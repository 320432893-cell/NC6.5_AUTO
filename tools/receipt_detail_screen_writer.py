# 职责：通过 JAB 选中收款单明细单元格，并用前台守卫键盘输入写入
# 不做什么：不决定业务字段顺序，不增删明细行，不读取 Excel
# 允许依赖层：标准库 ctypes/sys/time、tools.jab_probe、tools.tmp_receipt_cell_probe_run 的 SendInput 兼容函数
# 谁不应该 import：配置校验、Sheet 写入、收款匹配模块不应 import

import ctypes
import sys
import time

from tools.jab_probe import AccessibleTableCellInfo
from tools.tmp_receipt_cell_probe_run import (
    read_window_info,
    send_hotkey_ctrl_a,
    send_text,
    send_virtual_key,
)

KEYBOARD_INPUT_COMMIT_KEY = "Right"
KEY_WAIT_SECONDS = 0.035
REFERENCE_ACCEPT_SETTLE_SECONDS = 0.04
VK_KEYS = {
    "F2": 0x71,
    "Right": 0x27,
    "Left": 0x25,
    "Enter": 0x0D,
    "Delete": 0x2E,
}


def activate_window(hwnd):
    if sys.platform != "win32" or not hwnd:
        return {"ok": False, "reason": "必须在 Windows Python 下运行且需要 hwnd"}
    user32 = ctypes.windll.user32
    hwnd = int(hwnd)
    root = user32.GetAncestor(hwnd, 2) or hwnd
    user32.ShowWindow(root, 9)
    user32.BringWindowToTop(root)
    ok = bool(user32.SetForegroundWindow(root))
    time.sleep(0.05)
    return {
        "ok": ok,
        "hwnd": hwnd,
        "root_hwnd": int(root),
        "foreground": read_window_info(user32.GetForegroundWindow()),
    }


def get_cell_context(jab, vm_id, table_context, row, col):
    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    if not ok or not cell_info.accessibleContext:
        return None
    return cell_info.accessibleContext


def get_cell_info(jab, vm_id, table_context, row, col):
    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    return cell_info if ok else None


def focus_detail_cell(jab, located, row_index, col_index):
    best = located.get("best") or {}
    context, vm_id, owned, window_info = jab.find_context_by_path_once(
        best.get("path"),
        class_name=(best.get("window") or {}).get("class_name"),
        scope_hwnd=(best.get("window") or {}).get("hwnd"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {"ok": False, "reason": "按 path 重新取得明细表 context 失败"}

    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {"ok": False, "reason": "getAccessibleTableInfo 失败"}
        if row_index < 0 or row_index >= table_info.rowCount:
            return {
                "ok": False,
                "reason": f"目标行越界：{row_index} / {table_info.rowCount}",
            }
        if col_index < 0 or col_index >= table_info.columnCount:
            return {
                "ok": False,
                "reason": f"目标列越界：{col_index} / {table_info.columnCount}",
            }
        foreground = foreground_matches_table(window_info)
        activate = (
            {
                "ok": True,
                "skipped": True,
                "reason": "NC 已在前台，未重复激活窗口",
                "foreground": foreground,
            }
            if foreground.get("ok")
            else activate_window(window_info.get("hwnd"))
        )
        child_index = row_index * table_info.columnCount + col_index
        if not jab.has_selection_api():
            return {"ok": False, "reason": "JAB selection API 不可用"}
        jab.dll.clearAccessibleSelectionFromContext(vm_id, context)
        jab.dll.addAccessibleSelectionFromContext(vm_id, context, child_index)
        time.sleep(0.05)
        cell_info = get_cell_info(jab, vm_id, context, row_index, col_index)
        selected = bool(cell_info and cell_info.isSelected)
        cell_context = cell_info.accessibleContext if cell_info else None
        focus_results = []
        if hasattr(jab.dll, "requestFocus"):
            focus_results.append(
                {"target": "table", "ok": bool(jab.dll.requestFocus(vm_id, context))}
            )
            time.sleep(0.02)
            if cell_context:
                focus_results.append(
                    {
                        "target": "cell",
                        "ok": bool(jab.dll.requestFocus(vm_id, cell_context)),
                    }
                )
                time.sleep(0.02)
        return {
            "ok": bool(activate.get("ok")) and selected,
            "activate": activate,
            "child_index": child_index,
            "selected": selected,
            "request_focus": focus_results,
            "window": window_info,
            "target": {"row": row_index, "col": col_index},
            "reason": None
            if bool(activate.get("ok")) and selected
            else activate.get("reason") or "JAB 选中目标单元格后读回不匹配",
        }
    finally:
        jab.release_contexts(vm_id, owned)


def foreground_matches_table(table_window):
    if sys.platform != "win32":
        return {"ok": False, "reason": "必须在 Windows Python 下运行"}
    table_info = read_window_info((table_window or {}).get("hwnd"))
    foreground_info = read_window_info(ctypes.windll.user32.GetForegroundWindow())
    if not table_info or not foreground_info:
        return {
            "ok": False,
            "reason": "无法读取当前前台窗口或明细表窗口",
            "table_window": table_info,
            "foreground": foreground_info,
        }
    same_root = (
        foreground_info.get("hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("root_hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("hwnd") == table_info.get("hwnd")
    )
    return {
        "ok": bool(same_root),
        "reason": None if same_root else "当前前台窗口不是本次定位到的 NC 收款单窗口",
        "table_window": table_info,
        "foreground": foreground_info,
    }


def guarded_press_virtual_key(table_window, key_name):
    if key_name not in VK_KEYS:
        return {"ok": False, "reason": f"未知按键：{key_name}"}
    guard = foreground_matches_table(table_window)
    if not guard.get("ok"):
        return {**guard, "key": key_name, "ok": False}
    try:
        send_virtual_key(VK_KEYS[key_name], key_up=False)
        time.sleep(0.015)
        send_virtual_key(VK_KEYS[key_name], key_up=True)
    except Exception as exc:
        return {
            **guard,
            "ok": False,
            "key": key_name,
            "mode": f"SendInput({key_name})",
            "reason": f"{type(exc).__name__}: {exc}",
        }
    time.sleep(KEY_WAIT_SECONDS)
    return {
        **guard,
        "ok": True,
        "key": key_name,
        "mode": f"SendInput({key_name})",
    }


def keyboard_write_selected_cell(
    table_window,
    value,
    commit_key=KEYBOARD_INPUT_COMMIT_KEY,
    clear_only=False,
    accept_key=None,
):
    guard = foreground_matches_table(table_window)
    if not guard.get("ok"):
        return guard
    try:
        send_virtual_key(VK_KEYS["F2"])
        time.sleep(0.025)
        send_hotkey_ctrl_a()
        time.sleep(0.02)
        clear = None
        if clear_only:
            clear = guarded_press_virtual_key(table_window, "Delete")
            if not clear.get("ok"):
                return {
                    **guard,
                    "ok": False,
                    "mode": "keyboard",
                    "clear_only": clear_only,
                    "clear": clear,
                    "accept_key": accept_key,
                    "commit_key": commit_key,
                    "reason": clear.get("reason"),
                }
        else:
            send_text(value)
        time.sleep(0.025)
        accept = None
        if accept_key:
            accept = guarded_press_virtual_key(table_window, accept_key)
            if not accept.get("ok"):
                return {
                    **guard,
                    "ok": False,
                    "mode": "keyboard",
                    "clear_only": clear_only,
                    "accept_key": accept_key,
                    "accept": accept,
                    "commit_key": commit_key,
                    "reason": accept.get("reason"),
                }
            time.sleep(REFERENCE_ACCEPT_SETTLE_SECONDS)
        commit = guarded_press_virtual_key(table_window, commit_key)
    except Exception as exc:
        return {
            **guard,
            "ok": False,
            "reason": f"键盘输入失败：{type(exc).__name__}: {exc}",
            "accept_key": accept_key,
            "commit_key": commit_key,
        }
    return {
        **guard,
        "ok": bool(commit.get("ok")),
        "mode": "keyboard",
        "clear_only": clear_only,
        "clear": clear,
        "accept_key": accept_key,
        "accept": accept,
        "commit_key": commit_key,
        "commit": commit,
        "reason": None if commit.get("ok") else commit.get("reason"),
    }


def read_selected_cell(jab, located):
    best = located.get("best") or {}
    context, vm_id, owned, _window_info = jab.find_context_by_path_once(
        best.get("path"),
        class_name=(best.get("window") or {}).get("class_name"),
        scope_hwnd=(best.get("window") or {}).get("hwnd"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {"ok": False, "reason": "按 path 重新取得明细表 context 失败"}
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {"ok": False, "reason": "getAccessibleTableInfo 失败"}
        selected = []
        for row in range(table_info.rowCount):
            for col in range(table_info.columnCount):
                cell_info = AccessibleTableCellInfo()
                ok = jab.dll.getAccessibleTableCellInfo(
                    vm_id,
                    context,
                    row,
                    col,
                    ctypes.byref(cell_info),
                )
                if not ok or not cell_info.isSelected:
                    continue
                text = ""
                if cell_info.accessibleContext:
                    info = jab.get_context_info(vm_id, cell_info.accessibleContext)
                    if info:
                        text = info.name.strip() or info.description.strip()
                selected.append({"row": row, "col": col, "text": text})
        return {
            "ok": True,
            "selected": selected,
            "single": selected[0] if len(selected) == 1 else None,
        }
    finally:
        jab.release_contexts(vm_id, owned)
