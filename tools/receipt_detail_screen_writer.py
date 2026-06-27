# 职责：通过 JAB 选中收款单明细单元格，并用前台守卫键盘输入写入
# 不做什么：不决定业务字段顺序，不增删明细行，不读取 Excel
# 允许依赖层：标准库 ctypes/sys/time、core.jab_probe、tools.receipt_keyboard_utils
# 谁不应该 import：配置校验、Sheet 写入、收款匹配模块不应 import

import ctypes
import sys
import time

from core.jab_probe import AccessibleTableCellInfo
from tools.receipt_keyboard_utils import (
    get_clipboard_text,
    read_window_info,
    restore_clipboard_text,
    send_hotkey_ctrl_a,
    send_text,
    send_unicode_char,
    send_virtual_key,
    set_clipboard_text,
)

KEYBOARD_INPUT_COMMIT_KEY = "Enter"
KEY_WAIT_SECONDS = 0.035
REFERENCE_ACCEPT_SETTLE_SECONDS = 0.04
VK_KEYS = {
    "F2": 0x71,
    "Right": 0x27,
    "Left": 0x25,
    "Enter": 0x0D,
    "Delete": 0x2E,
}
VK_CONTROL = 0x11
VK_V = 0x56


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


def move_selected_cell_by_arrows(table_window, from_col, to_col):
    start = int(from_col)
    target = int(to_col)
    delta = target - start
    if delta == 0:
        return {
            "ok": True,
            "skipped": True,
            "from_col": start,
            "to_col": target,
            "steps": [],
            "reason": "当前已在目标列",
        }
    key_name = "Right" if delta > 0 else "Left"
    steps = []
    for _index in range(abs(delta)):
        sent = guarded_press_virtual_key(table_window, key_name)
        steps.append(sent)
        if not sent.get("ok"):
            return {
                "ok": False,
                "from_col": start,
                "to_col": target,
                "steps": steps,
                "reason": sent.get("reason"),
            }
    return {
        "ok": True,
        "from_col": start,
        "to_col": target,
        "key": key_name,
        "count": abs(delta),
        "steps": steps,
    }


def keyboard_write_selected_cell(
    table_window,
    value,
    commit_key=KEYBOARD_INPUT_COMMIT_KEY,
    clear_only=False,
    accept_key=None,
    typing_interval=0.0,
    edit_mode="editor",
    input_mode="paste",
    pre_commit_wait=0.025,
    recover_after_failure=None,
    _recovery_retry=False,
):
    timing = {
        "guard_seconds": 0.0,
        "edit_prepare_seconds": 0.0,
        "clear_seconds": 0.0,
        "clipboard_read_seconds": 0.0,
        "clipboard_set_seconds": 0.0,
        "paste_send_seconds": 0.0,
        "type_send_seconds": 0.0,
        "pre_commit_wait_seconds": 0.0,
        "accept_seconds": 0.0,
        "accept_settle_wait_seconds": 0.0,
        "commit_seconds": 0.0,
        "clipboard_restore_seconds": 0.0,
    }
    guard_started = time.perf_counter()
    guard = foreground_matches_table(table_window)
    if not guard.get("ok"):
        recovery = recover_after_failure() if recover_after_failure else None
        if recovery and recovery.get("attempted") and recovery.get("ok"):
            guard = foreground_matches_table(table_window)
        timing["guard_seconds"] = round(time.perf_counter() - guard_started, 4)
        if not guard.get("ok"):
            return {**guard, "modal_recovery": recovery, "screen_timing": timing}
    else:
        timing["guard_seconds"] = round(time.perf_counter() - guard_started, 4)

    def retry_current_cell_after_failure(failure):
        if _recovery_retry or recover_after_failure is None:
            return failure
        recovery = recover_after_failure()
        if not (recovery.get("attempted") and recovery.get("ok")):
            return {**failure, "retry_modal_recovery": recovery}
        retried = keyboard_write_selected_cell(
            table_window,
            value,
            commit_key=commit_key,
            clear_only=clear_only,
            accept_key=accept_key,
            typing_interval=typing_interval,
            edit_mode=edit_mode,
            input_mode=input_mode,
            pre_commit_wait=pre_commit_wait,
            recover_after_failure=recover_after_failure,
            _recovery_retry=True,
        )
        return {
            **retried,
            "retried_after_modal_recovery": True,
            "retry_modal_recovery": recovery,
            "first_failure": failure,
        }

    old_clipboard = None
    clipboard_restored = None
    retry_recovery = None
    try:
        if edit_mode != "selected":
            edit_started = time.perf_counter()
            send_virtual_key(VK_KEYS["F2"])
            time.sleep(0.025)
            send_hotkey_ctrl_a()
            time.sleep(0.02)
            timing["edit_prepare_seconds"] = round(
                time.perf_counter() - edit_started, 4
            )
        clear = None
        if clear_only:
            clear_started = time.perf_counter()
            clear = guarded_press_virtual_key(table_window, "Delete")
            timing["clear_seconds"] = round(time.perf_counter() - clear_started, 4)
            if not clear.get("ok"):
                return retry_current_cell_after_failure(
                    {
                        **guard,
                        "ok": False,
                        "mode": "keyboard",
                        "clear_only": clear_only,
                        "clear": clear,
                        "accept_key": accept_key,
                        "commit_key": commit_key,
                        "screen_timing": timing,
                        "reason": clear.get("reason"),
                    }
                )
        else:
            if input_mode == "paste":
                clipboard_read_started = time.perf_counter()
                old_clipboard = safe_clipboard_read()
                timing["clipboard_read_seconds"] = round(
                    time.perf_counter() - clipboard_read_started, 4
                )
                try:
                    clipboard_set_started = time.perf_counter()
                    set_clipboard_text(str(value))
                    timing["clipboard_set_seconds"] = round(
                        time.perf_counter() - clipboard_set_started, 4
                    )
                except RuntimeError as exc:
                    if (
                        str(exc) != "OpenClipboard failed"
                        or recover_after_failure is None
                    ):
                        raise
                    retry_recovery = recover_after_failure()
                    if not retry_recovery.get("ok"):
                        raise
                    time.sleep(0.05)
                    clipboard_set_started = time.perf_counter()
                    set_clipboard_text(str(value))
                    timing["clipboard_set_seconds"] = round(
                        time.perf_counter() - clipboard_set_started, 4
                    )
                paste_started = time.perf_counter()
                send_hotkey_ctrl_v()
                timing["paste_send_seconds"] = round(
                    time.perf_counter() - paste_started, 4
                )
            else:
                type_started = time.perf_counter()
                send_text_slow(value, typing_interval)
                timing["type_send_seconds"] = round(
                    time.perf_counter() - type_started, 4
                )
        pre_commit_wait_started = time.perf_counter()
        time.sleep(float(pre_commit_wait or 0))
        timing["pre_commit_wait_seconds"] = round(
            time.perf_counter() - pre_commit_wait_started, 4
        )
        accept = None
        if accept_key:
            accept_started = time.perf_counter()
            accept = guarded_press_virtual_key(table_window, accept_key)
            timing["accept_seconds"] = round(time.perf_counter() - accept_started, 4)
            if not accept.get("ok"):
                return retry_current_cell_after_failure(
                    {
                        **guard,
                        "ok": False,
                        "mode": "keyboard",
                        "clear_only": clear_only,
                        "accept_key": accept_key,
                        "accept": accept,
                        "commit_key": commit_key,
                        "screen_timing": timing,
                        "reason": accept.get("reason"),
                    }
                )
            accept_wait_started = time.perf_counter()
            time.sleep(REFERENCE_ACCEPT_SETTLE_SECONDS)
            timing["accept_settle_wait_seconds"] = round(
                time.perf_counter() - accept_wait_started, 4
            )
        commit_started = time.perf_counter()
        commit = guarded_press_virtual_key(table_window, commit_key)
        timing["commit_seconds"] = round(time.perf_counter() - commit_started, 4)
    except Exception as exc:
        return retry_current_cell_after_failure(
            {
                **guard,
                "ok": False,
                "reason": f"键盘输入失败：{type(exc).__name__}: {exc}",
                "accept_key": accept_key,
                "commit_key": commit_key,
                "screen_timing": timing,
                "retry_modal_recovery": retry_recovery,
            }
        )
    finally:
        if old_clipboard is not None:
            clipboard_restore_started = time.perf_counter()
            clipboard_restored = restore_clipboard_text(old_clipboard)
            timing["clipboard_restore_seconds"] = round(
                time.perf_counter() - clipboard_restore_started, 4
            )
    result = {
        **guard,
        "ok": bool(commit.get("ok")),
        "mode": "keyboard",
        "clear_only": clear_only,
        "clear": clear,
        "accept_key": accept_key,
        "typing_interval": float(typing_interval or 0),
        "edit_mode": edit_mode,
        "input_mode": input_mode,
        "pre_commit_wait": float(pre_commit_wait or 0),
        "retry_modal_recovery": retry_recovery,
        "clipboard_restored": clipboard_restored,
        "accept": accept,
        "commit_key": commit_key,
        "commit": commit,
        "screen_timing": timing,
        "reason": None if commit.get("ok") else commit.get("reason"),
    }
    if not result["ok"]:
        return retry_current_cell_after_failure(result)
    return result


def send_text_slow(value, interval=0.0):
    text = str(value)
    delay = float(interval or 0)
    if delay <= 0:
        send_text(text)
        return
    for char in text:
        send_unicode_char(char)
        time.sleep(delay)


def safe_clipboard_read():
    try:
        return get_clipboard_text()
    except Exception:
        return None


def send_hotkey_ctrl_v():
    send_virtual_key(VK_CONTROL, key_up=False)
    send_virtual_key(VK_V, key_up=False)
    send_virtual_key(VK_V, key_up=True)
    send_virtual_key(VK_CONTROL, key_up=True)
