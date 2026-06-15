# 职责: 归档历史表头收款银行账户参照搜索/带回探针
# 不做什么: 不作为正式收款单录入逻辑，不提供正式账号录入 fallback，不写 Excel/保存 NC
# 允许依赖层: core JAB 操作、历史现场探测工具
# 谁不应该 import: core、正式 tools 入口、tests、配置校验和 Excel/Sheet 写入模块不应 import

import argparse
import ctypes
import json
import os
import sys
import time
from ctypes import wintypes
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402

RESULT_TABLE_PATH = "0.0.1.0.0.0.1.3.0.2.0"
OK_PATH = "0.0.1.0.0.0.1.2.1.0"
STOP_HOTKEY = "Space"
VK_SPACE = 0x20
VK_A = 0x41
VK_F = 0x46
VK_RETURN = 0x0D
VK_MENU = 0x12
VK_CONTROL = 0x11
KEYEVENTF_KEYUP = 0x0002
GMEM_MOVEABLE = 0x0002
CF_UNICODETEXT = 13
_CLIPBOARD_API_CONFIGURED = False


def configure_clipboard_api():
    global _CLIPBOARD_API_CONFIGURED
    if _CLIPBOARD_API_CONFIGURED:
        return
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    user32.OpenClipboard.argtypes = [ctypes.c_void_p]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL
    user32.EmptyClipboard.argtypes = []
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE
    user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user32.SetClipboardData.restype = wintypes.HANDLE
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    _CLIPBOARD_API_CONFIGURED = True


def main():
    parser = argparse.ArgumentParser(
        description="Stage account search actions in 使用权参照."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument(
        "--stage",
        choices=(
            "find-dialog",
            "check-foreground",
            "type-search",
            "search",
            "read-table",
            "select-first",
            "read-select",
            "select-confirm",
            "full",
        ),
        default="find-dialog",
    )
    parser.add_argument("--hwnd", type=int)
    parser.add_argument("--account")
    parser.add_argument("--press-enter", action="store_true")
    parser.add_argument("--check-timeout", type=float, default=1.0)
    parser.add_argument("--poll-timeout", type=float, default=2.0)
    parser.add_argument("--stop-hotkey", default=STOP_HOTKEY)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    report = {"stage": args.stage, "hwnd": args.hwnd, "account": args.account}
    try:
        jab.ensure_started()
        if args.stage == "find-dialog":
            report["dialogs"] = collect_reference_dialogs(jab)
            report["foreground"] = get_foreground_window_info()
            report["ok"] = len(report["dialogs"]) == 1
        elif args.stage == "full":
            report.update(
                run_full_account_reference(
                    jab,
                    hwnd=args.hwnd,
                    account=args.account,
                    press_enter=args.press_enter,
                    check_timeout=args.check_timeout,
                    poll_timeout=args.poll_timeout,
                    stop_hotkey=args.stop_hotkey,
                )
            )
        else:
            hwnd_result = resolve_reference_hwnd(jab, args.hwnd)
            report["target"] = hwnd_result
            if not hwnd_result.get("ok"):
                report["ok"] = False
            else:
                hwnd = hwnd_result["hwnd"]
                if args.stage == "check-foreground":
                    report["foreground_check"] = wait_foreground_match(
                        hwnd, timeout=args.check_timeout
                    )
                    report["ok"] = bool(report["foreground_check"].get("ok"))
                elif args.stage in ("type-search", "search"):
                    if not args.account:
                        report["ok"] = False
                        report["reason"] = f"--account is required for {args.stage}"
                    else:
                        if args.stage == "search":
                            report["foreground_check"] = wait_foreground_match(
                                hwnd, timeout=args.check_timeout
                            )
                        report["search"] = focus_search_and_type(
                            jab, hwnd, args.account, args.press_enter
                        )
                        report["ok"] = bool(report["search"].get("ok"))
                elif args.stage == "read-table":
                    report["table"] = wait_table(jab, hwnd, args.poll_timeout)
                    report["ok"] = bool(report["table"].get("ok"))
                elif args.stage == "select-first":
                    report["selection"] = select_reference_result_first_row(jab, hwnd)
                    report["ok"] = bool(report["selection"].get("ok"))
                elif args.stage == "read-select":
                    report["table"] = wait_table(jab, hwnd, args.poll_timeout)
                    if (
                        report["table"].get("ok")
                        and report["table"].get("row_count", 0) > 0
                    ):
                        report["selection"] = select_reference_result_first_row(
                            jab, hwnd
                        )
                    report["ok"] = bool(
                        report["table"].get("ok")
                        and report["table"].get("row_count", 0) > 0
                        and report.get("selection", {}).get("ok")
                    )
                elif args.stage == "select-confirm":
                    selection = select_reference_result_first_row(jab, hwnd)
                    report["selection"] = selection
                    if selection.get("ok"):
                        report["confirm"] = confirm_reference_selection(jab, hwnd)
                    report["ok"] = bool(
                        selection.get("ok") and report.get("confirm", {}).get("ok")
                    )
    finally:
        jab.close()

    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


def run_full_account_reference(
    jab,
    hwnd=None,
    account=None,
    press_enter=True,
    check_timeout=1.0,
    poll_timeout=3.0,
    stop_hotkey=STOP_HOTKEY,
    confirm_selection=True,
):
    steps: list[dict[str, object]] = []
    result: dict[str, object] = {"steps": steps, "stop_hotkey": stop_hotkey}
    if not account:
        return {"ok": False, "reason": "--account is required for full"}

    target = resolve_reference_hwnd(jab, hwnd)
    result["target"] = target
    steps.append({"step": "find-dialog", **target})
    if not target.get("ok"):
        result["ok"] = False
        result["failed_step"] = "find-dialog"
        return result

    hwnd = target["hwnd"]
    if is_stop_hotkey_pressed():
        return stopped_result(result, "before-check-foreground")

    foreground = wait_foreground_match(hwnd, timeout=check_timeout)
    result["foreground_check"] = foreground
    steps.append({"step": "check-foreground", **foreground})
    if not foreground.get("ok"):
        result["ok"] = False
        result["failed_step"] = "check-foreground"
        result["reason"] = "target dialog is not foreground; no keyboard input sent"
        return result

    if is_stop_hotkey_pressed():
        return stopped_result(result, "before-type-search")

    search = focus_search_and_type(jab, hwnd, account, press_enter=press_enter)
    result["search"] = search
    steps.append({"step": "type-search", **search})
    if not search.get("ok"):
        result["ok"] = False
        result["failed_step"] = "type-search"
        return result

    if is_stop_hotkey_pressed():
        return stopped_result(result, "before-read-table")

    table = wait_table(jab, hwnd, poll_timeout)
    result["table"] = table
    if table.get("stopped_by_hotkey"):
        steps.append({"step": "read-table", **table})
        return stopped_result(result, "read-table")
    steps.append(
        {
            "step": "read-table",
            "ok": bool(table.get("ok") and table.get("row_count", 0) > 0),
            "row_count": table.get("row_count"),
            "col_count": table.get("col_count"),
            "stable": table.get("stable"),
            "reason": table.get("reason"),
        }
    )
    if not table.get("ok") or table.get("row_count", 0) <= 0:
        result["ok"] = False
        result["failed_step"] = "read-table"
        return result

    if is_stop_hotkey_pressed():
        return stopped_result(result, "before-select-first")

    selection = select_reference_result_first_row(jab, hwnd)
    result["selection"] = selection
    steps.append({"step": "select-first", **selection})
    if not selection.get("ok"):
        result["ok"] = False
        result["failed_step"] = "select-first"
        return result

    if not confirm_selection:
        result["ok"] = True
        result["confirmed"] = False
        result["reason"] = "selected first row; confirm skipped by test launcher"
        return result

    if is_stop_hotkey_pressed():
        return stopped_result(result, "before-confirm")

    confirm = confirm_reference_selection(jab, hwnd)
    result["confirm"] = confirm
    steps.append({"step": "confirm", **confirm})
    result["ok"] = bool(confirm.get("ok"))
    if not result["ok"]:
        result["failed_step"] = "confirm"
    return result


def stopped_result(result, step):
    result["ok"] = False
    result["stopped_by_hotkey"] = True
    result["failed_step"] = step
    result["reason"] = f"emergency stop hotkey pressed: {STOP_HOTKEY}"
    return result


def is_stop_hotkey_pressed():
    if os.name != "nt":
        return False
    user32 = ctypes.windll.user32
    return all(bool(user32.GetAsyncKeyState(vk) & 0x8000) for vk in (VK_SPACE,))


def focus_search_and_type(jab, hwnd, value, press_enter=False):
    foreground = check_foreground_match(hwnd)
    if not foreground.get("ok"):
        return {
            "ok": False,
            "reason": "target dialog is not foreground; refused global keyboard input",
            "target_hwnd": int(hwnd),
            "foreground": foreground,
        }
    send_hotkey(VK_MENU, VK_F)
    time.sleep(0.5)
    send_hotkey(VK_CONTROL, VK_A)
    time.sleep(0.1)
    clipboard = paste_text_via_clipboard(value)
    if press_enter:
        if is_stop_hotkey_pressed():
            return {
                "ok": False,
                "stopped_by_hotkey": True,
                "reason": f"emergency stop hotkey pressed: {STOP_HOTKEY}",
            }
        send_key(VK_RETURN)
    return {
        "ok": True,
        "foreground": foreground,
        "method": "winapi_alt_f_clipboard_paste",
        "clipboard": clipboard,
        "press_enter": bool(press_enter),
    }


def send_key(vk):
    user32 = ctypes.windll.user32
    user32.keybd_event(vk, 0, 0, 0)
    user32.keybd_event(vk, 0, 2, 0)


def send_hotkey(modifier_vk, key_vk):
    user32 = ctypes.windll.user32
    user32.keybd_event(modifier_vk, 0, 0, 0)
    user32.keybd_event(key_vk, 0, 0, 0)
    user32.keybd_event(key_vk, 0, 2, 0)
    user32.keybd_event(modifier_vk, 0, 2, 0)


def paste_text_via_clipboard(text):
    before = get_clipboard_text()
    set_clipboard_text(str(text))
    time.sleep(0.1)
    send_hotkey(VK_CONTROL, 0x56)
    time.sleep(0.2)
    restored = restore_clipboard_text(before)
    return {
        "ok": True,
        "method": "clipboard_ctrl_v",
        "previous_text_preserved": restored,
    }


def get_clipboard_text():
    configure_clipboard_api()
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    if not user32.OpenClipboard(None):
        return None
    try:
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return None
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return None
        try:
            return ctypes.wstring_at(ptr)
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


def set_clipboard_text(text):
    configure_clipboard_api()
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    encoded = str(text) + "\0"
    size = len(encoded) * ctypes.sizeof(ctypes.c_wchar)
    if not user32.OpenClipboard(None):
        raise RuntimeError("OpenClipboard failed")
    try:
        if not user32.EmptyClipboard():
            raise RuntimeError("EmptyClipboard failed")
        handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, size)
        if not handle:
            raise RuntimeError("GlobalAlloc failed")
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            raise RuntimeError("GlobalLock failed")
        try:
            ctypes.memmove(ptr, ctypes.create_unicode_buffer(encoded), size)
        finally:
            kernel32.GlobalUnlock(handle)
        if not user32.SetClipboardData(CF_UNICODETEXT, handle):
            raise RuntimeError("SetClipboardData failed")
    finally:
        user32.CloseClipboard()


def restore_clipboard_text(text):
    if text is None:
        return False
    try:
        set_clipboard_text(text)
    except RuntimeError:
        return False
    return True


def collect_reference_dialogs(jab):
    dialogs = []
    for hwnd, title, class_name, pid, visible in jab.get_scoped_windows(
        None, include_children=True
    ):
        if (
            visible
            and title == "使用权参照"
            and class_name == "SunAwtDialog"
            and jab.dll.isJavaWindow(hwnd)
        ):
            dialogs.append(
                {
                    "hwnd": int(hwnd),
                    "title": title,
                    "class_name": class_name,
                    "pid": pid,
                    "visible": visible,
                }
            )
    return dialogs


def resolve_reference_hwnd(jab, hwnd=None):
    if hwnd:
        return {"ok": True, "hwnd": int(hwnd), "method": "argument"}
    dialogs = collect_reference_dialogs(jab)
    if len(dialogs) == 1:
        return {"ok": True, "hwnd": dialogs[0]["hwnd"], "method": "single_dialog"}
    return {
        "ok": False,
        "reason": "expected exactly one visible 使用权参照 dialog",
        "dialogs": dialogs,
    }


def get_foreground_window_info():
    if os.name != "nt":
        return None

    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None

    length = user32.GetWindowTextLengthW(hwnd)
    window_title = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, window_title, length + 1)

    window_class = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, window_class, 256)
    return {
        "hwnd": int(hwnd),
        "title": window_title.value,
        "class_name": window_class.value,
    }


def check_foreground_match(hwnd):
    if os.name != "nt":
        return {"ok": False, "reason": "must run under Windows Python"}
    foreground = get_foreground_window_info()
    return {
        "ok": bool(foreground and foreground.get("hwnd") == int(hwnd)),
        "target_hwnd": int(hwnd),
        "foreground": foreground,
    }


def wait_foreground_match(hwnd, timeout=1.0):
    timeout = min(max(float(timeout), 0.0), 1.0)
    deadline = time.time() + timeout
    last = check_foreground_match(hwnd)
    while not last.get("ok") and time.time() < deadline:
        if is_stop_hotkey_pressed():
            return {
                "ok": False,
                "stopped_by_hotkey": True,
                "reason": f"emergency stop hotkey pressed: {STOP_HOTKEY}",
                "target_hwnd": int(hwnd),
            }
        time.sleep(0.1)
        last = check_foreground_match(hwnd)
    last["timeout"] = timeout
    return last


def read_table(jab, hwnd):
    context, vm_id, owned, window = jab.find_context_by_path_once(
        RESULT_TABLE_PATH,
        scope_hwnd=hwnd,
        role="table",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "reason": "result table not found",
            "path": RESULT_TABLE_PATH,
        }
    try:
        info = jab.get_table_info(vm_id, context)
        if not info:
            return {"ok": False, "reason": "table info unavailable", "window": window}
        rows = []
        for row in range(min(info.rowCount, 5)):
            cells = {}
            for col in range(min(info.columnCount, 10)):
                cells[str(col)] = jab.get_table_cell_text(vm_id, context, row, col)
            rows.append({"row": row, "cells": cells})
        return {
            "ok": True,
            "path": RESULT_TABLE_PATH,
            "window": window,
            "row_count": info.rowCount,
            "col_count": info.columnCount,
            "rows": rows,
        }
    finally:
        jab.release_contexts(vm_id, owned)


def select_reference_result_first_row(jab, hwnd):
    context, vm_id, owned, window = jab.find_context_by_path_once(
        RESULT_TABLE_PATH,
        scope_hwnd=hwnd,
        role="table",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "reason": "result table not found",
            "path": RESULT_TABLE_PATH,
        }
    try:
        info = jab.get_table_info(vm_id, context)
        if not info:
            return {"ok": False, "reason": "table info unavailable", "window": window}
        if info.rowCount <= 0:
            return {
                "ok": False,
                "reason": "no rows",
                "row_count": info.rowCount,
                "col_count": info.columnCount,
            }
        before = jab.get_selected_child_indexes(
            vm_id, context, info.rowCount * info.columnCount
        )
        jab.dll.clearAccessibleSelectionFromContext(vm_id, context)
        jab.dll.addAccessibleSelectionFromContext(vm_id, context, 0)
        time.sleep(0.3)
        after = jab.get_selected_child_indexes(
            vm_id, context, info.rowCount * info.columnCount
        )
        cells = {
            str(col): jab.get_table_cell_text(vm_id, context, 0, col)
            for col in range(min(info.columnCount, 10))
        }
        return {
            "ok": 0 in after,
            "path": RESULT_TABLE_PATH,
            "window": window,
            "row_count": info.rowCount,
            "col_count": info.columnCount,
            "selected_before": before,
            "selected_after": after,
            "first_row_cells": cells,
        }
    finally:
        jab.release_contexts(vm_id, owned)


def confirm_reference_selection(jab, hwnd):
    ok = jab.do_action_by_path(
        OK_PATH,
        scope_hwnd=hwnd,
        action_name="单击",
        wait=1.2,
        timeout=3.0,
        require_showing=True,
        require_valid_bounds=False,
    )
    return {"ok": bool(ok), "path": OK_PATH, "method": "jab_action"}


def wait_table(jab, hwnd, timeout):
    deadline = time.time() + timeout
    last = None
    stable_count = 0
    while time.time() < deadline:
        if is_stop_hotkey_pressed():
            return {
                "ok": False,
                "stopped_by_hotkey": True,
                "reason": f"emergency stop hotkey pressed: {STOP_HOTKEY}",
            }
        current = read_table(jab, hwnd)
        if current.get("ok"):
            signature = (
                current.get("row_count"),
                current.get("col_count"),
                json.dumps(current.get("rows", []), ensure_ascii=False, sort_keys=True),
            )
            if current.get("row_count", 0) > 0 and signature == last:
                stable_count += 1
                if stable_count >= 2:
                    current["stable"] = True
                    return current
            else:
                stable_count = 0
            last = signature
        time.sleep(0.8)
    current = read_table(jab, hwnd)
    current["stable"] = False
    current["waited_timeout"] = timeout
    return current


if __name__ == "__main__":
    raise SystemExit(main())
