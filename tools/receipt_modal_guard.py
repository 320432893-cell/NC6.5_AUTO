# 职责: 识别收款单保存前可取消的 NC Java 模态弹窗，并用 Alt+C 恢复页面
# 不做什么: 不写表头/明细，不保存/暂存，不枚举业务弹窗白名单
# 允许依赖层: 标准库 ctypes/time、tools.jab_probe、tools.receipt_keyboard_utils
# 谁不应该 import: core、配置校验、Excel/Sheet 写入模块不应 import

import ctypes
from ctypes import wintypes
import sys
import time

from tools.jab_probe import JOBJECT, enum_windows
from tools.receipt_keyboard_utils import send_hotkey_alt_c


def recover_cancelable_modal_now(jab, stage="", settle_timeout=0.25):
    dialogs = collect_visible_java_dialogs(jab)
    recoverable = [item for item in dialogs if item.get("cancel_controls")]
    if not recoverable:
        return {
            "ok": True,
            "attempted": False,
            "stage": stage,
            "dialog_count": len(dialogs),
            "reason": "未发现带取消按钮的 Java 弹窗",
        }
    event = {
        "ok": False,
        "attempted": True,
        "stage": stage,
        "method": "Alt+C",
        "dialogs": recoverable[:5],
    }
    event["focus"] = focus_window(recoverable[0].get("hwnd"))
    send_hotkey_alt_c()
    time.sleep(float(settle_timeout or 0))
    after = collect_visible_java_dialogs(jab)
    still_recoverable = [item for item in after if item.get("cancel_controls")]
    event["after_dialogs"] = after[:5]
    event["ok"] = not still_recoverable
    event["reason"] = None if event["ok"] else "Alt+C 后仍存在带取消按钮的 Java 弹窗"
    return event


def focus_window(hwnd):
    if sys.platform != "win32" or not hwnd:
        return {"ok": False, "reason": "必须在 Windows Python 下运行且需要 hwnd"}
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    hwnd = int(hwnd)
    foreground = int(user32.GetForegroundWindow() or 0)
    current_thread = int(kernel32.GetCurrentThreadId())
    target_thread = int(user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), None))
    foreground_thread = (
        int(user32.GetWindowThreadProcessId(wintypes.HWND(foreground), None))
        if foreground
        else 0
    )
    attached = []
    try:
        user32.ShowWindow(wintypes.HWND(hwnd), 9)
        user32.BringWindowToTop(wintypes.HWND(hwnd))
        for thread_id in {target_thread, foreground_thread}:
            if thread_id and thread_id != current_thread:
                if user32.AttachThreadInput(current_thread, thread_id, True):
                    attached.append(thread_id)
        ok = bool(user32.SetForegroundWindow(wintypes.HWND(hwnd)))
        user32.SetFocus(wintypes.HWND(hwnd))
        user32.SetActiveWindow(wintypes.HWND(hwnd))
        time.sleep(0.05)
        after = int(user32.GetForegroundWindow() or 0)
        return {
            "ok": bool(ok or after == hwnd),
            "hwnd": hwnd,
            "foreground_before": foreground,
            "foreground_after": after,
            "attached_threads": attached,
        }
    finally:
        for thread_id in attached:
            user32.AttachThreadInput(current_thread, thread_id, False)


def collect_visible_java_dialogs(jab):
    dialogs = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if not visible or class_name != "SunAwtDialog":
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue
        item = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": bool(visible),
            "root_hwnd": root_hwnd(hwnd),
        }
        item.update(scan_dialog_controls(jab, hwnd))
        dialogs.append(item)
    return dialogs


def scan_dialog_controls(jab, hwnd):
    vm_id_ref = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        hwnd,
        ctypes.byref(vm_id_ref),
        ctypes.byref(root_context),
    ):
        return {"error": "getAccessibleContextFromHWND failed"}
    buttons = []
    owned = [root_context.value]
    try:
        collect_buttons(jab, vm_id_ref.value, root_context.value, [], buttons, owned, 0)
    finally:
        jab.release_contexts(vm_id_ref.value, list(dict.fromkeys(owned)))
    cancel_controls = [
        item
        for item in buttons
        if "取消" in item.get("name", "") or "Alt+C" in item.get("description", "")
    ]
    return {"buttons": buttons[:20], "cancel_controls": cancel_controls}


def collect_buttons(jab, vm_id, context, path, buttons, owned, depth):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    role = (info.role_en_US.strip() or info.role.strip()).lower()
    states = (info.states_en_US.strip() or info.states.strip()).lower()
    if role == "push button" and "showing" in states:
        item = info_to_dict(info)
        item["path"] = ".".join(map(str, path))
        buttons.append(item)
    if depth >= min(jab.max_depth, 12):
        return
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        owned.append(child)
        collect_buttons(jab, vm_id, child, path + [index], buttons, owned, depth + 1)


def info_to_dict(info):
    states = info.states_en_US.strip() or info.states.strip()
    return {
        "name": info.name.strip(),
        "description": info.description.strip(),
        "role": info.role_en_US.strip() or info.role.strip(),
        "states": states,
        "showing": "showing" in states.lower(),
        "bounds": [info.x, info.y, info.width, info.height],
    }


def root_hwnd(hwnd):
    if sys.platform != "win32" or not hwnd:
        return 0
    return int(ctypes.windll.user32.GetAncestor(wintypes.HWND(int(hwnd)), 2) or 0)
