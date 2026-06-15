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


def recover_cancelable_modal_before_save(jab, page_probe, timeout=0.8):
    before_page = page_probe(jab)
    if before_page.get("ok"):
        return {
            "ok": True,
            "skipped": True,
            "reason": "当前仍是收款单录入页，无需 Alt+C 恢复",
            "before_page": before_page,
        }
    dialogs = collect_visible_java_dialogs(jab)
    recoverable = [item for item in dialogs if item.get("cancel_controls")]
    if not recoverable:
        return {
            "ok": False,
            "reason": "当前不是收款单录入页，且未找到带取消按钮的可恢复弹窗",
            "before_page": before_page,
            "dialogs": dialogs,
        }
    send_hotkey_alt_c()
    time.sleep(float(timeout or 0))
    after_page = page_probe(jab)
    after_dialogs = collect_visible_java_dialogs(jab)
    return {
        "ok": bool(after_page.get("ok")),
        "method": "Alt+C",
        "before_page": before_page,
        "before_dialogs": dialogs,
        "recoverable_dialog_count": len(recoverable),
        "after_page": after_page,
        "after_dialogs": after_dialogs,
        "reason": None if after_page.get("ok") else "Alt+C 后仍未恢复收款单录入页",
    }


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
