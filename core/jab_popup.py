# 职责：跟踪和清理 NC/JAB 下拉菜单类 SunAwtWindow popup。
# 不做什么：不实现收款单/凭证业务匹配，不全局强清理所有小窗。
# 允许依赖层：JABOperator-like 对象、core.jab_probe 窗口枚举/JOBJECT、Win32 ctypes。
# 谁不应该 import：Excel/Sheet 读写和纯数据解析模块不应 import。

from __future__ import annotations

import ctypes
import os
import time
from ctypes import wintypes
from typing import Any

from core.jab_probe import JOBJECT, enum_windows


def collect_visible_popup_windows(jab, max_depth=8, max_children=120) -> list[dict]:
    if os.name != "nt" or not hasattr(ctypes, "WinDLL"):
        return []

    windows = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if class_name != "SunAwtWindow" or not visible:
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            continue

        window = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
            "is_java": True,
            "root": summarize_context(jab, vm_id.value, root_context.value, "0"),
            "controls": [],
            "all_controls": [],
        }
        if is_visible_sun_awt_popup(window):
            collect_controls(
                jab,
                vm_id.value,
                root_context.value,
                path="0",
                controls=window["controls"],
                all_controls=window["all_controls"],
                depth=0,
                max_depth=max_depth,
                max_children=max_children,
            )
        windows.append(window)
    return windows


def wait_for_new_popup_with_menu_item(
    jab,
    before_windows: list[dict],
    menu_name: str,
    timeout=0.8,
    interval=0.08,
) -> dict:
    deadline = time.monotonic() + timeout
    last_windows = []
    while time.monotonic() < deadline:
        last_windows = collect_new_visible_popup_windows(jab, before_windows)
        popup = find_new_popup_with_menu_item(before_windows, last_windows, menu_name)
        if popup:
            return {"ok": True, "popup": popup, "windows": last_windows}
        time.sleep(interval)

    full_windows = collect_new_visible_popup_windows(jab, before_windows)
    popup = find_new_popup_with_menu_item(before_windows, full_windows, menu_name)
    return {
        "ok": bool(popup),
        "popup": popup,
        "windows": full_windows if popup else last_windows,
        "reason": None if popup else f"未检测到菜单项 {menu_name!r} 的 popup",
    }


def collect_new_visible_popup_windows(jab, before_windows: list[dict]) -> list[dict]:
    before_signatures = {
        window_key(item): window_signature(item) for item in before_windows
    }
    windows = []
    for window in collect_visible_popup_windows(jab):
        key = window_key(window)
        if key in before_signatures and before_signatures[key] == window_signature(
            window
        ):
            continue
        windows.append(window)
    return windows


def find_new_popup_with_menu_item(
    before_windows: list[dict],
    after_windows: list[dict],
    menu_name: str,
) -> dict | None:
    before_signatures = {
        window_key(item): window_signature(item) for item in before_windows
    }
    candidates = []
    for item in after_windows:
        key = window_key(item)
        if key in before_signatures and before_signatures[key] == window_signature(
            item
        ):
            continue
        if not is_visible_sun_awt_popup(item):
            continue
        menu_names = {
            control.get("name")
            for control in item.get("all_controls", [])
            if str(control.get("role", "")).lower() == "menu item"
        }
        if menu_name in menu_names:
            candidates.append(item)
    candidates.sort(key=lambda item: (item.get("root") or {}).get("bounds", [0, 0])[1])
    if not candidates:
        return None
    item = candidates[0]
    return {
        "hwnd": item.get("hwnd"),
        "class_name": item.get("class_name"),
        "title": item.get("title"),
        "visible": item.get("visible"),
        "root": item.get("root"),
        "menu_items": [
            summarize_control(control)
            for control in item.get("all_controls", [])
            if str(control.get("role", "")).lower() == "menu item"
        ],
    }


def click_menu_item_in_popup(jab, windows: list[dict], menu_name: str, popup_hwnd=None):
    candidates = []
    for window in windows:
        if not window.get("is_java") or not window.get("visible"):
            continue
        if popup_hwnd is not None and window.get("hwnd") != popup_hwnd:
            continue
        for control in window.get("all_controls", []):
            if not is_current_visible_control(control):
                continue
            if not control.get("accessibleAction"):
                continue
            if (
                str(control.get("role", "")).lower() == "menu item"
                and control.get("name") == menu_name
            ):
                candidates.append({"window": window, "control": control})

    if not candidates:
        return {
            "ok": False,
            "reason": f"菜单项未找到: {menu_name}",
            "candidate_count": 0,
        }

    target = candidates[0]
    ok = do_action_by_window_path(
        jab,
        target["window"]["hwnd"],
        target["control"]["path"],
        action_name=choose_click_action(target["control"].get("actions", [])),
    )
    return {
        "ok": bool(ok),
        "target": {
            "window": {
                key: target["window"].get(key)
                for key in ("hwnd", "title", "class_name", "visible")
            },
            "control": summarize_control(target["control"]),
        },
        "candidate_count": len(candidates),
    }


def close_popup_hwnd(hwnd):
    if os.name != "nt":
        return {"ok": False, "reason": "Windows only", "hwnd": hwnd}
    if not hwnd:
        return {"ok": False, "reason": "missing hwnd"}
    user32 = ctypes.windll.user32
    hwnd_obj = wintypes.HWND(int(hwnd))
    before = describe_hwnd(user32, hwnd_obj)
    if not before.get("exists"):
        return {"ok": True, "reason": "already gone", "before": before}
    if before.get("class_name") != "SunAwtWindow":
        return {"ok": False, "reason": "class mismatch", "before": before}
    user32.ShowWindow(hwnd_obj, 0)
    user32.SetWindowPos(
        hwnd_obj,
        0,
        -32000,
        -32000,
        0,
        0,
        0x0001 | 0x0010 | 0x0080 | 0x0200,
    )
    user32.PostMessageW(hwnd_obj, 0x0010, 0, 0)
    return {"ok": True, "before": before, "after": describe_hwnd(user32, hwnd_obj)}


def is_visible_sun_awt_popup(window):
    if window.get("class_name") != "SunAwtWindow" or not window.get("visible"):
        return False
    root = window.get("root") or {}
    bounds = root.get("bounds") or []
    if len(bounds) != 4:
        return False
    _x, _y, width, height = bounds
    if width <= 0 or height <= 0:
        return False
    return width <= 500 and height <= 500


def collect_controls(
    jab,
    vm_id,
    context,
    path,
    controls,
    all_controls,
    depth,
    max_depth,
    max_children,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    item = summarize_info(jab, vm_id, context, info, path)
    all_controls.append(item)
    if keep_control(item):
        controls.append(item)
    if depth >= max_depth:
        return
    for index in range(min(info.childrenCount, max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        collect_controls(
            jab,
            vm_id,
            child,
            f"{path}.{index}",
            controls,
            all_controls,
            depth + 1,
            max_depth,
            max_children,
        )
        jab.release_contexts(vm_id, [child])


def keep_control(item):
    role = str(item.get("role", "")).lower()
    return role in {"menu item", "menu"} or bool(item.get("accessibleAction"))


def summarize_context(jab, vm_id, context, path):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    return summarize_info(jab, vm_id, context, info, path)


def summarize_info(jab, vm_id, context, info, path):
    role = info.role_en_US.strip() or info.role.strip()
    item = {
        "path": path,
        "role": role,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "bounds": [info.x, info.y, info.width, info.height],
        "children": info.childrenCount,
        "accessibleAction": bool(info.accessibleAction),
        "actions": [],
    }
    if info.accessibleAction:
        item["actions"] = get_action_names(jab, vm_id, context)
    return item


def get_action_names(jab, vm_id, context):
    if not hasattr(jab.dll, "getAccessibleActions"):
        return []
    return jab.get_action_names(vm_id, context)


def is_current_visible_control(control):
    states = str(control.get("states") or "").lower()
    bounds = control.get("bounds") or []
    if "visible" not in states or "showing" not in states:
        return False
    return (
        isinstance(bounds, list)
        and len(bounds) == 4
        and bounds[2] > 0
        and bounds[3] > 0
    )


def choose_click_action(actions):
    if not actions:
        return None
    for preferred in ("单击", "click", "press"):
        if preferred in actions:
            return preferred
    return actions[0]


def do_action_by_window_path(jab, hwnd, path, action_name=None):
    result = jab.find_context_by_path_once(
        path,
        scope_hwnd=hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    context, vm_id, owned, _window_info = result
    if not context:
        return False
    try:
        return jab.do_action(
            vm_id,
            context,
            action_name=action_name,
            cleanup_blank_awt=False,
        )
    finally:
        jab.release_contexts(vm_id, owned)


def summarize_control(control):
    return {
        key: control.get(key)
        for key in (
            "path",
            "role",
            "name",
            "description",
            "states",
            "bounds",
            "accessibleAction",
            "actions",
        )
    }


def window_key(window):
    return (
        window.get("hwnd"),
        window.get("class_name"),
        window.get("title"),
    )


def window_signature(window):
    controls = []
    for item in window.get("controls", []):
        controls.append(
            (
                item.get("path"),
                item.get("role"),
                item.get("name"),
                item.get("description"),
                item.get("states"),
            )
        )
    root = window.get("root") or {}
    return (tuple(controls), tuple(root.get("bounds") or []))


def describe_hwnd(user32: Any, hwnd):
    if not user32.IsWindow(hwnd):
        return {"exists": False}

    class Rect(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    title_len = user32.GetWindowTextLengthW(hwnd)
    title = ctypes.create_unicode_buffer(title_len + 1)
    user32.GetWindowTextW(hwnd, title, title_len + 1)
    class_name = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_name, 256)
    rect = Rect()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return {
        "exists": True,
        "hwnd": int(hwnd.value),
        "visible": bool(user32.IsWindowVisible(hwnd)),
        "class_name": class_name.value,
        "title": title.value,
        "rect": [rect.left, rect.top, rect.right, rect.bottom],
        "width": rect.right - rect.left,
        "height": rect.bottom - rect.top,
    }
