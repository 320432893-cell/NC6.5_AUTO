# 职责：新增弹窗/窗口枚举与控件收集、前台根标注、窗口签名差分
# 不做什么：不做 CLI 解析/不起 JABOperator(那是 receipt_new_probe 主入口)
# 允许依赖层：标准库、core JAB(经 jab 参数)、tools.jab_probe、tools.receipt_new_* 同层
# 谁不应该 import：core 层模块不应 import

import ctypes
import json
import os
import sys

from tools.jab_probe import JOBJECT, enum_windows

from tools.receipt_new_menu import ENTRY_STATE_NAMES, SELF_MADE_NAMES



class _ProbeNamespace:
    # 调用时从已加载的 receipt_new_probe 读顶层函数,使测试对
    # tools.receipt_new_probe.<name> 的 monkeypatch 与拆分前一致生效,且不在加载期 import 成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_new_probe"], name)


_probe = _ProbeNamespace()


def collect_entry_context_snapshot(jab):
    windows = _probe.collect_receipt_new_windows_compat(jab, max_depth=0, max_children=0)
    anchor = _probe.resolve_current_canvas_header_anchor(jab, windows)
    state = {
        "ok": False,
        "partial_ok": bool(anchor.get("ok")),
        "names": [],
        "hits": [],
        "reason": "self-made action succeeded; header 财务组织(O) anchor will confirm entry scope",
    }
    if anchor.get("ok"):
        state["reason"] = "财务组织(O) anchor resolved in current canvas"
        state["hits"].append(
            {
                "window": anchor.get("window") or {},
                "control": {
                    "path": anchor.get("label_path"),
                    "role": "label",
                    "name": ((anchor.get("anchor_text") or {}).get("name") or ""),
                    "description": (
                        (anchor.get("anchor_text") or {}).get("description") or ""
                    ),
                    "dynamic_index": anchor.get("dynamic_index"),
                    "dynamic_prefix": anchor.get("dynamic_prefix"),
                },
            }
        )
    return {
        "ok": True,
        "confirmed": bool(anchor.get("ok")),
        "state": state,
        "windows": windows,
        "anchor": anchor,
    }



def resolve_current_canvas_header_anchor(jab, windows):
    canvas_hwnds = [
        int(window["hwnd"])
        for window in windows or []
        if window.get("is_java")
        and window.get("visible")
        and window.get("hwnd")
        and window.get("class_name") == "SunAwtCanvas"
    ]
    if not canvas_hwnds:
        return {"ok": False, "reason": "current SunAwtCanvas not found"}
    from tools.receipt_self_made_fill_trial import (
        resolve_receipt_header_anchor_in_canvas,
    )

    attempts = []
    for hwnd in canvas_hwnds:
        attempt = resolve_receipt_header_anchor_in_canvas(jab, hwnd, timeout=0.4)
        attempts.append(attempt)
        if attempt.get("ok"):
            return attempt
    return {
        "ok": False,
        "reason": "财务组织(O) anchor not resolved in current canvas",
        "attempts": attempts,
    }



def collect_new_visible_popup_windows(jab, before, max_depth=8, max_children=120):
    if os.name != "nt" or not hasattr(ctypes, "WinDLL"):
        return _probe.collect_receipt_new_windows(jab)

    before_signatures = {_probe.window_key(item): _probe.window_signature(item) for item in before}
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
            "root": _probe.summarize_context(jab, vm_id.value, root_context.value, "0"),
            "controls": [],
            "all_controls": [],
        }
        key = _probe.window_key(window)
        if not _probe.is_visible_sun_awt_popup(window):
            windows.append(window)
            continue
        _probe.collect_controls(
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
        if key in before_signatures and before_signatures[key] == _probe.window_signature(
            window
        ):
            continue
        windows.append(window)
    return windows



def collect_receipt_new_windows_compat(jab, **kwargs):
    try:
        return _probe.collect_receipt_new_windows(jab, **kwargs)
    except TypeError as exc:
        if "unexpected keyword argument" not in str(exc):
            raise
        return _probe.collect_receipt_new_windows(jab)



def find_named_controls_in_windows(
    windows,
    name_query="新增",
    role=None,
    class_name=None,
    require_action=True,
):
    results = []
    name_query = str(name_query or "").lower()
    role = role.lower() if role else None
    for window in windows or []:
        if class_name and window.get("class_name") != class_name:
            continue
        if not window.get("is_java"):
            continue
        for control in window.get("all_controls", []):
            control_role = control.get("role", "").lower()
            if role and control_role != role:
                continue
            text = f"{control.get('name', '')} {control.get('description', '')}".lower()
            if name_query and name_query not in text:
                continue
            if require_action and not control.get("accessibleAction"):
                continue
            results.append(
                {
                    "window": {
                        key: window.get(key)
                        for key in ("hwnd", "title", "class_name", "visible")
                    },
                    "control": control,
                }
            )
    return results



def annotate_foreground_root_for_targets(targets, foreground):
    for item in targets or []:
        window = item.get("window") or {}
        window["root_hwnd"] = _probe.root_hwnd(window.get("hwnd"))
        window["is_foreground_root"] = bool(
            foreground
            and foreground.get("root")
            and window.get("root_hwnd") == foreground.get("root")
        )



def annotate_foreground_root(windows, foreground):
    for window in windows or []:
        window["root_hwnd"] = _probe.root_hwnd(window.get("hwnd"))
        window["is_foreground_root"] = bool(
            foreground
            and foreground.get("root")
            and window.get("root_hwnd") == foreground.get("root")
        )



def collect_receipt_new_windows(jab, max_depth=25, max_children=1000):
    windows = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if not class_name.startswith(("SunAwt", "Yonyou")):
            continue
        is_java = bool(jab.dll.isJavaWindow(hwnd))
        window = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
            "is_java": is_java,
            "root": None,
            "controls": [],
            "all_controls": [],
        }
        windows.append(window)
        if not is_java:
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            continue
        window["root"] = _probe.summarize_context(jab, vm_id.value, root_context.value, "0")
        _probe.collect_controls(
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
    return windows



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
    item = _probe.summarize_info(jab, vm_id, context, info, path)
    all_controls.append(item)
    if _probe.keep_control(item):
        controls.append(item)
    if depth >= max_depth:
        return
    for index in range(min(info.childrenCount, max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        _probe.collect_controls(
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
    role = item["role"].lower()
    if role in {
        "menu item",
        "menu",
        "push button",
        "text",
        "page tab",
        "page tab list",
    }:
        return True
    texts = {item["name"], item["description"]}
    if texts & SELF_MADE_NAMES or _probe.normalize_entry_state_names(item) & ENTRY_STATE_NAMES:
        return True
    if item["accessibleAction"]:
        return True
    return False



def find_new_visible_popup(before, after):
    before_signatures = {_probe.window_key(item): _probe.window_signature(item) for item in before}
    candidates = []
    for item in after:
        key = _probe.window_key(item)
        if key in before_signatures and before_signatures[key] == _probe.window_signature(
            item
        ):
            continue
        if not _probe.is_visible_sun_awt_popup(item):
            continue
        menu_names = {
            control.get("name")
            for control in item.get("all_controls", [])
            if control.get("role", "").lower() == "menu item"
        }
        if SELF_MADE_NAMES & menu_names or "应收单" in menu_names:
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
            _probe.summarize_control(control)
            for control in item.get("all_controls", [])
            if control.get("role", "").lower() == "menu item"
        ],
    }



def diff_windows(before, after):
    before_signatures = {_probe.window_key(item): _probe.window_signature(item) for item in before}
    changed = []
    for item in after:
        key = _probe.window_key(item)
        sig = _probe.window_signature(item)
        if key not in before_signatures or before_signatures[key] != sig:
            changed.append(item)
    return changed



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
                tuple(item.get("actions") or []),
            )
        )
    return (
        json.dumps(window.get("root"), ensure_ascii=False, sort_keys=True),
        tuple(controls),
    )
