# 职责：JAB 无障碍树遍历小工具——按窗口/标签/文本定位 context、根窗口句柄、提交动作
# 不做什么：不算 path 模板,不做表头字段写入,不解析 scope,不读客户名候选
# 允许依赖层：标准库 + ctypes + tools.jab_probe(惰性) + tools.receipt_header_paths(label 匹配)
# 谁不应该 import：receipt_header_paths 不应 import 本模块(会成环)

import ctypes
from ctypes import wintypes
import os
import time

from tools.receipt_header_paths import (
    header_label_text_matches,
    header_scope_anchor_text_matches,
)


def find_context_with_window(
    jab,
    name,
    roles=(),
    timeout=None,
    require_showing=False,
    window_title=None,
    window_class=None,
    visible_only=True,
    scope_hwnd=None,
):
    deadline = time.time() + (timeout or jab.search_timeout)
    normalized_roles = {role.lower() for role in roles}
    while time.time() < deadline:
        windows = jab.get_scoped_windows(scope_hwnd, include_children=True)
        for hwnd, title, class_name, pid, visible in windows:
            if visible_only and not visible:
                continue
            if (
                scope_hwnd is None
                and window_title is not None
                and title != window_title
            ):
                continue
            if (
                scope_hwnd is None
                and window_class is not None
                and class_name != window_class
            ):
                continue
            if not jab.dll.isJavaWindow(hwnd):
                continue
            from tools.jab_probe import JOBJECT

            vm_id_ref = ctypes.c_long()
            root_context = JOBJECT()
            if not jab.dll.getAccessibleContextFromHWND(
                hwnd,
                ctypes.byref(vm_id_ref),
                ctypes.byref(root_context),
            ):
                continue
            context, owned_contexts, owned_indexes = jab.find_in_tree_with_path(
                vm_id_ref.value,
                root_context.value,
                name,
                normalized_roles,
                require_showing,
                depth=0,
                owned_contexts=[],
                owned_indexes=[],
            )
            if context:
                return (
                    context,
                    vm_id_ref.value,
                    owned_contexts,
                    owned_indexes,
                    {
                        "hwnd": int(hwnd),
                        "title": title,
                        "class_name": class_name,
                        "pid": pid,
                        "visible": visible,
                    },
                )
            jab.release_contexts(vm_id_ref.value, [root_context.value])
        time.sleep(0.2)
    return None, None, [], [], {}


def find_label_following_text(jab, vm_id, context, label, path, depth, owned_contexts):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if depth >= jab.max_depth or role == "table":
        return None

    child_infos = []
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_info = jab.get_context_info(vm_id, child)
        child_path = f"{path}.{index}"
        if child_info:
            child_infos.append((index, child, child_info, child_path))
        else:
            jab.release_contexts(vm_id, [child])

    for position, (_index, child, child_info, child_path) in enumerate(child_infos):
        child_role = (child_info.role_en_US.strip() or child_info.role.strip()).lower()
        child_states = (
            child_info.states_en_US.strip() or child_info.states.strip()
        ).lower()
        child_texts = {child_info.name.strip(), child_info.description.strip()}
        if child_role == "label" and label in child_texts and "visible" in child_states:
            for _next_index, next_child, _next_info, next_path in child_infos[
                position + 1 :
            ]:
                text_context, text_owned, text_path = first_text_descendant(
                    jab,
                    vm_id,
                    next_child,
                    next_path,
                    depth + 1,
                )
                if text_context:
                    keep = [item[1] for item in child_infos] + text_owned
                    return text_context, owned_contexts + keep, text_path, child_path

    for _index, child, _child_info, child_path in child_infos:
        result = find_label_following_text(
            jab,
            vm_id,
            child,
            label,
            child_path,
            depth + 1,
            owned_contexts + [item[1] for item in child_infos],
        )
        if result:
            return result
    for _index, child, _child_info, _child_path in child_infos:
        jab.release_contexts(vm_id, [child])
    return None


def first_text_descendant(jab, vm_id, context, path, depth):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None, [], None
    role = (info.role_en_US.strip() or info.role.strip()).lower()
    states = (info.states_en_US.strip() or info.states.strip()).lower()
    if role == "text" and "visible" in states and "editable" in states:
        return context, [], path
    if depth >= jab.max_depth or role == "table":
        return None, [], None
    owned = []
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_path = f"{path}.{index}"
        found, found_owned, found_path = first_text_descendant(
            jab,
            vm_id,
            child,
            child_path,
            depth + 1,
        )
        if found:
            return found, owned + [child] + found_owned, found_path
        jab.release_contexts(vm_id, [child])
    return None, owned, None


def find_header_label_context_with_window(
    jab,
    label,
    timeout=1.5,
    require_showing=True,
    scope_hwnd=None,
    strict_anchor=False,
):
    deadline = time.time() + max(float(timeout or 0), 0.0)
    last_window_count = 0
    while time.time() < deadline:
        windows = jab.get_scoped_windows(scope_hwnd, include_children=True)
        last_window_count = len(windows)
        for hwnd, title, class_name, pid, visible in windows:
            if not visible or class_name != "SunAwtCanvas":
                continue
            if not jab.dll.isJavaWindow(hwnd):
                continue
            from tools.jab_probe import JOBJECT

            vm_id_ref = ctypes.c_long()
            root_context = JOBJECT()
            if not jab.dll.getAccessibleContextFromHWND(
                hwnd,
                ctypes.byref(vm_id_ref),
                ctypes.byref(root_context),
            ):
                continue
            context, owned_contexts, owned_indexes = find_header_label_in_tree(
                jab,
                vm_id_ref.value,
                root_context.value,
                label,
                require_showing,
                strict_anchor,
                depth=0,
                owned_contexts=[],
                owned_indexes=[],
            )
            if context:
                return (
                    context,
                    vm_id_ref.value,
                    owned_contexts,
                    owned_indexes,
                    {
                        "hwnd": int(hwnd),
                        "title": title,
                        "class_name": class_name,
                        "pid": pid,
                        "visible": visible,
                    },
                )
            jab.release_contexts(vm_id_ref.value, [root_context.value])
        time.sleep(0.1)
    return None, None, [], [], {"window_count": last_window_count}


def find_header_label_in_tree(
    jab,
    vm_id,
    context,
    label,
    require_showing,
    strict_anchor,
    depth,
    owned_contexts,
    owned_indexes,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None, [], []
    role = (info.role_en_US.strip() or info.role.strip()).lower()
    states = (info.states_en_US.strip() or info.states.strip()).lower()
    if (
        role == "label"
        and (
            header_scope_anchor_text_matches(info)
            if strict_anchor
            else header_label_text_matches(info, label)
        )
        and (not require_showing or ("visible" in states and "showing" in states))
    ):
        return context, list(owned_contexts), list(owned_indexes)
    if depth >= min(jab.max_depth, 50):
        return None, [], []
    if role == "table":
        return None, [], []
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        found, found_contexts, found_indexes = find_header_label_in_tree(
            jab,
            vm_id,
            child,
            label,
            require_showing,
            strict_anchor,
            depth + 1,
            owned_contexts + [child],
            owned_indexes + [index],
        )
        if found:
            return found, found_contexts, found_indexes
        jab.release_contexts(vm_id, [child])
    return None, [], []


def window_root_hwnd(hwnd):
    if os.name != "nt" or not hasattr(ctypes, "windll") or not hwnd:
        return 0
    return int(ctypes.windll.user32.GetAncestor(wintypes.HWND(int(hwnd)), 2) or 0)


def do_context_commit_action(jab, vm_id, context):
    actions = jab.get_action_names(vm_id, context)
    preferred = ("确认", "确定", "提交", "单击", "click", "press")
    for action_name in preferred:
        if action_name not in actions:
            continue
        try:
            ok = bool(jab.do_action(vm_id, context, action_name=action_name))
        except Exception as exc:
            return {
                "ok": False,
                "action": action_name,
                "exception": repr(exc),
                "actions": actions,
            }
        if ok:
            return {"ok": True, "action": action_name, "actions": actions}
    return {"ok": False, "reason": "no commit action", "actions": actions}
