# 职责：封装基于 label 位置关系的 JAB 文本控件查找、诊断和邻近控件动作
# 不做什么：不处理 context path，不读取表格，不发送全局键盘输入，不解释业务字段
# 允许依赖层：标准库 ctypes/time、JABOperator 暴露的 dll/context/action/text 能力、tools.jab_probe 枚举窗口能力
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import

import ctypes
import time

from core.logger import log
from core.utils import check_abort
from tools.jab_probe import JOBJECT, enum_windows


def set_text_near_label(
    jab,
    label,
    text,
    title=None,
    class_name=None,
    wait=None,
    timeout=None,
    require_showing=True,
):
    jab.ensure_started()
    deadline = time.time() + (timeout or jab.search_timeout)

    while time.time() < deadline:
        check_abort()
        result = find_text_context_near_label_once(
            jab,
            label,
            title=title,
            class_name=class_name,
            require_showing=require_showing,
        )
        context, vm_id, owned_contexts, label_info, text_info, window_info = result
        if context:
            try:
                if not jab.set_text_context(vm_id, context, text):
                    return False
                log.info(
                    "JAB label 文本输入成功: "
                    f"label={label!r} text={text!r} "
                    f"label_bounds={label_info.x},{label_info.y},{label_info.width},{label_info.height} "
                    f"text_bounds={text_info.x},{text_info.y},{text_info.width},{text_info.height} "
                    f"hwnd={window_info.get('hwnd')} title={window_info.get('title')!r}"
                )
                time.sleep(jab.menu_wait if wait is None else wait)
                return True
            finally:
                jab.release_contexts(vm_id, owned_contexts)
        time.sleep(0.2)

    log.warning(
        f"JAB 未找到 label 右侧文本控件: label={label!r} title={title!r} class={class_name!r}"
    )
    return False


def find_text_context_near_label_once(
    jab,
    label,
    title=None,
    class_name=None,
    require_showing=True,
):
    windows = enum_windows(include_children=True)
    for hwnd, window_title, window_class, pid, visible in windows:
        if title is not None and window_title != title:
            continue
        if class_name is not None and window_class != class_name:
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

        found = find_text_near_label_by_bounds(
            jab,
            vm_id.value,
            root_context.value,
            label,
            require_showing=require_showing,
        )
        if found[0]:
            context, owned_contexts, label_info, text_info = found
            return (
                context,
                vm_id.value,
                owned_contexts,
                label_info,
                text_info,
                {
                    "hwnd": int(hwnd),
                    "title": window_title,
                    "class": window_class,
                    "pid": pid,
                    "visible": visible,
                },
            )

    return None, None, [], None, None, {}


def describe_controls_near_label(
    jab,
    label,
    title=None,
    class_name=None,
    require_showing=True,
    max_vertical_distance=28,
    max_right_distance=420,
):
    """Read-only diagnostic for controls near a label row."""
    jab.ensure_started()
    windows = []

    for hwnd, window_title, window_class, pid, visible in enum_windows(
        include_children=True
    ):
        if title is not None and window_title != title:
            continue
        if class_name is not None and window_class != class_name:
            continue

        window_result = {
            "hwnd": int(hwnd),
            "title": window_title,
            "class": window_class,
            "pid": pid,
            "visible": visible,
            "is_java": bool(jab.dll.isJavaWindow(hwnd)),
            "labels": [],
        }
        windows.append(window_result)
        if not window_result["is_java"]:
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            window_result["error"] = "getAccessibleContextFromHWND failed"
            continue

        controls = []
        owned = []
        try:
            collect_controls_for_bounds_scan(
                jab,
                vm_id.value,
                root_context.value,
                controls,
                owned,
                require_showing=require_showing,
                depth=0,
            )
            labels = [
                item
                for item in controls
                if item[1].role_en_US.strip().lower() == "label"
                and item[1].name.strip() == label
            ]
            labels.sort(key=lambda item: (item[1].y, item[1].x))
            for _label_context, label_info in labels:
                nearby = [
                    {
                        "control": info_to_dict(item["info"]),
                        "dy": item["dy"],
                        "right_distance": item["right_distance"],
                        "actions": item["actions"],
                    }
                    for item in controls_near_label_info(
                        jab,
                        vm_id.value,
                        controls,
                        label_info,
                        max_vertical_distance=max_vertical_distance,
                        max_right_distance=max_right_distance,
                    )
                ]
                nearby.sort(
                    key=lambda item: (
                        item["control"]["y"],
                        item["control"]["x"],
                        item["control"]["role"],
                    )
                )
                window_result["labels"].append(
                    {
                        "label": info_to_dict(label_info),
                        "nearby": nearby,
                    }
                )
        finally:
            jab.release_contexts(vm_id.value, owned)

    return windows


def controls_near_label_info(
    jab,
    vm_id,
    controls,
    label_info,
    max_vertical_distance=28,
    max_right_distance=420,
):
    label_mid_y = label_info.y + label_info.height / 2
    label_right = label_info.x + label_info.width
    nearby = []
    for control_context, control_info in controls:
        if control_info is label_info:
            continue
        control_mid_y = control_info.y + control_info.height / 2
        dy = control_mid_y - label_mid_y
        right_distance = control_info.x - label_right
        if abs(dy) > max_vertical_distance:
            continue
        if right_distance < -40 or right_distance > max_right_distance:
            continue
        nearby.append(
            {
                "context": control_context,
                "info": control_info,
                "dy": dy,
                "right_distance": right_distance,
                "actions": jab.get_action_names(vm_id, control_context),
            }
        )
    nearby.sort(
        key=lambda item: (
            item["info"].y,
            item["info"].x,
            item["info"].role_en_US.strip(),
        )
    )
    return nearby


def click_control_near_label(
    jab,
    label,
    role,
    index=0,
    title=None,
    class_name=None,
    require_showing=True,
    max_vertical_distance=28,
    max_right_distance=420,
    action_name=None,
    wait=None,
):
    jab.ensure_started()
    role = str(role).strip().lower()

    for hwnd, window_title, window_class, pid, visible in enum_windows(
        include_children=True
    ):
        if title is not None and window_title != title:
            continue
        if class_name is not None and window_class != class_name:
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

        controls = []
        owned = []
        try:
            collect_controls_for_bounds_scan(
                jab,
                vm_id.value,
                root_context.value,
                controls,
                owned,
                require_showing=require_showing,
                depth=0,
            )
            labels = [
                item
                for item in controls
                if item[1].role_en_US.strip().lower() == "label"
                and item[1].name.strip() == label
            ]
            labels.sort(key=lambda item: (item[1].y, item[1].x))
            for _label_context, label_info in labels:
                candidates = controls_near_label_info(
                    jab,
                    vm_id.value,
                    controls,
                    label_info,
                    max_vertical_distance=max_vertical_distance,
                    max_right_distance=max_right_distance,
                )
                filtered = [
                    item
                    for item in candidates
                    if item["info"].role_en_US.strip().lower() == role
                ]
                filtered.sort(key=lambda item: (item["info"].x, item["info"].y))
                if index >= len(filtered):
                    continue

                target = filtered[index]
                context = target["context"]
                info = target["info"]
                actions = jab.get_action_names(vm_id.value, context)
                ok = jab.do_action(vm_id.value, context, action_name=action_name)
                if wait is not None:
                    time.sleep(wait)
                return {
                    "ok": bool(ok),
                    "method": "action",
                    "window": {
                        "hwnd": int(hwnd),
                        "title": window_title,
                        "class": window_class,
                        "pid": pid,
                        "visible": visible,
                    },
                    "label": info_to_dict(label_info),
                    "control": info_to_dict(info),
                    "actions": actions,
                    "index": index,
                }
        finally:
            jab.release_contexts(vm_id.value, owned)

    return {"ok": False, "role": role, "index": index}


def describe_text_near_label(
    jab,
    label,
    title=None,
    class_name=None,
    require_showing=True,
):
    """Read-only diagnostic for label-to-text matching."""
    jab.ensure_started()
    windows = []

    for hwnd, window_title, window_class, pid, visible in enum_windows(
        include_children=True
    ):
        if title is not None and window_title != title:
            continue
        if class_name is not None and window_class != class_name:
            continue

        window_result = {
            "hwnd": int(hwnd),
            "title": window_title,
            "class": window_class,
            "pid": pid,
            "visible": visible,
            "is_java": bool(jab.dll.isJavaWindow(hwnd)),
            "labels": [],
        }
        windows.append(window_result)
        if not window_result["is_java"]:
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            window_result["error"] = "getAccessibleContextFromHWND failed"
            continue

        labels = []
        texts = []
        owned = []
        try:
            collect_labels_and_texts(
                jab,
                vm_id.value,
                root_context.value,
                label,
                labels,
                texts,
                owned,
                require_showing=require_showing,
                depth=0,
            )
            labels.sort(key=lambda item: (item[1].y, item[1].x))
            for _label_context, label_info in labels:
                label_mid_y = label_info.y + label_info.height / 2
                label_result = {
                    "label": info_to_dict(label_info),
                    "selected": None,
                    "candidates": [],
                }
                for _text_context, text_info in texts:
                    text_mid_y = text_info.y + text_info.height / 2
                    dy = text_mid_y - label_mid_y
                    rejected = []
                    if text_info.x <= label_info.x + label_info.width:
                        rejected.append("not_right_of_label")
                    if abs(dy) > 6:
                        rejected.append("different_row")
                    label_result["candidates"].append(
                        {
                            "text": info_to_dict(text_info),
                            "dy": dy,
                            "rejected": rejected,
                        }
                    )

                label_result["candidates"].sort(
                    key=lambda item: (
                        bool(item["rejected"]),
                        item["text"]["x"],
                        abs(item["dy"]),
                        item["text"]["y"],
                    )
                )
                usable = [
                    item for item in label_result["candidates"] if not item["rejected"]
                ]
                if usable:
                    label_result["selected"] = usable[0]
                window_result["labels"].append(label_result)
        finally:
            jab.release_contexts(vm_id.value, owned)

    return windows


def collect_controls_for_bounds_scan(
    jab,
    vm_id,
    context,
    controls,
    owned,
    require_showing=True,
    depth=0,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if role == "table" or depth >= jab.max_depth:
        return

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_info = jab.get_context_info(vm_id, child)
        if not child_info:
            jab.release_contexts(vm_id, [child])
            continue

        owned.append(child)
        states = (child_info.states_en_US.strip() or child_info.states.strip()).lower()
        showing = "visible" in states and "showing" in states
        if not require_showing or showing:
            controls.append((child, child_info))

        collect_controls_for_bounds_scan(
            jab,
            vm_id,
            child,
            controls,
            owned,
            require_showing=require_showing,
            depth=depth + 1,
        )


def info_to_dict(info):
    return {
        "name": info.name.strip(),
        "description": info.description.strip(),
        "role": (info.role_en_US.strip() or info.role.strip()),
        "states": (info.states_en_US.strip() or info.states.strip()),
        "index_in_parent": info.indexInParent,
        "children_count": info.childrenCount,
        "x": info.x,
        "y": info.y,
        "width": info.width,
        "height": info.height,
    }


def find_text_near_label_by_bounds(
    jab,
    vm_id,
    root_context,
    label,
    require_showing=True,
):
    labels = []
    texts = []
    owned = []
    collect_labels_and_texts(
        jab,
        vm_id,
        root_context,
        label,
        labels,
        texts,
        owned,
        require_showing=require_showing,
        depth=0,
    )
    if not labels:
        jab.release_contexts(vm_id, owned)
        return None, [], None, None

    labels.sort(key=lambda item: (item[1].y, item[1].x))
    for label_context, label_info in labels:
        if not jab.context_info_has_valid_bounds(label_info):
            continue
        label_mid_y = label_info.y + label_info.height / 2
        candidates = []
        for text_context, text_info in texts:
            if not jab.context_info_has_valid_bounds(text_info):
                continue
            text_mid_y = text_info.y + text_info.height / 2
            if text_info.x <= label_info.x + label_info.width:
                continue
            if abs(text_mid_y - label_mid_y) > 6:
                continue
            candidates.append((text_context, text_info))
        if candidates:
            candidates.sort(key=lambda item: (item[1].x, item[1].y))
            text_context, text_info = candidates[0]
            keep = {label_context, text_context}
            release = [context for context in owned if context not in keep]
            jab.release_contexts(vm_id, release)
            return (
                text_context,
                [label_context, text_context],
                label_info,
                text_info,
            )

    jab.release_contexts(vm_id, owned)
    return None, [], None, None


def collect_labels_and_texts(
    jab,
    vm_id,
    context,
    target_label,
    labels,
    texts,
    owned,
    require_showing=True,
    depth=0,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if role == "table" or depth >= jab.max_depth:
        return

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_info = jab.get_context_info(vm_id, child)
        if not child_info:
            jab.release_contexts(vm_id, [child])
            continue

        owned.append(child)
        child_role = (child_info.role_en_US.strip() or child_info.role.strip()).lower()
        child_states = (
            child_info.states_en_US.strip() or child_info.states.strip()
        ).lower()
        showing = "visible" in child_states and "showing" in child_states

        if (
            child_role == "label"
            and child_info.name.strip() == target_label
            and (not require_showing or showing)
            and jab.context_info_has_valid_bounds(child_info)
        ):
            labels.append((child, child_info))
        elif (
            child_role == "text"
            and "editable" in child_states
            and (not require_showing or showing)
            and jab.context_info_has_valid_bounds(child_info)
        ):
            texts.append((child, child_info))

        collect_labels_and_texts(
            jab,
            vm_id,
            child,
            target_label,
            labels,
            texts,
            owned,
            require_showing=require_showing,
            depth=depth + 1,
        )


def find_text_near_label_in_tree(
    jab,
    vm_id,
    context,
    label,
    require_showing=True,
    depth=0,
    owned_path=None,
):
    owned_path = owned_path or []
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None, [], None, None

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if role == "table":
        return None, [], None, None

    children = []
    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_info = jab.get_context_info(vm_id, child)
        if child_info:
            children.append((child, child_info))

    for label_context, label_info in children:
        label_role = (label_info.role_en_US.strip() or label_info.role.strip()).lower()
        label_states = (
            label_info.states_en_US.strip() or label_info.states.strip()
        ).lower()
        if label_role != "label" or label_info.name.strip() != label:
            continue
        if require_showing and (
            "visible" not in label_states or "showing" not in label_states
        ):
            continue

        candidates = find_text_candidates_right_of_label(
            jab,
            vm_id,
            context,
            label_info,
            require_showing=require_showing,
            depth=0,
            owned_path=[],
            skip_contexts={label_context},
        )

        if candidates:
            candidates.sort(key=lambda item: (item[1].x, item[1].y))
            text_context, text_info, text_owned = candidates[0]
            owned = list(owned_path) + [label_context] + text_owned
            for extra_context, _extra_info in children:
                if extra_context not in (label_context, text_owned[0]):
                    jab.release_contexts(vm_id, [extra_context])
            return text_context, owned, label_info, text_info

    if depth >= jab.max_depth:
        jab.release_contexts(vm_id, [child for child, _info in children])
        return None, [], None, None

    for child, child_info in children:
        child_role = (child_info.role_en_US.strip() or child_info.role.strip()).lower()
        if child_role == "table":
            jab.release_contexts(vm_id, [child])
            continue
        found = find_text_near_label_in_tree(
            jab,
            vm_id,
            child,
            label,
            require_showing=require_showing,
            depth=depth + 1,
            owned_path=owned_path + [child],
        )
        if found[0]:
            for sibling, _info in children:
                if sibling != child:
                    jab.release_contexts(vm_id, [sibling])
            return found
        jab.release_contexts(vm_id, [child])

    return None, [], None, None


def find_text_candidates_right_of_label(
    jab,
    vm_id,
    context,
    label_info,
    require_showing=True,
    depth=0,
    owned_path=None,
    skip_contexts=None,
):
    owned_path = owned_path or []
    skip_contexts = skip_contexts or set()
    info = jab.get_context_info(vm_id, context)
    if not info:
        return []

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    states = (info.states_en_US.strip() or info.states.strip()).lower()
    label_mid_y = label_info.y + label_info.height / 2

    candidates = []
    if context not in skip_contexts and role == "text" and "editable" in states:
        if not require_showing or ("visible" in states and "showing" in states):
            text_mid_y = info.y + info.height / 2
            if info.x > label_info.x + label_info.width and abs(
                text_mid_y - label_mid_y
            ) <= max(label_info.height, 24):
                return [(context, info, list(owned_path))]

    if depth >= jab.max_depth or role == "table":
        return []

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_candidates = find_text_candidates_right_of_label(
            jab,
            vm_id,
            child,
            label_info,
            require_showing=require_showing,
            depth=depth + 1,
            owned_path=owned_path + [child],
            skip_contexts=skip_contexts,
        )
        if child_candidates:
            candidates.extend(child_candidates)
        else:
            jab.release_contexts(vm_id, [child])

    return candidates
