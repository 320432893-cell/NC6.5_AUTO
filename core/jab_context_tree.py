# 职责：封装 JAB context 信息读取、控件匹配、树递归查找和 context 释放
# 不做什么：不枚举窗口，不执行控件 action，不读取表格单元格，不处理业务字段
# 允许依赖层：标准库 ctypes、JABOperator 暴露的 dll/max_depth/max_children、tools.jab_probe 基础结构
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import

import ctypes

from core.logger import log
from tools.jab_probe import AccessibleContextInfo


def find_in_tree(
    jab, vm_id, context, name, normalized_roles, require_showing, depth, owned_path
):
    info = get_context_info(jab, vm_id, context)
    if not info:
        return None, []

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    control_name = info.name.strip()
    desc = info.description.strip()
    states = (info.states_en_US.strip() or info.states.strip()).lower()

    if matches_control(
        control_name,
        desc,
        role,
        states,
        name,
        normalized_roles,
        require_showing,
    ):
        return context, list(owned_path)

    if depth >= jab.max_depth:
        return None, []
    if role == "table" and "table" not in normalized_roles:
        return None, []

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue

        found, found_owned = find_in_tree(
            jab,
            vm_id,
            child,
            name,
            normalized_roles,
            require_showing,
            depth + 1,
            owned_path + [child],
        )
        if found:
            return found, found_owned

        release_contexts(jab, vm_id, [child])

    return None, []


def find_in_tree_with_path(
    jab,
    vm_id,
    context,
    name,
    normalized_roles,
    require_showing,
    depth,
    owned_contexts,
    owned_indexes,
):
    info = get_context_info(jab, vm_id, context)
    if not info:
        return None, [], []

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    control_name = info.name.strip()
    desc = info.description.strip()
    states = (info.states_en_US.strip() or info.states.strip()).lower()

    if matches_control(
        control_name,
        desc,
        role,
        states,
        name,
        normalized_roles,
        require_showing,
    ):
        return context, list(owned_contexts), list(owned_indexes)

    if depth >= jab.max_depth:
        return None, [], []
    if role == "table" and "table" not in normalized_roles:
        return None, [], []

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue

        found, found_contexts, found_indexes = find_in_tree_with_path(
            jab,
            vm_id,
            child,
            name,
            normalized_roles,
            require_showing,
            depth + 1,
            owned_contexts + [child],
            owned_indexes + [index],
        )
        if found:
            return found, found_contexts, found_indexes

        release_contexts(jab, vm_id, [child])

    return None, [], []


def get_context_info(jab, vm_id, context):
    info = AccessibleContextInfo()
    if not jab.dll.getAccessibleContextInfo(vm_id, context, ctypes.byref(info)):
        return None
    return info


def matches_control(
    control_name,
    desc,
    role,
    states,
    expected_name,
    normalized_roles,
    require_showing,
):
    if normalized_roles and role not in normalized_roles:
        return False
    if require_showing and ("visible" not in states or "showing" not in states):
        return False
    return control_name == expected_name or desc == expected_name


def release_contexts(jab, vm_id, contexts):
    if not vm_id or not hasattr(jab.dll, "releaseJavaObject"):
        return
    for context in reversed(contexts):
        try:
            jab.dll.releaseJavaObject(vm_id, context)
        except Exception as exc:
            # 句柄释放失败不阻断其余释放,但留 DEBUG 痕迹,避免 JAB 句柄泄漏完全无线索
            log.debug(f"释放 JAB context 失败(忽略): vm_id={vm_id} err={exc}")
