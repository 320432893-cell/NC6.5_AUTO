# 职责：封装基于 JAB context path 的控件定位、动作触发和文本读写
# 不做什么：不做 near-label 搜索，不读取表格，不管理业务保存流程，不发送全局键盘输入
# 允许依赖层：标准库 ctypes/threading/time、JABOperator 暴露的 dll/上下文/动作能力、core.jab_probe 基础结构
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import

import ctypes
import threading
import time

from core.logger import log
from core.utils import check_abort
from core.jab_probe import AccessibleTextInfo, JOBJECT, enum_windows


def do_action_by_path(
    jab,
    path,
    title=None,
    class_name=None,
    scope_hwnd=None,
    name=None,
    role=None,
    action_name=None,
    click_mode=None,
    wait=None,
    timeout=None,
    require_showing=True,
    require_valid_bounds=True,
    cleanup_blank_awt=False,
):
    jab.ensure_started()
    if not path:
        raise ValueError("JAB action path is required")

    deadline = time.time() + (timeout or jab.search_timeout)
    while time.time() < deadline:
        check_abort()
        result = find_context_by_path_once(
            jab,
            path,
            title=title,
            class_name=class_name,
            scope_hwnd=scope_hwnd,
            name=name,
            role=role,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
        )
        context, vm_id, owned_contexts, window_info = result
        if context:
            try:
                if click_mode == "bounds":
                    log.warning(
                        "JAB bounds 点击已禁用: "
                        f"path={path} hwnd={window_info.get('hwnd')} "
                        f"title={window_info.get('title')!r}"
                    )
                    ok = False
                else:
                    ok = jab.do_action(
                        vm_id,
                        context,
                        action_name=action_name,
                        cleanup_blank_awt=cleanup_blank_awt,
                    )
                if ok:
                    log.info(
                        "JAB path 动作成功: "
                        f"path={path} mode={click_mode or 'action'} "
                        f"action={action_name or '<default>'} "
                        f"hwnd={window_info.get('hwnd')} "
                        f"title={window_info.get('title')!r}"
                    )
                    time.sleep(jab.menu_wait if wait is None else wait)
                return ok
            finally:
                jab.release_contexts(vm_id, owned_contexts)
        time.sleep(0.2)

    log.warning(
        "JAB path 未找到可执行控件: "
        f"path={path} title={title!r} class={class_name!r} scope_hwnd={scope_hwnd!r}"
    )
    return False


def set_text_by_path(
    jab,
    path,
    text,
    title=None,
    class_name=None,
    scope_hwnd=None,
    name=None,
    role=None,
    guard_path=None,
    guard_name=None,
    guard_role=None,
    wait=None,
    timeout=None,
    require_showing=True,
    require_valid_bounds=True,
):
    jab.ensure_started()
    if not path:
        raise ValueError("JAB text path is required")

    deadline = time.time() + (timeout or jab.search_timeout)
    while time.time() < deadline:
        check_abort()
        if guard_path:
            guard = find_context_by_path_once(
                jab,
                guard_path,
                title=title,
                class_name=class_name,
                scope_hwnd=scope_hwnd,
                name=guard_name,
                role=guard_role,
                require_showing=require_showing,
                require_valid_bounds=require_valid_bounds,
            )
            guard_context, guard_vm_id, guard_owned_contexts, _guard_window = guard
            if guard_context:
                jab.release_contexts(guard_vm_id, guard_owned_contexts)
            else:
                time.sleep(0.2)
                continue

        result = find_context_by_path_once(
            jab,
            path,
            title=title,
            class_name=class_name,
            scope_hwnd=scope_hwnd,
            name=name,
            role=role,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
        )
        context, vm_id, owned_contexts, window_info = result
        if context:
            try:
                if not set_text_context(jab, vm_id, context, text):
                    return False
                log.info(
                    "JAB path 文本输入成功: "
                    f"path={path} text={text!r} "
                    f"hwnd={window_info.get('hwnd')} "
                    f"title={window_info.get('title')!r}"
                )
                time.sleep(jab.menu_wait if wait is None else wait)
                return True
            finally:
                jab.release_contexts(vm_id, owned_contexts)
        time.sleep(0.2)

    log.warning(
        "JAB path 未找到文本控件: "
        f"path={path} title={title!r} class={class_name!r} scope_hwnd={scope_hwnd!r}"
    )
    return False


def get_text_by_path(
    jab,
    path,
    title=None,
    class_name=None,
    scope_hwnd=None,
    name=None,
    role=None,
    timeout=None,
    require_showing=True,
    require_valid_bounds=False,
):
    jab.ensure_started()
    if not path:
        raise ValueError("JAB text path is required")

    deadline = time.time() + (timeout or jab.search_timeout)
    while time.time() < deadline:
        check_abort()
        result = find_context_by_path_once(
            jab,
            path,
            title=title,
            class_name=class_name,
            scope_hwnd=scope_hwnd,
            name=name,
            role=role,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
        )
        context, vm_id, owned_contexts, _window_info = result
        if context:
            try:
                info = jab.get_context_info(vm_id, context)
                if not info:
                    return None
                return info.description.strip() or info.name.strip()
            finally:
                jab.release_contexts(vm_id, owned_contexts)
        time.sleep(0.2)
    return None


def set_text_context(jab, vm_id, context, text):
    if not hasattr(jab.dll, "setTextContents"):
        log.warning("当前 JAB DLL 不支持 setTextContents，拒绝使用全局键盘输入")
        return False

    if hasattr(jab.dll, "requestFocus") and not jab.dll.requestFocus(vm_id, context):
        log.warning("JAB requestFocus 失败，拒绝写入文本")
        return False

    value = str(text)
    ok = jab.dll.setTextContents(vm_id, context, value)
    if not ok:
        log.warning(f"JAB setTextContents 失败: text={value!r}")
        return False
    return True


def get_text_context_value(jab, vm_id, context):
    if not (
        hasattr(jab.dll, "getAccessibleTextInfo")
        and hasattr(jab.dll, "getAccessibleTextRange")
    ):
        return None

    text_info = AccessibleTextInfo()
    if not jab.dll.getAccessibleTextInfo(
        vm_id,
        context,
        ctypes.byref(text_info),
        0,
        0,
    ):
        return None

    if text_info.charCount <= 0:
        return ""

    buffer_len = min(text_info.charCount + 1, 4096)
    buffer = ctypes.create_unicode_buffer(buffer_len)
    if not jab.dll.getAccessibleTextRange(
        vm_id,
        context,
        0,
        min(text_info.charCount, buffer_len - 1),
        buffer,
        buffer_len,
    ):
        return None

    return buffer.value


def trigger_action_by_path_async(
    jab,
    path,
    title=None,
    class_name=None,
    scope_hwnd=None,
    name=None,
    role=None,
    action_name=None,
    timeout=None,
    require_showing=True,
    require_valid_bounds=True,
    cleanup_blank_awt=False,
):
    jab.ensure_started()
    result = find_context_by_path_once(
        jab,
        path,
        title=title,
        class_name=class_name,
        name=name,
        role=role,
        require_showing=require_showing,
        require_valid_bounds=require_valid_bounds,
    )
    context, vm_id, owned_contexts, window_info = result
    if not context:
        log.warning(
            "JAB async path 未找到可执行控件: "
            f"path={path} title={title!r} class={class_name!r}"
        )
        return None

    def target():
        try:
            ok = jab.do_action(
                vm_id,
                context,
                action_name=action_name,
                cleanup_blank_awt=cleanup_blank_awt,
            )
            log.info(
                "JAB async path 动作返回: "
                f"path={path} action={action_name or '<default>'} ok={ok} "
                f"hwnd={window_info.get('hwnd')} title={window_info.get('title')!r}"
            )
        finally:
            jab.release_contexts(vm_id, owned_contexts)

    thread = threading.Thread(target=target, daemon=True)
    thread.start()
    thread.join(timeout or 0)
    return thread


def find_context_by_path_once(
    jab,
    path,
    title=None,
    class_name=None,
    scope_hwnd=None,
    name=None,
    role=None,
    require_showing=False,
    require_valid_bounds=False,
):
    jab.ensure_started()
    parts = jab.parse_context_path(path)
    windows = enum_windows(include_children=True)
    normalized_role = role.lower() if role else None

    for hwnd, window_title, window_class, pid, visible in windows:
        if scope_hwnd is not None and int(hwnd) != int(scope_hwnd):
            continue
        if title and window_title != title:
            continue
        if class_name and window_class != class_name:
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

        context = root_context.value
        owned_contexts = [context]
        for index in parts[1:]:
            child = jab.dll.getAccessibleChildFromContext(vm_id.value, context, index)
            if not child:
                jab.release_contexts(vm_id.value, owned_contexts)
                context = None
                break
            context = child
            owned_contexts.append(child)

        if not context:
            continue

        if require_showing or require_valid_bounds or name or normalized_role:
            info = jab.get_context_info(vm_id.value, context)
            if not info:
                jab.release_contexts(vm_id.value, owned_contexts)
                continue
            control_name = info.name.strip()
            desc = info.description.strip()
            control_role = (info.role_en_US.strip() or info.role.strip()).lower()
            if name and control_name != name and desc != name:
                jab.release_contexts(vm_id.value, owned_contexts)
                continue
            if normalized_role and control_role != normalized_role:
                jab.release_contexts(vm_id.value, owned_contexts)
                continue
            if require_showing and not jab.context_info_is_showing(info):
                jab.release_contexts(vm_id.value, owned_contexts)
                continue
            if require_valid_bounds and not jab.context_info_has_valid_bounds(info):
                jab.release_contexts(vm_id.value, owned_contexts)
                continue

        return (
            context,
            vm_id.value,
            owned_contexts,
            {
                "hwnd": int(hwnd),
                "title": window_title,
                "class": window_class,
                "pid": pid,
                "visible": visible,
            },
        )

    return None, None, [], {}


def wait_context_by_path(
    jab,
    path,
    title=None,
    class_name=None,
    name=None,
    role=None,
    require_showing=True,
    require_valid_bounds=True,
    timeout=None,
    scope_hwnd=None,
):
    jab.ensure_started()
    deadline = time.time() + (timeout or jab.search_timeout)
    while time.time() < deadline:
        check_abort()
        result = find_context_by_path_once(
            jab,
            path,
            title=title,
            class_name=class_name,
            scope_hwnd=scope_hwnd,
            name=name,
            role=role,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
        )
        context, vm_id, owned_contexts, window_info = result
        if context:
            jab.release_contexts(vm_id, owned_contexts)
            return window_info
        time.sleep(0.1)
    return None
