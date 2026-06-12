# 职责：封装 JAB 控件查找、动作执行、context 树代理和保存按钮兼容入口
# 不做什么：不负责 JAB 生命周期加载，不发送全局键盘输入，不处理表格业务匹配
# 允许依赖层：标准库 ctypes/os/time、core.jab_context_tree、tools.jab_probe JAB 结构
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应直接 import

import ctypes
import os
from ctypes import wintypes
import time
from typing import TYPE_CHECKING, Any

from core import jab_context_tree
from core.jab_helpers import (
    context_info_has_valid_bounds,
    context_info_is_showing,
    normalize_amount,
    normalize_text,
    text_matches,
)
from core.logger import log
from core.utils import check_abort
from tools.jab_probe import (
    AccessibleActions,
    AccessibleActionsToDo,
    JOBJECT,
    enum_windows,
)


class JABControlMixin:
    if TYPE_CHECKING:
        dll: Any
        save_button_path: str | None
        save_button_title: str
        save_button_class: str
        search_timeout: float

        def ensure_started(self) -> None: ...

        def hide_blank_awt_windows(self): ...

        def do_action_by_path(self, *args, **kwargs): ...

    def click_control(
        self,
        name,
        roles=(),
        timeout=None,
        action_name=None,
        require_showing=False,
        window_title=None,
        window_class=None,
        visible_only=True,
        scope_hwnd=None,
    ):
        self.ensure_started()
        context, vm_id, owned_contexts = self.find_context(
            name,
            roles=roles,
            timeout=timeout,
            require_showing=require_showing,
            window_title=window_title,
            window_class=window_class,
            visible_only=visible_only,
            scope_hwnd=scope_hwnd,
        )
        if not context:
            log.warning(f"JAB 未找到控件: {name}")
            return False

        try:
            return self.do_action(vm_id, context, action_name=action_name)
        finally:
            self.release_contexts(vm_id, owned_contexts)

    def click_save(self, timeout=None):
        self.ensure_started()
        if self.save_button_path:
            with_path = self.do_action_by_path(
                self.save_button_path,
                title=self.save_button_title,
                class_name=self.save_button_class,
                name="保存(Ctrl+S)",
                role="push button",
                timeout=timeout or self.search_timeout,
                wait=0,
            )
            if with_path:
                return True
            log.warning(
                "JAB 保存按钮 path 快路径失败，回退按名称查找: "
                f"path={self.save_button_path}"
            )

        ok, path = self.click_control_with_path(
            "保存(Ctrl+S)",
            roles=("push button",),
            timeout=timeout or self.search_timeout,
            window_title=self.save_button_title,
            window_class=self.save_button_class,
        )
        if ok and path:
            log.info(f"JAB 保存按钮候选 path: {path}")
            self.save_button_path = path
        return ok

    def wait_for_control(self, name, roles=(), timeout=None, require_showing=False):
        self.ensure_started()
        context, vm_id, owned_contexts = self.find_context(
            name,
            roles=roles,
            timeout=timeout,
            require_showing=require_showing,
        )
        if not context:
            return False
        self.release_contexts(vm_id, owned_contexts)
        return True

    def wait_save_success(self, timeout=None):
        return self.wait_for_control("保存成功", timeout=timeout or self.search_timeout)

    def normalize_amount(self, value):
        return normalize_amount(value)

    def normalize_text(self, value):
        return normalize_text(value)

    def text_matches(self, value, target, match_mode):
        return text_matches(value, target, match_mode)

    def find_context(
        self,
        name,
        roles=(),
        timeout=None,
        require_showing=False,
        window_title=None,
        window_class=None,
        visible_only=True,
        scope_hwnd=None,
    ):
        deadline = time.time() + (timeout or self.search_timeout)
        normalized_roles = {role.lower() for role in roles}

        while time.time() < deadline:
            check_abort()
            result = self.find_context_once(
                name,
                normalized_roles,
                require_showing=require_showing,
                window_title=window_title,
                window_class=window_class,
                visible_only=visible_only,
                scope_hwnd=scope_hwnd,
            )
            if result[0]:
                return result
            time.sleep(0.2)

        return None, None, []

    def click_control_with_path(
        self,
        name,
        roles=(),
        timeout=None,
        action_name=None,
        require_showing=False,
        window_title=None,
        window_class=None,
        visible_only=True,
        scope_hwnd=None,
    ):
        self.ensure_started()
        context, vm_id, owned_contexts, owned_indexes = self.find_context_with_path(
            name,
            roles=roles,
            timeout=timeout,
            require_showing=require_showing,
            window_title=window_title,
            window_class=window_class,
            visible_only=visible_only,
            scope_hwnd=scope_hwnd,
        )
        if not context:
            log.warning(f"JAB 未找到控件: {name}")
            return False, None

        try:
            return self.do_action(vm_id, context, action_name=action_name), (
                "0" + "".join(f".{index}" for index in owned_indexes)
            )
        finally:
            self.release_contexts(vm_id, owned_contexts)

    def find_context_with_path(
        self,
        name,
        roles=(),
        timeout=None,
        require_showing=False,
        window_title=None,
        window_class=None,
        visible_only=True,
        scope_hwnd=None,
    ):
        deadline = time.time() + (timeout or self.search_timeout)
        normalized_roles = {role.lower() for role in roles}

        while time.time() < deadline:
            check_abort()
            result = self.find_context_once_with_path(
                name,
                normalized_roles,
                require_showing=require_showing,
                window_title=window_title,
                window_class=window_class,
                visible_only=visible_only,
                scope_hwnd=scope_hwnd,
            )
            if result[0]:
                return result
            time.sleep(0.2)

        return None, None, [], []

    def find_context_once_with_path(
        self,
        name,
        normalized_roles,
        require_showing=False,
        window_title=None,
        window_class=None,
        visible_only=True,
        scope_hwnd=None,
    ):
        windows = self.get_scoped_windows(scope_hwnd, include_children=True)

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
            if not self.dll.isJavaWindow(hwnd):
                continue

            vm_id = ctypes.c_long()
            root_context = JOBJECT()
            if not self.dll.getAccessibleContextFromHWND(
                hwnd,
                ctypes.byref(vm_id),
                ctypes.byref(root_context),
            ):
                continue

            context, owned_contexts, owned_indexes = self.find_in_tree_with_path(
                vm_id.value,
                root_context.value,
                name,
                normalized_roles,
                require_showing,
                depth=0,
                owned_contexts=[],
                owned_indexes=[],
            )
            if context:
                log.debug(
                    f"JAB 找到控件 {name}: hwnd={int(hwnd)} pid={pid} "
                    f"class={class_name!r} title={title!r} visible={visible}"
                )
                return context, vm_id.value, owned_contexts, owned_indexes

        return None, None, [], []

    def find_context_once(
        self,
        name,
        normalized_roles,
        require_showing=False,
        window_title=None,
        window_class=None,
        visible_only=True,
        scope_hwnd=None,
    ):
        windows = self.get_scoped_windows(scope_hwnd, include_children=True)

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
            if not self.dll.isJavaWindow(hwnd):
                continue

            vm_id = ctypes.c_long()
            root_context = JOBJECT()
            if not self.dll.getAccessibleContextFromHWND(
                hwnd,
                ctypes.byref(vm_id),
                ctypes.byref(root_context),
            ):
                continue

            context, owned_contexts = self.find_in_tree(
                vm_id.value,
                root_context.value,
                name,
                normalized_roles,
                require_showing,
                depth=0,
                owned_path=[],
            )
            if context:
                log.debug(
                    f"JAB 找到控件 {name}: hwnd={int(hwnd)} pid={pid} "
                    f"class={class_name!r} title={title!r} visible={visible}"
                )
                return context, vm_id.value, owned_contexts

        return None, None, []

    def get_scoped_windows(self, scope_hwnd=None, include_children=True):
        windows = enum_windows(include_children=include_children)
        if scope_hwnd is None or os.name != "nt":
            return windows

        user32 = ctypes.windll.user32
        scope_pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(
            wintypes.HWND(scope_hwnd), ctypes.byref(scope_pid)
        )
        scope_pid_value = int(scope_pid.value)
        if not scope_pid_value:
            return windows

        child_hwnds = set()
        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def child_callback(hwnd, _lparam):
            child_hwnds.add(int(hwnd))
            return True

        user32.EnumChildWindows(wintypes.HWND(scope_hwnd), enum_proc(child_callback), 0)

        scoped = []
        for hwnd, title, class_name, pid, visible in windows:
            hwnd_value = int(hwnd)
            if hwnd_value == int(scope_hwnd) or hwnd_value in child_hwnds:
                scoped.append((hwnd, title, class_name, pid, visible))
                continue
            if pid == scope_pid_value and class_name in (
                "SunAwtCanvas",
                "SunAwtFrame",
                "SunAwtDialog",
                "YonyouUWnd",
            ):
                scoped.append((hwnd, title, class_name, pid, visible))

        return scoped

    def find_in_tree(
        self, vm_id, context, name, normalized_roles, require_showing, depth, owned_path
    ):
        return jab_context_tree.find_in_tree(
            self,
            vm_id,
            context,
            name,
            normalized_roles,
            require_showing,
            depth,
            owned_path,
        )

    def find_in_tree_with_path(
        self,
        vm_id,
        context,
        name,
        normalized_roles,
        require_showing,
        depth,
        owned_contexts,
        owned_indexes,
    ):
        return jab_context_tree.find_in_tree_with_path(
            self,
            vm_id,
            context,
            name,
            normalized_roles,
            require_showing,
            depth,
            owned_contexts,
            owned_indexes,
        )

    def get_context_info(self, vm_id, context):
        return jab_context_tree.get_context_info(self, vm_id, context)

    def matches_control(
        self,
        control_name,
        desc,
        role,
        states,
        expected_name,
        normalized_roles,
        require_showing,
    ):
        return jab_context_tree.matches_control(
            control_name,
            desc,
            role,
            states,
            expected_name,
            normalized_roles,
            require_showing,
        )

    @staticmethod
    def context_info_is_showing(info):
        return context_info_is_showing(info)

    @staticmethod
    def context_info_has_valid_bounds(info):
        return context_info_has_valid_bounds(info)

    def get_action_names(self, vm_id, context):
        actions = AccessibleActions()
        if not self.dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
            return []
        return [
            actions.actionInfo[index].name.strip()
            for index in range(actions.actionsCount)
        ]

    def do_action(self, vm_id, context, action_name=None, cleanup_blank_awt=False):
        if not hasattr(self.dll, "getAccessibleActions") or not hasattr(
            self.dll, "doAccessibleActions"
        ):
            log.warning("当前 JAB DLL 不支持 AccessibleActions")
            return False

        action_names = self.get_action_names(vm_id, context)
        if not action_names:
            log.warning("JAB 控件没有可执行动作")
            return False

        chosen_action = action_name or action_names[0]
        if chosen_action not in action_names:
            log.warning(
                f"JAB 控件不支持动作 {chosen_action!r}，可用动作: {action_names}"
            )
            return False

        todo = AccessibleActionsToDo()
        todo.actionsCount = 1
        todo.actions[0].name = chosen_action
        failure = ctypes.c_int(-1)
        ok = self.dll.doAccessibleActions(
            vm_id,
            context,
            ctypes.byref(todo),
            ctypes.byref(failure),
        )
        if ok:
            time.sleep(0.2)
            if cleanup_blank_awt:
                self.hide_blank_awt_windows()
        log.debug(
            f"JAB 执行动作 {chosen_action!r}: ok={bool(ok)} failure={failure.value}"
        )
        return bool(ok)

    def release_contexts(self, vm_id, contexts):
        jab_context_tree.release_contexts(self, vm_id, contexts)
