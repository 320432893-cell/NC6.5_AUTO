# 职责：承载 JABOperator 的 context path 兼容方法代理
# 不做什么：不实现 path 查找算法，不读取表格，不发送全局键盘输入
# 允许依赖层：core.jab_path_ops、core.jab_helpers
# 谁不应该 import：业务匹配、Sheet 写入、CLI 模块不应直接 import

from core import jab_path_ops
from core.jab_helpers import parse_context_path


class JABPathMixin:
    def do_action_by_path(
        self,
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
        return jab_path_ops.do_action_by_path(
            self,
            path,
            title=title,
            class_name=class_name,
            scope_hwnd=scope_hwnd,
            name=name,
            role=role,
            action_name=action_name,
            click_mode=click_mode,
            wait=wait,
            timeout=timeout,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
            cleanup_blank_awt=cleanup_blank_awt,
        )

    def set_text_by_path(
        self,
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
        return jab_path_ops.set_text_by_path(
            self,
            path,
            text,
            title=title,
            class_name=class_name,
            scope_hwnd=scope_hwnd,
            name=name,
            role=role,
            guard_path=guard_path,
            guard_name=guard_name,
            guard_role=guard_role,
            wait=wait,
            timeout=timeout,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
        )

    def get_text_by_path(
        self,
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
        return jab_path_ops.get_text_by_path(
            self,
            path,
            title=title,
            class_name=class_name,
            scope_hwnd=scope_hwnd,
            name=name,
            role=role,
            timeout=timeout,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
        )

    def set_text_context(self, vm_id, context, text):
        return jab_path_ops.set_text_context(self, vm_id, context, text)

    def get_text_context_value(self, vm_id, context):
        return jab_path_ops.get_text_context_value(self, vm_id, context)

    def trigger_action_by_path_async(
        self,
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
        return jab_path_ops.trigger_action_by_path_async(
            self,
            path,
            title=title,
            class_name=class_name,
            scope_hwnd=scope_hwnd,
            name=name,
            role=role,
            action_name=action_name,
            timeout=timeout,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
            cleanup_blank_awt=cleanup_blank_awt,
        )

    def find_context_by_path_once(
        self,
        path,
        title=None,
        class_name=None,
        scope_hwnd=None,
        name=None,
        role=None,
        require_showing=False,
        require_valid_bounds=False,
    ):
        return jab_path_ops.find_context_by_path_once(
            self,
            path,
            title=title,
            class_name=class_name,
            scope_hwnd=scope_hwnd,
            name=name,
            role=role,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
        )

    def wait_context_by_path(
        self,
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
        return jab_path_ops.wait_context_by_path(
            self,
            path,
            title=title,
            class_name=class_name,
            name=name,
            role=role,
            require_showing=require_showing,
            require_valid_bounds=require_valid_bounds,
            timeout=timeout,
            scope_hwnd=scope_hwnd,
        )

    @staticmethod
    def parse_context_path(path):
        return parse_context_path(path)
