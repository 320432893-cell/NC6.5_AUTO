# 职责：承载 JABOperator 的 near-label 兼容方法代理
# 不做什么：不实现 near-label 搜索算法，不处理表格读写，不发送全局键盘输入
# 允许依赖层：core.jab_near_label、core.logger
# 谁不应该 import：收款匹配、Excel/Sheet 读写、CLI 模块不应直接 import

from core import jab_near_label
from typing import TYPE_CHECKING


class JABNearLabelMixin:
    if TYPE_CHECKING:

        def get_context_info(self, vm_id, context): ...

    def set_text_near_label(
        self,
        label,
        text,
        title=None,
        class_name=None,
        hwnd=None,
        wait=None,
        timeout=None,
        require_showing=True,
    ):
        return jab_near_label.set_text_near_label(
            self,
            label,
            text,
            title=title,
            class_name=class_name,
            hwnd=hwnd,
            wait=wait,
            timeout=timeout,
            require_showing=require_showing,
        )

    def describe_controls_near_label(
        self,
        label,
        title=None,
        class_name=None,
        hwnd=None,
        require_showing=True,
        max_vertical_distance=28,
        max_right_distance=420,
    ):
        return jab_near_label.describe_controls_near_label(
            self,
            label,
            title=title,
            class_name=class_name,
            hwnd=hwnd,
            require_showing=require_showing,
            max_vertical_distance=max_vertical_distance,
            max_right_distance=max_right_distance,
        )

    def click_control_near_label(
        self,
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
        return jab_near_label.click_control_near_label(
            self,
            label,
            role,
            index=index,
            title=title,
            class_name=class_name,
            require_showing=require_showing,
            max_vertical_distance=max_vertical_distance,
            max_right_distance=max_right_distance,
            action_name=action_name,
            wait=wait,
        )

    def describe_text_near_label(
        self,
        label,
        title=None,
        class_name=None,
        require_showing=True,
    ):
        return jab_near_label.describe_text_near_label(
            self,
            label,
            title=title,
            class_name=class_name,
            require_showing=require_showing,
        )

    @staticmethod
    def info_to_dict(info):
        return jab_near_label.info_to_dict(info)
