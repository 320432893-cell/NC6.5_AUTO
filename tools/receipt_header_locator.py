# 职责：收款单表头字段定位 finders——按动态 path 定位、按语义标签反推文本 path 定位、path 落空后的语义兜底定位
# 不做什么：不写值/不粘贴(委托 writer/finance_org),不算 path 模板规则(委托 paths),不做纯状态判定(委托 state),不读 Excel/不解析 CLI
# 允许依赖层：core/JAB(jab.find_context_by_path_once) + tools.receipt_header_paths(纯计算/常量);
#             被 monkeypatch 的协作者(find_header_label_context_with_window /
#             find_receipt_header_field_by_semantic_label)经 _trial 代理读 tools.receipt_self_made_flow
# 谁不应该 import：receipt_header_{paths,state} 不应 import 本模块(会成环);writer 经门面 re-export 本模块,不反向被本模块 import

import sys

from tools.receipt_header_paths import (
    HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT,
    HEADER_SCOPE_ANCHOR_LABEL,
    build_receipt_header_dynamic_label_path,
    build_receipt_header_dynamic_path,
    build_receipt_header_label_path_from_template,
    build_receipt_header_path_from_template,
    infer_header_text_path_from_label_path,
    receipt_header_dynamic_prefix,
)


class _TrialNamespace:
    # 按调用时从已加载的入口模块 tools.receipt_self_made_flow 取属性：
    # 让测试对 trial 上 find_header_label_context_with_window /
    # find_receipt_header_field_by_semantic_label 的 monkeypatch 与拆分前一致地生效；
    # 不在加载期 import 入口模块以避免成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_self_made_flow"], name)


_trial = _TrialNamespace()


def find_receipt_header_field_by_live_semantic(
    jab,
    label,
    scope_hwnd=None,
    timeout=HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT,
):
    found = _trial.find_receipt_header_field_by_semantic_label(
        jab,
        label,
        scope_hwnd=scope_hwnd,
        timeout=timeout,
    )
    if found.get("ok"):
        found["source"] = "semantic-live-after-path-miss"
        return found
    return {
        **found,
        "source": "semantic-live-after-path-miss",
        "timeout": timeout,
    }


def find_receipt_header_field_by_semantic_label(
    jab,
    label,
    scope_hwnd=None,
    timeout=1.5,
):
    label_found = _trial.find_header_label_context_with_window(
        jab,
        label,
        timeout=timeout,
        require_showing=label != HEADER_SCOPE_ANCHOR_LABEL,
        scope_hwnd=scope_hwnd,
    )
    label_context, vm_id, owned_contexts, owned_indexes, window = label_found
    if not label_context:
        return {"ok": False, "label": label, "reason": "semantic label not found"}
    label_path = None
    if owned_indexes:
        label_path = "0" + "".join(f".{index}" for index in owned_indexes)
    jab.release_contexts(vm_id, owned_contexts)
    if not label_path:
        return {"ok": False, "label": label, "reason": "semantic label path missing"}
    text_path = infer_header_text_path_from_label_path(label, label_path)
    if not text_path:
        return {
            "ok": False,
            "label": label,
            "label_path": label_path,
            "reason": "semantic label path cannot infer text path",
        }
    label_window_hwnd = (window or {}).get("hwnd") or scope_hwnd
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        text_path,
        class_name="SunAwtCanvas",
        scope_hwnd=label_window_hwnd,
        role="text",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "label": label,
            "label_path": label_path,
            "path": text_path,
            "reason": "semantic inferred text path not found",
            "label_window": window,
            "path_scope_hwnd": label_window_hwnd,
        }
    return {
        "ok": True,
        "label": label,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "path": text_path,
        "label_path": label_path,
        "window": window_info,
    }


def find_receipt_header_field_by_dynamic_path(
    jab,
    label,
    dynamic_index,
    scope_hwnd=None,
    require_showing=True,
    require_valid_bounds=True,
    path_template=None,
):
    text_path = (
        build_receipt_header_path_from_template(dynamic_index, label, path_template)
        if path_template
        else build_receipt_header_dynamic_path(dynamic_index, label)
    )
    if not text_path:
        return {
            "ok": False,
            "reason": "header dynamic path not configured",
            "label": label,
            "dynamic_index": dynamic_index,
        }
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        text_path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="text",
        require_showing=require_showing,
        require_valid_bounds=require_valid_bounds,
    )
    if not context:
        return {
            "ok": False,
            "reason": "header dynamic path not found",
            "label": label,
            "path": text_path,
            "dynamic_index": dynamic_index,
        }
    return {
        "ok": True,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "path": text_path,
        "label_path": (
            build_receipt_header_label_path_from_template(
                dynamic_index,
                label,
                path_template,
            )
            if path_template
            else build_receipt_header_dynamic_label_path(dynamic_index, label)
        ),
        "window": window_info,
        "dynamic_index": dynamic_index,
        "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
        "path_template": path_template,
    }
