# 职责：收款单表头 scope/anchor 解析——定位财务组织锚点、校验锚点、按客户名纠正动态索引、语义推断 scope
# 不做什么：不写表头字段(委托 writer),不算 path 模板(委托 paths),不读 Excel/不解析 CLI
# 允许依赖层：tools.receipt_header_paths(纯计算);被 monkeypatch 的协作者(writer/tree 的字段定位
#             及本模块自身被 patch 的函数)经 _trial 代理读 tools.receipt_self_made_fill_trial
# 谁不应该 import：receipt_header_paths 不应 import 本模块(会成环)

import sys

from tools.receipt_header_paths import (
    HEADER_SCOPE_ANCHOR_LABEL,
    build_receipt_header_dynamic_label_path,
    extract_receipt_header_dynamic_index,
    header_scope_anchor_text_matches,
    receipt_header_dynamic_prefix,
)


class _TrialNamespace:
    # 按调用时从已加载的 tools.receipt_self_made_fill_trial 取属性：
    # 让测试对 trial 上 validate_receipt_header_scope_anchor /
    # correct_header_anchor_dynamic_index_by_customer / infer_receipt_header_scope_by_semantic /
    # find_header_label_context_with_window / find_receipt_header_field_by_dynamic_path /
    # find_receipt_header_field_by_semantic_label 的 monkeypatch 与拆分前一致地生效，
    # 同时打破 scope<->writer 的加载期循环依赖(不在加载期 import 入口模块)。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_self_made_fill_trial"], name)


_trial = _TrialNamespace()


def resolve_receipt_header_scope(
    jab, scope_hwnd=None, dynamic_index=None, anchor_path=None
):
    cached = getattr(jab, "_receipt_header_scope_cache", None)
    if cached and (scope_hwnd is None or cached.get("scope_hwnd") == scope_hwnd):
        return {**cached, "cached": True}
    if scope_hwnd and dynamic_index is not None:
        scoped = _trial.validate_receipt_header_scope_anchor(
            jab,
            scope_hwnd,
            dynamic_index,
            anchor_path=anchor_path,
        )
        if scoped.get("ok"):
            try:
                setattr(jab, "_receipt_header_scope_cache", scoped)
            except AttributeError:
                pass
            return scoped
        return scoped
    return {
        "ok": False,
        "mode": "provided-canvas-anchor",
        "reason": "正式表头缺少当前 canvas scope 或 dynamic_index，停止；不走语义兜底",
        "scope_hwnd": scope_hwnd,
        "dynamic_index": dynamic_index,
        "label_path": anchor_path,
    }


def validate_receipt_header_scope_anchor(
    jab, scope_hwnd, dynamic_index, anchor_path=None
):
    label_path = anchor_path or build_receipt_header_dynamic_label_path(
        dynamic_index,
        HEADER_SCOPE_ANCHOR_LABEL,
    )
    if not label_path:
        return {
            "ok": False,
            "mode": "provided-canvas-anchor",
            "reason": "财务组织(O) label path not configured",
            "scope_hwnd": scope_hwnd,
            "dynamic_index": dynamic_index,
        }
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        label_path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="label",
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "mode": "provided-canvas-anchor",
            "reason": "当前 canvas 未找到财务组织(O) 锚点",
            "scope_hwnd": scope_hwnd,
            "dynamic_index": dynamic_index,
            "label_path": label_path,
        }
    try:
        info = jab.get_context_info(vm_id, context)
        anchor_ok = bool(info and header_scope_anchor_text_matches(info))
        anchor_text = {
            "name": info.name.strip() if info else "",
            "description": info.description.strip() if info else "",
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)
    if not anchor_ok:
        return {
            "ok": False,
            "mode": "provided-canvas-anchor",
            "reason": "当前 canvas 财务组织(O) 锚点文本不匹配",
            "scope_hwnd": scope_hwnd,
            "dynamic_index": dynamic_index,
            "label_path": label_path,
            "anchor_text": anchor_text,
        }
    return {
        "ok": True,
        "scope_hwnd": scope_hwnd,
        "mode": "provided-canvas-anchor",
        "dynamic_index": dynamic_index,
        "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
        "matched_labels": [HEADER_SCOPE_ANCHOR_LABEL],
        "semantic_label_path": label_path,
        "anchor_text": anchor_text,
        "window": window_info,
    }


def resolve_receipt_header_anchor_in_canvas(jab, scope_hwnd, timeout=0.6):
    label_found = _trial.find_header_label_context_with_window(
        jab,
        HEADER_SCOPE_ANCHOR_LABEL,
        timeout=timeout,
        require_showing=False,
        scope_hwnd=scope_hwnd,
        strict_anchor=True,
    )
    label_context, vm_id, owned_contexts, owned_indexes, window = label_found
    if not label_context:
        return {
            "ok": False,
            "reason": "当前 canvas 未找到财务组织(O) 锚点",
            "scope_hwnd": scope_hwnd,
            "window": window,
        }
    label_path = ".".join(["0", *[str(index) for index in owned_indexes]])
    try:
        info = jab.get_context_info(vm_id, label_context)
        anchor_ok = bool(info and header_scope_anchor_text_matches(info))
        anchor_text = {
            "name": info.name.strip() if info else "",
            "description": info.description.strip() if info else "",
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)
    if not anchor_ok:
        return {
            "ok": False,
            "reason": "当前 canvas 财务组织(O) 锚点文本不匹配",
            "scope_hwnd": scope_hwnd,
            "label_path": label_path,
            "anchor_text": anchor_text,
            "window": window,
        }
    dynamic_index = extract_receipt_header_dynamic_index(label_path)
    if dynamic_index is None:
        return {
            "ok": False,
            "reason": "当前 canvas 财务组织(O) 锚点无法推出动态前缀",
            "scope_hwnd": scope_hwnd,
            "label_path": label_path,
            "anchor_text": anchor_text,
            "window": window,
        }
    corrected = _trial.correct_header_anchor_dynamic_index_by_customer(
        jab,
        scope_hwnd,
        dynamic_index,
    )
    if corrected.get("ok") and corrected.get("dynamic_index") != dynamic_index:
        corrected_index = corrected.get("dynamic_index")
        return {
            "ok": True,
            "scope_hwnd": scope_hwnd,
            "dynamic_index": corrected_index,
            "dynamic_prefix": receipt_header_dynamic_prefix(corrected_index),
            "label_path": label_path,
            "anchor_text": anchor_text,
            "window": corrected.get("window") or window,
            "mode": "current-canvas-anchor-corrected-by-customer",
            "initial_dynamic_index": dynamic_index,
            "initial_dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
            "correction": corrected,
        }
    return {
        "ok": True,
        "scope_hwnd": scope_hwnd,
        "dynamic_index": dynamic_index,
        "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
        "label_path": label_path,
        "anchor_text": anchor_text,
        "window": window,
        "mode": "current-canvas-anchor",
    }


def correct_header_anchor_dynamic_index_by_customer(jab, scope_hwnd, dynamic_index):
    current = _trial.find_receipt_header_field_by_dynamic_path(
        jab,
        "客户",
        dynamic_index,
        scope_hwnd=scope_hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    if current.get("ok"):
        jab.release_contexts(current["vm_id"], current["owned_contexts"])
        return {
            "ok": True,
            "source": "current-anchor-customer-path",
            "dynamic_index": dynamic_index,
            "path": current.get("path"),
            "window": current.get("window"),
        }
    semantic = _trial.find_receipt_header_field_by_semantic_label(
        jab,
        "客户",
        scope_hwnd=scope_hwnd,
    )
    if not semantic.get("ok"):
        return {
            "ok": False,
            "source": "customer-semantic-correction",
            "dynamic_index": dynamic_index,
            "current_attempt": current,
            "semantic_attempt": semantic,
        }
    corrected_index = extract_receipt_header_dynamic_index(semantic.get("path"))
    path = semantic.get("path")
    window = semantic.get("window")
    jab.release_contexts(semantic["vm_id"], semantic["owned_contexts"])
    if corrected_index is None:
        return {
            "ok": False,
            "source": "customer-semantic-correction",
            "dynamic_index": dynamic_index,
            "current_attempt": current,
            "semantic_path": path,
            "reason": "客户语义 path 无法推出 dynamic_index",
        }
    return {
        "ok": True,
        "source": "customer-semantic-correction",
        "dynamic_index": corrected_index,
        "path": path,
        "window": window,
        "current_attempt": current,
    }


def locate_receipt_header_scope(jab, scope_hwnd=None):
    semantic = _trial.infer_receipt_header_scope_by_semantic(jab, scope_hwnd=scope_hwnd)
    if semantic.get("ok"):
        return semantic
    return {
        "ok": False,
        "semantic_attempt": semantic,
        "reason": "未能用语义路径推断当前收款单表头 scope",
    }


def infer_receipt_header_scope_by_semantic(jab, scope_hwnd=None):
    found = _trial.find_receipt_header_field_by_semantic_label(
        jab,
        HEADER_SCOPE_ANCHOR_LABEL,
        scope_hwnd=scope_hwnd,
    )
    if not found.get("ok"):
        return {
            "ok": False,
            "mode": "semantic-path-inference",
            "reason": found.get("reason") or "财务组织语义定位失败",
            "attempt": found,
        }
    dynamic_index = extract_receipt_header_dynamic_index(found.get("path"))
    scope_hwnd = ((found.get("window") or {}).get("hwnd")) or None
    jab.release_contexts(found["vm_id"], found["owned_contexts"])
    if dynamic_index is None or not scope_hwnd:
        return {
            "ok": False,
            "mode": "semantic-path-inference",
            "reason": "财务组织语义路径无法推出动态前缀或窗口",
            "attempt": {
                "path": found.get("path"),
                "label_path": found.get("label_path"),
                "window": found.get("window"),
            },
        }
    return {
        "ok": True,
        "scope_hwnd": scope_hwnd,
        "mode": "semantic-path-inference",
        "dynamic_index": dynamic_index,
        "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
        "matched_labels": [HEADER_SCOPE_ANCHOR_LABEL],
        "semantic_label_path": found.get("label_path"),
        "semantic_text_path": found.get("path"),
    }
