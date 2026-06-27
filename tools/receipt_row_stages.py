# 职责：收款单单行流程里可独立切出的阶段步骤(表头 scope 解析、保存),供 run_one_row 编排调用
# 不做什么：不自建 JABOperator、不做整批编排、不读 Excel
# 允许依赖层：tools 收款定位缓存/表头组件;锁与模态恢复闭包由调用方传入
# 谁不应该 import：core 层模块不应 import

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools.receipt_locator_cache import (  # noqa: E402
    cache_receipt_header_scope,
    extract_entry_anchor_path,
    extract_entry_dynamic_index,
    extract_entry_scope_hwnd,
)
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    HEADER_SCOPE_ANCHOR_LABEL,
    find_finance_org_header_scope_by_paths,
)


def resolve_entry_header_scope(
    jab,
    jab_lock,
    run_with_jab_lock,
    open_step,
    header_scope_cache,
    timings,
    row_report,
):
    """解析当前 canvas 的表头 scope_hwnd / dynamic_index / anchor_path。

    成功返回 {"ok": True, "scope_hwnd", "dynamic_index", "anchor_path"};
    失败返回 {"ok": False, "reason"}。期间把诊断写进 row_report。
    """
    entry_scope_hwnd = extract_entry_scope_hwnd(open_step)
    entry_dynamic_index = extract_entry_dynamic_index(open_step)
    entry_anchor_path = extract_entry_anchor_path(open_step)
    entry_dynamic_index_source = (
        "entry-state" if entry_dynamic_index is not None else None
    )
    if entry_scope_hwnd and entry_dynamic_index is None:
        preferred_cached_dynamic_index = None
        cached_scope = {}
        if isinstance(header_scope_cache, dict):
            cached_scope = header_scope_cache
        if not cached_scope.get("ok"):
            cached_scope = getattr(jab, "_receipt_header_scope_cache", None) or {}
        if cached_scope.get("ok") and cached_scope.get("dynamic_index") is not None:
            preferred_cached_dynamic_index = cached_scope.get("dynamic_index")
        if (
            cached_scope.get("ok")
            and cached_scope.get("scope_hwnd") == entry_scope_hwnd
            and cached_scope.get("dynamic_index") is not None
        ):
            entry_dynamic_index = cached_scope.get("dynamic_index")
            entry_dynamic_index_source = "header-scope-cache"
            entry_anchor_path = (
                cached_scope.get("label_path")
                or cached_scope.get("semantic_label_path")
                or entry_anchor_path
            )
            row_report["entry_header_scope_cache"] = {
                "ok": True,
                "scope_hwnd": entry_scope_hwnd,
                "dynamic_index": entry_dynamic_index,
                "dynamic_prefix": cached_scope.get("dynamic_prefix"),
                "label_path": entry_anchor_path,
                "source": cached_scope.get("mode") or "receipt-header-scope-cache",
            }
        else:
            finance_scope = timings.measure(
                "header.finance-org-fast-scope",
                run_with_jab_lock,
                jab_lock,
                find_finance_org_header_scope_by_paths,
                jab,
                entry_scope_hwnd,
                preferred_dynamic_index=preferred_cached_dynamic_index,
                min_index=1,
                max_index=10,
            )
            row_report["entry_finance_org_fast_scope"] = {
                **finance_scope,
                "purpose": (
                    "开单快速确认只提供当前 Canvas；优先用财务组织(O)"
                    "稳定 path 解析表头 dynamic_index"
                )
            }
            if finance_scope.get("ok"):
                entry_dynamic_index = finance_scope.get("dynamic_index")
                entry_dynamic_index_source = "finance-org-fast-scope"
                finance_semantic_label_path = finance_scope.get("semantic_label_path")
                finance_label_path = finance_scope.get("label_path")
                entry_anchor_path = (
                    finance_label_path
                    or finance_semantic_label_path
                    or entry_anchor_path
                )
                if entry_dynamic_index is not None:
                    cache_receipt_header_scope(
                        jab,
                        header_scope_cache,
                        {
                            "ok": True,
                            "scope_hwnd": entry_scope_hwnd,
                            "mode": "finance-org-fast-scope",
                            "dynamic_index": entry_dynamic_index,
                            "dynamic_prefix": finance_scope.get("dynamic_prefix"),
                            "matched_labels": [HEADER_SCOPE_ANCHOR_LABEL],
                            "semantic_label_path": (
                                finance_semantic_label_path
                                or finance_label_path
                                or entry_anchor_path
                            ),
                            "label_path": (
                                finance_label_path
                                or finance_semantic_label_path
                                or entry_anchor_path
                            ),
                            "text_path": finance_scope.get("text_path"),
                            "variant": finance_scope.get("variant"),
                        },
                    )
    row_report["entry_scope_hwnd"] = entry_scope_hwnd
    row_report["entry_dynamic_index"] = entry_dynamic_index
    row_report["entry_dynamic_index_source"] = entry_dynamic_index_source
    row_report["entry_anchor_path"] = entry_anchor_path
    row_report["locator_policy"] = {
        "header": (
            "财务组织用于确认当前 canvas/scope 并缓存语义锚点；其它表头字段优先复用"
            "该 scope 做容器内标签定位，失败才单字段语义兜底"
        ),
        "body": "明细表优先复用已定位表格，必要时按表格语义扫描重新定位",
    }
    if not entry_scope_hwnd or entry_dynamic_index is None:
        return {
            "ok": False,
            "reason": "当前 canvas 未解析到财务组织(O) 前缀，停止；不走语义兜底",
        }
    return {
        "ok": True,
        "scope_hwnd": entry_scope_hwnd,
        "dynamic_index": entry_dynamic_index,
        "anchor_path": entry_anchor_path,
    }
