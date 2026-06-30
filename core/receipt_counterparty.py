# 职责：收款单往来对象(客户)字段的定位/读取/写入/校验全部 JAB 逻辑
# 不做什么：不做整单编排、不保存、不读 Excel
# 允许依赖层：core JAB(jab_probe)、core 收款表头/定位缓存(receipt_self_made_fill_trial/receipt_locator_cache);谁不应 import：tools 入口层(不反依赖 receipt_full_flow_entry)

import ctypes
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_probe import JOBJECT  # noqa: E402
from core.receipt_self_made_fill_trial import (  # noqa: E402
    receipt_header_dynamic_prefix,
)


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.receipt_locator_cache import (  # noqa: E402
    resolve_body_table_by_dynamic_prefix,
)

COUNTERPARTY_LABEL = "往来对象"

COUNTERPARTY_EXPECTED = "客户"

COUNTERPARTY_KNOWN_OPTIONS = {"客户", "部门", "业务员", "供应商"}

COUNTERPARTY_NEARBY_MAX_VERTICAL_DISTANCE = 36

COUNTERPARTY_NEARBY_MAX_RIGHT_DISTANCE = 700

COUNTERPARTY_STATE_OK = "ok"

COUNTERPARTY_STATE_REPAIRABLE = "repairable-empty"

COUNTERPARTY_STATE_REPAIRABLE_CONFLICT = "repairable-conflict"

COUNTERPARTY_STATE_DETAIL_UNREADABLE = "detail-unreadable"

_COUNTERPARTY_NEARBY_SUFFIX_CACHE = {}


def ensure_header_counterparty_customer(
    jab,
    dynamic_index,
    scope_hwnd=None,
    located=None,
    recover_after_failure=None,
):
    started_at = time.perf_counter()
    if dynamic_index is None:
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "dynamic_index": dynamic_index,
            "reason": "往来对象 dynamic_index 未配置",
            "seconds": round(time.perf_counter() - started_at, 3),
        }

    detail = read_detail_counterparty_value(
        jab,
        dynamic_index,
        scope_hwnd=scope_hwnd,
        located=located,
        row=0,
        col=0,
    )
    detail_value = normalize_counterparty_value(
        detail.get("value"),
        detail.get("text"),
    )
    if detail_value == COUNTERPARTY_EXPECTED:
        return {
            "ok": True,
            "skipped": True,
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "actual": COUNTERPARTY_EXPECTED,
            "path": None,
            "dynamic_index": dynamic_index,
            "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
            "detail": detail,
            "state": {
                "state": COUNTERPARTY_STATE_OK,
                "actual": detail_value,
                "source": "detail-row0-col0",
                "repairable": False,
            },
            "source": "detail-row0-col0",
            "reason": "明细表第 0 行往来对象为客户，跳过",
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    if detail.get("ok") is False and not detail_value:
        reason = str(detail.get("reason") or "")
        unreadable = any(
            marker in reason for marker in ("未定位", "命中失败", "行列不足", "异常")
        )
        if unreadable:
            snapshot = {
                "combo": {},
                "embedded": {},
                "detail": detail,
                "selected": "",
                "combo_text": "",
                "detail_value": "",
                "state": {
                    "state": COUNTERPARTY_STATE_DETAIL_UNREADABLE,
                    "actual": "",
                    "source": "detail-row0-col0",
                    "repairable": False,
                },
            }
            return {
                "ok": False,
                "label": COUNTERPARTY_LABEL,
                "expected": COUNTERPARTY_EXPECTED,
                "actual": "",
                "path": None,
                "dynamic_index": dynamic_index,
                "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
                "before": {},
                "embedded": {},
                "detail": detail,
                "state": snapshot["state"],
                "readback_trusted": False,
                "reason": summarize_counterparty_failure(snapshot),
                "seconds": round(time.perf_counter() - started_at, 3),
            }

    found = find_counterparty_combo(jab, dynamic_index, scope_hwnd=scope_hwnd)
    found_path = found.get("path")
    recovery_after_find = None
    if not found.get("ok") and recover_after_failure is not None:
        recovery_after_find = recover_after_failure()
        if recovery_after_find.get("attempted") and recovery_after_find.get("ok"):
            found = find_counterparty_combo(jab, dynamic_index, scope_hwnd=scope_hwnd)
            found_path = found.get("path")
    if not found.get("ok"):
        return {
            **found,
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "actual": detail_value,
            "dynamic_index": dynamic_index,
            "detail": detail,
            "modal_recovery": recovery_after_find,
            "seconds": round(time.perf_counter() - started_at, 3),
        }

    try:
        combo = read_counterparty_combo_state(
            jab,
            found["vm_id"],
            found["context"],
        )
        embedded = read_counterparty_selected_option(
            jab,
            found["vm_id"],
            found["context"],
        )
        snapshot = {
            "combo": combo,
            "embedded": embedded,
            "detail": detail,
            "selected": normalize_counterparty_value(embedded.get("selected")),
            "combo_text": normalize_counterparty_value(
                combo.get("description"),
                combo.get("text"),
                combo.get("name"),
            ),
            "detail_value": detail_value,
        }
        state = classify_counterparty_snapshot(snapshot)
        snapshot["state"] = state
        if state["state"] == COUNTERPARTY_STATE_OK:
            return {
                "ok": True,
                "skipped": True,
                "label": COUNTERPARTY_LABEL,
                "expected": COUNTERPARTY_EXPECTED,
                "actual": COUNTERPARTY_EXPECTED,
                "path": found_path,
                "dynamic_index": dynamic_index,
                "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
                "before": snapshot["combo"],
                "embedded": snapshot["embedded"],
                "detail": snapshot["detail"],
                "state": state,
                "source": "detail-row0-col0",
                "reason": "明细表第 0 行往来对象为客户，跳过",
                "seconds": round(time.perf_counter() - started_at, 3),
            }

        if state["state"] in {
            COUNTERPARTY_STATE_REPAIRABLE,
            COUNTERPARTY_STATE_REPAIRABLE_CONFLICT,
        }:
            return repair_counterparty_to_customer(
                jab,
                found,
                found_path,
                snapshot,
                state,
                dynamic_index,
                scope_hwnd,
                located,
                started_at,
            )

        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "actual": state.get("actual") or "",
            "path": found_path,
            "dynamic_index": dynamic_index,
            "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
            "before": snapshot["combo"],
            "embedded": snapshot["embedded"],
            "detail": snapshot["detail"],
            "state": state,
            "readback_trusted": False,
            "reason": summarize_counterparty_failure(snapshot),
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    finally:
        jab.release_contexts(found["vm_id"], found["owned_contexts"])


def classify_counterparty_snapshot(snapshot):
    selected = (snapshot or {}).get("selected") or ""
    combo_text = (snapshot or {}).get("combo_text") or ""
    detail_value = (snapshot or {}).get("detail_value") or ""
    detail = (snapshot or {}).get("detail") or {}

    if detail_value == COUNTERPARTY_EXPECTED:
        return {
            "state": COUNTERPARTY_STATE_OK,
            "actual": detail_value,
            "source": "detail-row0-col0",
            "repairable": False,
        }

    for source, value in (
        ("detail-row0-col0", detail_value),
        ("combo-text", combo_text),
        ("embedded-selected-option", selected),
    ):
        if value in COUNTERPARTY_KNOWN_OPTIONS and value != COUNTERPARTY_EXPECTED:
            return {
                "state": COUNTERPARTY_STATE_REPAIRABLE_CONFLICT,
                "actual": value,
                "source": source,
                "repairable": True,
            }

    if detail.get("ok") is False and not detail_value:
        reason = str(detail.get("reason") or "")
        unreadable = any(
            marker in reason for marker in ("未定位", "命中失败", "行列不足", "异常")
        )
        if unreadable:
            return {
                "state": COUNTERPARTY_STATE_DETAIL_UNREADABLE,
                "actual": "",
                "source": "detail-row0-col0",
                "repairable": False,
            }

    return {
        "state": COUNTERPARTY_STATE_REPAIRABLE,
        "actual": detail_value,
        "source": "detail-row0-col0",
        "repairable": True,
    }


def repair_counterparty_to_customer(
    jab,
    found,
    found_path,
    snapshot,
    state,
    dynamic_index,
    scope_hwnd,
    located,
    started_at,
):
    repair = select_counterparty_customer_embedded(
        jab,
        found["vm_id"],
        found["context"],
        press_enter=True,
    )
    time.sleep(0.12)
    after_detail = read_detail_counterparty_value(
        jab,
        dynamic_index,
        scope_hwnd=scope_hwnd,
        located=located,
        row=0,
        col=0,
    )
    after_value = normalize_counterparty_value(
        after_detail.get("value"),
        after_detail.get("text"),
    )
    if after_value == COUNTERPARTY_EXPECTED:
        from_conflict = state.get("state") == COUNTERPARTY_STATE_REPAIRABLE_CONFLICT
        return {
            "ok": True,
            "repaired": True,
            "repaired_from_conflict": bool(from_conflict),
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "actual": after_value,
            "path": found_path,
            "dynamic_index": dynamic_index,
            "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
            "before": snapshot["combo"],
            "embedded": snapshot["embedded"],
            "detail": snapshot["detail"],
            "after_detail": after_detail,
            "repair": repair,
            "state": state,
            "source": "embedded-selection-api",
            "reason": (
                "往来对象为非客户选项，已通过子列表 selection API 选择客户并验证明细表"
                if from_conflict
                else "往来对象为空，已通过子列表 selection API 选择客户并验证明细表"
            ),
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    return {
        "ok": False,
        "label": COUNTERPARTY_LABEL,
        "expected": COUNTERPARTY_EXPECTED,
        "actual": after_value,
        "path": found_path,
        "dynamic_index": dynamic_index,
        "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
        "before": snapshot["combo"],
        "embedded": snapshot["embedded"],
        "detail": snapshot["detail"],
        "after_detail": after_detail,
        "repair": repair,
        "state": state,
        "readback_trusted": False,
        "reason": summarize_counterparty_failure(snapshot, after_detail),
        "seconds": round(time.perf_counter() - started_at, 3),
    }


def normalize_counterparty_value(*values):
    for value in values:
        text = str(value or "").strip()
        if text in COUNTERPARTY_KNOWN_OPTIONS:
            return text
    return ""


def select_counterparty_customer_embedded(jab, vm_id, combo_context, press_enter=True):
    target = find_counterparty_embedded_list(jab, vm_id, combo_context)
    try:
        if not target.get("ok"):
            return {
                "ok": False,
                "method": "embedded-selection-api",
                "reason": target.get("reason") or "往来对象子列表未找到",
            }
        if not hasattr(jab.dll, "addAccessibleSelectionFromContext"):
            return {
                "ok": False,
                "method": "embedded-selection-api",
                "reason": "JAB selection API unavailable",
                "target": embedded_counterparty_target_summary(target),
            }
        list_context = target["list_context"]
        customer_index = next(
            (
                int(item.get("index"))
                for item in target.get("labels") or []
                if item.get("name") == COUNTERPARTY_EXPECTED
            ),
            None,
        )
        if customer_index is None:
            return {
                "ok": False,
                "method": "embedded-selection-api",
                "reason": "往来对象子列表没有客户选项",
                "target": embedded_counterparty_target_summary(target),
            }
        if hasattr(jab.dll, "clearAccessibleSelectionFromContext"):
            jab.dll.clearAccessibleSelectionFromContext(vm_id, list_context)
        selected_ok = bool(
            jab.dll.addAccessibleSelectionFromContext(
                vm_id,
                list_context,
                customer_index,
            )
        )
        focus_ok = request_focus_context(jab, vm_id, list_context)
        enter_ok = None
        if press_enter:
            enter_ok = press_counterparty_commit_keys(jab)
        return {
            "ok": bool(selected_ok),
            "method": "embedded-selection-api",
            "selected_index": customer_index,
            "target": embedded_counterparty_target_summary(target),
            "request_focus_list": focus_ok,
            "commit": enter_ok,
        }
    except Exception as exc:
        return {
            "ok": False,
            "method": "embedded-selection-api",
            "error": repr(exc),
        }
    finally:
        if target.get("owned_contexts"):
            jab.release_contexts(vm_id, target["owned_contexts"])


def request_focus_context(jab, vm_id, context):
    if not hasattr(jab, "dll") or not hasattr(jab.dll, "requestFocus"):
        return {"ok": None, "reason": "requestFocus unavailable"}
    try:
        return {"ok": bool(jab.dll.requestFocus(vm_id, context))}
    except Exception as exc:
        return {"ok": False, "error": repr(exc)}


def press_counterparty_commit_keys(jab):
    sent = []
    try:
        jab.press_key("home", wait=0.02)
        sent.append("home")
        jab.press_key("enter", wait=0)
        sent.append("enter")
        return {"ok": True, "keys": sent}
    except Exception as exc:
        return {"ok": False, "keys": sent, "error": repr(exc)}


def embedded_counterparty_target_summary(target):
    return {
        "list": target.get("list"),
        "popup": target.get("popup"),
        "labels": [
            {
                "index": item.get("index"),
                "name": item.get("name"),
                "states": item.get("states"),
            }
            for item in target.get("labels") or []
        ],
    }


def first_non_empty_counterparty_text(*values):
    return normalize_counterparty_value(*values)


def read_detail_counterparty_value(
    jab,
    dynamic_index,
    scope_hwnd=None,
    row=0,
    col=0,
    located=None,
):
    try:
        if located is None:
            located = resolve_body_table_by_dynamic_prefix(
                jab,
                dynamic_index,
                scope_hwnd=scope_hwnd,
            )
        best = (located or {}).get("best") or {}
        path = best.get("path")
        if not path:
            return {
                "ok": False,
                "source": "detail-row0-col0",
                "row": row,
                "col": col,
                "located": slim_counterparty_located(located),
                "reason": "明细表 path 未定位",
            }

        window = best.get("window") or {}
        context, vm_id, owned, window_info = jab.find_context_by_path_once(
            path,
            class_name=window.get("class_name") or "SunAwtCanvas",
            scope_hwnd=scope_hwnd or window.get("hwnd"),
            role="table",
            require_showing=False,
            require_valid_bounds=False,
        )
        if not context:
            return {
                "ok": False,
                "source": "detail-row0-col0",
                "row": row,
                "col": col,
                "path": path,
                "located": slim_counterparty_located(located),
                "reason": "明细表 path 命中失败",
            }
        try:
            table_info = jab.get_table_info(vm_id, context)
            table = {
                "path": path,
                "window": window_info or window,
                "row_count": int(getattr(table_info, "rowCount", 0) or 0)
                if table_info
                else None,
                "col_count": int(getattr(table_info, "columnCount", 0) or 0)
                if table_info
                else None,
            }
            schema = detail_table_schema_snapshot(best)
            table["schema"] = schema
            if table.get("col_count") is not None and table.get("col_count") < 12:
                return {
                    "ok": False,
                    "source": "detail-row0-col0",
                    "row": row,
                    "col": col,
                    "path": path,
                    "table": table,
                    "located": slim_counterparty_located(located),
                    "reason": "明细表列数不足，不像收款单明细表",
                }
            if table_info and (
                int(table_info.rowCount) <= row or int(table_info.columnCount) <= col
            ):
                return {
                    "ok": False,
                    "source": "detail-row0-col0",
                    "row": row,
                    "col": col,
                    "path": path,
                    "table": table,
                    "located": slim_counterparty_located(located),
                    "reason": "明细表行列不足，无法读取往来对象",
                }
            text, is_selected = jab.get_table_cell_text_and_selection(
                vm_id,
                context,
                row,
                col,
            )
            value = first_non_empty_counterparty_text(text)
            return {
                "ok": bool(value),
                "source": "detail-row0-col0",
                "row": row,
                "col": col,
                "value": value,
                "text": str(text or "").strip(),
                "is_selected": bool(is_selected),
                "path": path,
                "table": table,
                "located": slim_counterparty_located(located),
                "reason": None if value else "明细表往来对象单元格为空",
            }
        finally:
            jab.release_contexts(vm_id, owned)
    except Exception as exc:
        return {
            "ok": False,
            "source": "detail-row0-col0",
            "row": row,
            "col": col,
            "reason": "读取明细表往来对象异常",
            "error": repr(exc),
        }


def slim_counterparty_located(located):
    if not located:
        return None
    best = (located or {}).get("best") or {}
    return {
        "cache_hit": bool(located.get("cache_hit")),
        "fallback_used": bool(located.get("fallback_used")),
        "source": located.get("source"),
        "path": best.get("path"),
        "window": best.get("window"),
        "row_count": best.get("row_count"),
        "col_count": best.get("col_count"),
        "reason": located.get("reason"),
    }


def detail_table_schema_snapshot(best):
    rows = (best or {}).get("rows") or []
    first = rows[0] if rows else {}
    cells = (first or {}).get("cells") or first.get("values") or first
    if not isinstance(cells, dict):
        cells = {}
    key_cells = {
        str(index): str(cells.get(str(index), cells.get(index, "")) or "").strip()
        for index in (0, 1, 2, 3, 4, 5, 7, 11)
    }
    return {
        "row0_key_cells": key_cells,
        "looks_like_receipt_detail": (
            key_cells.get("0") in {"", *COUNTERPARTY_KNOWN_OPTIONS}
            and bool(key_cells.get("1") or key_cells.get("2") or key_cells.get("5"))
        ),
    }


def summarize_counterparty_failure(snapshot, after_detail=None):
    snapshot = snapshot or {}
    detail = after_detail or snapshot.get("detail") or {}
    detail_value = normalize_counterparty_value(
        detail.get("value"),
        detail.get("text"),
    )
    raw_detail = (detail or {}).get("text") or ""
    detail_reason = (detail or {}).get("reason") or ""
    state = snapshot.get("state") or {}
    parts = [
        f"header_selected={snapshot.get('selected') or ''}",
        f"combo_text={snapshot.get('combo_text') or ''}",
        f"detail_row0_col0={detail_value}",
    ]
    if raw_detail and raw_detail != detail_value:
        parts.append(f"detail_raw={raw_detail}")
    if detail_reason:
        parts.append(f"detail_reason={detail_reason}")
    if state.get("state"):
        parts.append(f"state={state.get('state')}")
    return f"往来对象未确认客户；{'; '.join(parts)}；已禁用旧下拉键盘方案"


def read_counterparty_selected_option(jab, vm_id, combo_context):
    found = find_counterparty_embedded_list(jab, vm_id, combo_context)
    try:
        if not found.get("ok"):
            return found
        labels = found.get("labels") or []
        selected = next(
            (
                item.get("name")
                for item in labels
                if "selected" in str(item.get("states") or "").lower()
            ),
            "",
        )
        return {
            "ok": True,
            "selected": selected,
            "options": [item.get("name") for item in labels if item.get("name")],
            "list": found.get("list"),
            "popup": found.get("popup"),
        }
    finally:
        if found.get("owned_contexts"):
            jab.release_contexts(vm_id, found["owned_contexts"])


def find_counterparty_embedded_list(jab, vm_id, combo_context):
    result = {
        "ok": False,
        "reason": "往来对象子列表未找到",
        "owned_contexts": [],
    }
    best = None

    def visit(context, path, depth, ancestors):
        nonlocal best
        info = jab.get_context_info(vm_id, context)
        if not info:
            return
        role = (info.role_en_US.strip() or info.role.strip()).lower()
        owned = []
        children = []
        if depth > 0:
            for index in range(
                min(info.childrenCount, getattr(jab, "max_children", 1000))
            ):
                child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
                if not child:
                    continue
                owned.append(child)
                child_info = jab.get_context_info(vm_id, child)
                if not child_info:
                    continue
                child_role = (
                    child_info.role_en_US.strip() or child_info.role.strip()
                ).lower()
                child_name = child_info.name.strip()
                children.append((index, child, child_info, child_role, child_name))

        if role == "list":
            label_items = [
                {
                    "index": index,
                    "path": f"{path}.{index}",
                    "name": child_name,
                    "description": child_info.description.strip(),
                    "role": child_info.role_en_US.strip() or child_info.role.strip(),
                    "states": child_info.states_en_US.strip()
                    or child_info.states.strip(),
                    "bounds": [
                        child_info.x,
                        child_info.y,
                        child_info.width,
                        child_info.height,
                    ],
                }
                for index, _child, child_info, child_role, child_name in children
                if child_role == "label"
            ]
            names = [item["name"] for item in label_items if item.get("name")]
            if COUNTERPARTY_EXPECTED in names:
                popup = next(
                    (
                        info_to_counterparty_dict(ancestor_info, "ancestor")
                        for _ancestor_context, ancestor_info in reversed(ancestors)
                        if (
                            ancestor_info.role_en_US.strip()
                            or ancestor_info.role.strip()
                        ).lower()
                        == "popup menu"
                    ),
                    None,
                )
                keep = [context]
                keep.extend(child for _index, child, *_rest in children)
                best = {
                    "ok": True,
                    "list": info_to_counterparty_dict(info, path),
                    "list_context": context,
                    "popup": popup,
                    "labels": label_items,
                    "owned_contexts": unique_contexts(keep + owned),
                }
                return

        try:
            if depth > 0 and best is None:
                next_ancestors = ancestors + [(context, info)]
                for index, child, _child_info, _child_role, _child_name in children:
                    visit(child, f"{path}.{index}", depth - 1, next_ancestors)
                    if best is not None:
                        break
        finally:
            if best is None:
                jab.release_contexts(vm_id, owned)

    visit(combo_context, "target", 8, [])
    return best or result


def info_to_counterparty_dict(info, path):
    if not info:
        return None
    return {
        "path": path,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "role": info.role_en_US.strip() or info.role.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "bounds": [info.x, info.y, info.width, info.height],
        "children_count": info.childrenCount,
    }


def unique_contexts(contexts):
    result = []
    seen = set()
    for context in contexts or []:
        key = context_key(context)
        if key in seen:
            continue
        seen.add(key)
        result.append(context)
    return result


def context_key(context):
    try:
        value = getattr(context, "value", context)
        return ("int", int(value))
    except Exception:
        return ("repr", repr(context))


def find_counterparty_combo(jab, dynamic_index, scope_hwnd=None):
    prefix = receipt_header_dynamic_prefix(dynamic_index)
    cached_path = build_cached_counterparty_nearby_path(dynamic_index, scope_hwnd)
    if cached_path:
        cached = find_counterparty_combo_by_path(
            jab,
            cached_path,
            scope_hwnd=scope_hwnd,
        )
        if cached.get("ok"):
            return {
                **cached,
                "source": "nearby-cache-path",
                "cached_path": cached_path,
            }

    nearby = find_counterparty_combo_nearby(
        jab,
        dynamic_index,
        scope_hwnd=scope_hwnd,
    )
    if nearby.get("ok"):
        cache_counterparty_nearby_suffix(
            dynamic_index,
            scope_hwnd,
            prefix,
            nearby.get("path"),
        )
        return nearby

    return {
        "ok": False,
        "label": COUNTERPARTY_LABEL,
        "source": "not-found",
        "nearby_attempt": slim_found(nearby),
        "reason": nearby.get("reason") or "往来对象 nearby 定位失败",
    }


def build_cached_counterparty_nearby_path(dynamic_index, scope_hwnd=None):
    prefix = receipt_header_dynamic_prefix(dynamic_index)
    if not prefix:
        return None
    cached = _COUNTERPARTY_NEARBY_SUFFIX_CACHE.get(counterparty_cache_key(scope_hwnd))
    if not cached:
        cached = _COUNTERPARTY_NEARBY_SUFFIX_CACHE.get("last")
    suffix = (cached or {}).get("suffix")
    if not suffix:
        return None
    return f"{prefix}.{suffix}"


def cache_counterparty_nearby_suffix(dynamic_index, scope_hwnd, prefix, path):
    if not prefix or not path or not str(path).startswith(f"{prefix}."):
        return None
    suffix = str(path)[len(prefix) + 1 :]
    cached = {
        "dynamic_index": dynamic_index,
        "scope_hwnd": scope_hwnd,
        "suffix": suffix,
        "path": path,
        "source": "nearby",
    }
    _COUNTERPARTY_NEARBY_SUFFIX_CACHE[counterparty_cache_key(scope_hwnd)] = cached
    _COUNTERPARTY_NEARBY_SUFFIX_CACHE["last"] = cached
    return cached


def counterparty_cache_key(scope_hwnd):
    return int(scope_hwnd) if scope_hwnd is not None else None


def find_counterparty_combo_nearby(jab, dynamic_index, scope_hwnd=None):
    if scope_hwnd is None:
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "reason": "nearby 定位缺少 scope_hwnd",
        }
    dll = getattr(jab, "dll", None)
    if not dll or not hasattr(dll, "isJavaWindow") or not dll.isJavaWindow(scope_hwnd):
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "scope_hwnd": scope_hwnd,
            "reason": "nearby scope 不是 Java 窗口",
        }

    vm_id = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        int(scope_hwnd),
        ctypes.byref(vm_id),
        ctypes.byref(root_context),
    ):
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "scope_hwnd": scope_hwnd,
            "reason": "nearby 读取 scope root 失败",
        }

    controls = []
    owned = []
    selected_contexts = set()
    try:
        collect_counterparty_controls_for_bounds_scan(
            jab,
            vm_id.value,
            root_context.value,
            controls,
            owned,
            require_showing=True,
            depth=0,
        )
        labels = [
            (context, info, _path)
            for context, info, _path in controls
            if control_role(info) == "label"
            and info.name.strip() == COUNTERPARTY_LABEL
            and jab.context_info_has_valid_bounds(info)
        ]
        labels.sort(key=lambda item: (item[1].y, item[1].x))
        prefix = receipt_header_dynamic_prefix(dynamic_index)
        candidates = []
        for label_context, label_info, _label_path in labels:
            label_mid_y = label_info.y + label_info.height / 2
            label_right = label_info.x + label_info.width
            row_candidates = []
            for context, info, control_path in controls:
                if control_role(info) != "combo box":
                    continue
                if not jab.context_info_has_valid_bounds(info):
                    continue
                mid_y = info.y + info.height / 2
                right_distance = info.x - label_right
                dy = abs(mid_y - label_mid_y)
                if right_distance <= 0:
                    continue
                if dy > COUNTERPARTY_NEARBY_MAX_VERTICAL_DISTANCE:
                    continue
                if right_distance > COUNTERPARTY_NEARBY_MAX_RIGHT_DISTANCE:
                    continue
                actions = jab.get_action_names(vm_id.value, context)
                row_candidates.append(
                    {
                        "context": context,
                        "info": info,
                        "path": control_path,
                        "score": (right_distance, dy),
                        "actions": actions,
                    }
                )
            row_candidates.sort(key=lambda item: item["score"])
            for item in row_candidates:
                candidates.append(
                    {
                        "label": jab.info_to_dict(label_info),
                        "control": jab.info_to_dict(item["info"]),
                        "path": item["path"],
                        "actions": item["actions"],
                        "score": list(item["score"]),
                    }
                )
            if row_candidates:
                target = row_candidates[0]
                selected_contexts = {target["context"], label_context}
                release = [
                    context for context in owned if context not in selected_contexts
                ]
                jab.release_contexts(vm_id.value, release)
                return {
                    "ok": True,
                    "label": COUNTERPARTY_LABEL,
                    "source": "nearby",
                    "context": target["context"],
                    "vm_id": vm_id.value,
                    "owned_contexts": list(selected_contexts),
                    "window": {
                        "hwnd": int(scope_hwnd),
                        "class_name": "SunAwtCanvas",
                    },
                    "path": target["path"],
                    "dynamic_prefix": prefix,
                    "target": {
                        "label": jab.info_to_dict(label_info),
                        "control": jab.info_to_dict(target["info"]),
                        "actions": target["actions"],
                        "score": list(target["score"]),
                    },
                    "candidate_count": len(candidates),
                    "candidates": candidates[:8],
                }
    finally:
        if not selected_contexts:
            jab.release_contexts(vm_id.value, owned)

    return {
        "ok": False,
        "label": COUNTERPARTY_LABEL,
        "source": "nearby",
        "scope_hwnd": scope_hwnd,
        "candidate_count": len(candidates) if "candidates" in locals() else 0,
        "reason": "未在往来对象标签右侧找到 combo box",
    }


def collect_counterparty_controls_for_bounds_scan(
    jab,
    vm_id,
    context,
    controls,
    owned,
    require_showing=True,
    depth=0,
    path="0",
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    role = control_role(info)
    if role == "table" or depth >= jab.max_depth:
        return

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_path = f"{path}.{index}"
        child_info = jab.get_context_info(vm_id, child)
        if not child_info:
            jab.release_contexts(vm_id, [child])
            continue

        owned.append(child)
        states = (child_info.states_en_US.strip() or child_info.states.strip()).lower()
        showing = "visible" in states and "showing" in states
        if not require_showing or showing:
            controls.append((child, child_info, child_path))

        collect_counterparty_controls_for_bounds_scan(
            jab,
            vm_id,
            child,
            controls,
            owned,
            require_showing=require_showing,
            depth=depth + 1,
            path=child_path,
        )


def control_role(info):
    return (info.role_en_US.strip() or info.role.strip()).lower()


def slim_found(found):
    return {
        key: value
        for key, value in (found or {}).items()
        if key not in {"context", "vm_id", "owned_contexts", "candidates"}
    }


def find_counterparty_combo_by_path(jab, path, scope_hwnd=None):
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="combo box",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "path": path,
            "reason": "往来对象下拉控件未找到",
        }
    return {
        "ok": True,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "window": window_info,
        "path": path,
    }


def read_counterparty_combo_state(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    text = jab.get_text_context_value(vm_id, context)
    return {
        "name": info.name.strip() if info else "",
        "description": info.description.strip() if info else "",
        "text": str(text or "").strip(),
        "role": (info.role_en_US.strip() or info.role.strip()) if info else "",
        "states": (info.states_en_US.strip() or info.states.strip()) if info else "",
    }
