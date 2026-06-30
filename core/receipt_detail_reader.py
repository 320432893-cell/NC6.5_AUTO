# 职责：读取收款单明细表缓存 path、指定行单元格和行数
# 不做什么：不写入 NC，不发送键盘输入，不决定业务字段顺序
# 允许依赖层：core.receipt_body_table_locator、core.receipt_self_made_fill_trial
# 谁不应该 import：配置校验、Excel/Sheet 写入模块不应 import

import time

from core.receipt_body_table_locator import (
    locate_receipt_body_table_cached,
    read_receipt_body_table_by_cached_path,
)
from core.receipt_self_made_fill_trial import read_body_table


def read_body_table_by_path(jab, located, step, max_rows=5, semantic_fallback=False):
    result = read_receipt_body_table_by_cached_path(
        jab,
        located,
        max_rows=max(
            max_rows,
            int(((located or {}).get("best") or {}).get("row_count") or 0),
        ),
    )
    if result.get("ok"):
        return {
            "step": step,
            "ok": True,
            "fast_path": True,
            "semantic_fallback_used": False,
            "path": result.get("path"),
            "row_count": result.get("row_count"),
            "col_count": result.get("col_count"),
            "rows": result.get("rows"),
        }
    if semantic_fallback:
        refreshed = locate_receipt_body_table_cached(
            jab,
            cached=located,
            max_rows=max_rows,
            scope_hwnd=(((located or {}).get("best") or {}).get("window") or {}).get(
                "hwnd"
            ),
        )
        best = refreshed.get("best")
        if best:
            if isinstance(located, dict):
                located["best"] = best
            return {
                "step": step,
                "ok": True,
                "fast_path": bool(refreshed.get("cache_hit")),
                "semantic_fallback_used": bool(refreshed.get("fallback_used")),
                "path": best.get("path"),
                "row_count": best.get("row_count"),
                "col_count": best.get("col_count"),
                "rows": best.get("rows"),
                "path_validation": refreshed.get("path_validation"),
            }
    return {"step": step, "ok": False, "reason": result.get("reason")}


def read_first_row_cells(jab, located=None):
    snapshot = read_body_table_by_path(jab, located, "field_readback")
    if not snapshot.get("ok"):
        snapshot = read_body_table(
            jab,
            "field_readback",
            scope_hwnd=(((located or {}).get("best") or {}).get("window") or {}).get(
                "hwnd"
            ),
        )
    if not snapshot.get("ok"):
        return snapshot, {}
    rows = snapshot.get("rows") or []
    cells = (rows[0].get("cells") if rows else {}) or {}
    return snapshot, cells


def read_row_cells(jab, row_index, located=None, semantic_fallback=False):
    snapshot = read_body_table_by_path(
        jab,
        located,
        f"row_{row_index}_readback",
        semantic_fallback=semantic_fallback,
    )
    if not snapshot.get("ok"):
        snapshot = read_body_table(
            jab,
            f"row_{row_index}_readback",
            scope_hwnd=(((located or {}).get("best") or {}).get("window") or {}).get(
                "hwnd"
            ),
        )
    if not snapshot.get("ok"):
        return snapshot, {}
    rows = snapshot.get("rows") or []
    for row in rows:
        if int(row.get("row_index", -1)) == int(row_index):
            return snapshot, (row.get("cells") or {})
    return snapshot, {}


def read_table_row_count_by_path(jab, located, semantic_fallback=False):
    result = read_receipt_body_table_by_cached_path(jab, located, max_rows=0)
    if not result.get("ok"):
        if semantic_fallback:
            refreshed = locate_receipt_body_table_cached(
                jab,
                cached=located,
                max_rows=0,
                scope_hwnd=(
                    ((located or {}).get("best") or {}).get("window") or {}
                ).get("hwnd"),
            )
            best = refreshed.get("best")
            if best:
                if isinstance(located, dict):
                    located["best"] = best
                return {
                    "ok": True,
                    "row_count": best.get("row_count"),
                    "col_count": best.get("col_count"),
                    "semantic_fallback_used": bool(refreshed.get("fallback_used")),
                    "path": best.get("path"),
                    "path_validation": refreshed.get("path_validation"),
                }
        return {"ok": False, "reason": result.get("reason")}
    return {
        "ok": True,
        "row_count": result.get("row_count"),
        "col_count": result.get("col_count"),
        "semantic_fallback_used": False,
    }


def wait_table_row_count_by_path(
    jab, located, expected_rows, label, timeout=0.45, interval=0.035
):
    started_at = time.perf_counter()
    deadline = time.perf_counter() + timeout
    last = {}
    while True:
        last = read_table_row_count_by_path(jab, located)
        rows = int(last.get("row_count") or 0) if last.get("ok") else 0
        if last.get("ok") and rows == expected_rows:
            return {
                "ok": True,
                "label": label,
                "seconds": round(time.perf_counter() - started_at, 3),
                "expected_rows": expected_rows,
                "actual_rows": rows,
                "snapshot": last,
            }
        if time.perf_counter() >= deadline:
            return {
                "ok": False,
                "label": label,
                "seconds": round(time.perf_counter() - started_at, 3),
                "expected_rows": expected_rows,
                "actual_rows": rows,
                "snapshot": last,
                "reason": f"等待行数变为 {expected_rows} 超时，实际 {rows}",
            }
        time.sleep(interval)


def read_located_body_table(jab, located, step, scope_hwnd=None, max_rows=3):
    snapshot = read_body_table_by_path(jab, located, step, max_rows=max_rows)
    if snapshot.get("ok"):
        return snapshot
    refreshed = locate_receipt_body_table_cached(
        jab,
        cached=located,
        max_rows=max_rows,
        scope_hwnd=scope_hwnd
        or (((located or {}).get("best") or {}).get("window") or {}).get("hwnd"),
    )
    best = refreshed.get("best")
    if not best:
        return {
            "step": step,
            "ok": False,
            "reason": snapshot.get("reason") or "body table not found",
            "path_validation": refreshed.get("path_validation"),
            "candidates": refreshed.get("candidates", [])[:3],
        }
    return {
        "step": step,
        "ok": True,
        "path": best.get("path"),
        "row_count": best.get("row_count"),
        "col_count": best.get("col_count"),
        "rows": best.get("rows"),
        "cache_hit": refreshed.get("cache_hit"),
        "fallback_used": refreshed.get("fallback_used"),
    }


def wait_body_row_count(
    jab,
    located,
    expected_rows,
    label,
    scope_hwnd=None,
    timeout=0.75,
    interval=0.06,
):
    started_at = time.perf_counter()
    deadline = time.perf_counter() + timeout
    last = {}
    while True:
        last = read_located_body_table(jab, located, label, scope_hwnd=scope_hwnd)
        rows = int(last.get("row_count") or 0) if last.get("ok") else 0
        if last.get("ok") and rows == expected_rows:
            return {
                "ok": True,
                "seconds": round(time.perf_counter() - started_at, 3),
                "expected_rows": expected_rows,
                "actual_rows": rows,
                "snapshot": last,
            }
        if time.perf_counter() >= deadline:
            return {
                "ok": False,
                "seconds": round(time.perf_counter() - started_at, 3),
                "expected_rows": expected_rows,
                "actual_rows": rows,
                "snapshot": last,
                "reason": f"等待行数变为 {expected_rows} 超时，实际 {rows}",
            }
        time.sleep(interval)
