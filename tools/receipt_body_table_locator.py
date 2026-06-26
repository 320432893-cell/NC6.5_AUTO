import argparse
import ctypes
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_probe import JOBJECT, enum_windows  # noqa: E402


KEY_COLUMNS = {
    0: "往来对象",
    1: "收款业务类型",
    2: "收款性质",
    3: "币种",
    4: "收款银行账户",
    5: "科目",
    6: "汇率",
    7: "贷方原币金额",
    11: "结算方式",
    13: "订单客户",
    19: "客户",
}


def main():
    parser = argparse.ArgumentParser(
        description="Locate NC receipt body table by JAB table signature."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--max-rows", type=int, default=5)
    parser.add_argument("--scope-hwnd", type=int, default=None)
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        report = locate_receipt_body_table(
            jab,
            max_rows=args.max_rows,
            scope_hwnd=args.scope_hwnd,
        )
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)
    return 0 if report["best"] else 1


def locate_receipt_body_table(jab, max_rows=5, scope_hwnd=None):
    jab.ensure_started()
    candidates = []
    tables = find_tables_with_index_paths(jab, scope_hwnd=scope_hwnd)
    for table_index, table in enumerate(tables):
        context = table["context"]
        vm_id = table["vm_id"]
        owned = table["owned_contexts"]
        info = table["table_info"]
        window = table["window"]
        try:
            rows = read_key_rows(jab, vm_id, context, info, max_rows=max_rows)
            score, reasons = score_receipt_body_table(info, rows, window)
            candidates.append(
                {
                    "table_index": table_index,
                    "path": table["path"],
                    "window": window,
                    "row_count": info.rowCount,
                    "col_count": info.columnCount,
                    "score": score,
                    "reasons": reasons,
                    "key_columns": KEY_COLUMNS,
                    "rows": rows,
                    "selected_indexes": jab.get_selected_child_indexes(
                        vm_id,
                        context,
                        info.rowCount * info.columnCount,
                    ),
                    "bounds": table_bounds(jab, vm_id, context),
                }
            )
        finally:
            jab.release_contexts(vm_id, owned)

    ranked = sorted(candidates, key=lambda item: item["score"], reverse=True)
    best = ranked[0] if ranked and ranked[0]["score"] >= 5 else None
    return {"best": best, "candidates": ranked, "scope_hwnd": scope_hwnd}


def locate_receipt_body_table_cached(
    jab,
    cached=None,
    max_rows=5,
    scope_hwnd=None,
    required_cols=25,
):
    """Use a previously inferred table path first, then fall back to semantic scan."""
    fast = read_receipt_body_table_by_cached_path(
        jab,
        cached,
        max_rows=max_rows,
        scope_hwnd=scope_hwnd,
        required_cols=required_cols,
    )
    if fast.get("ok"):
        previous_best = dict(((cached or {}).get("best") or {}))
        best = {
            **previous_best,
            "path": fast.get("path"),
            "window": fast.get("window") or previous_best.get("window"),
            "row_count": fast.get("row_count"),
            "col_count": fast.get("col_count"),
            "rows": fast.get("rows"),
            "cache_hit": True,
            "validated_by_path": True,
        }
        return {
            "best": best,
            "candidates": [best],
            "scope_hwnd": scope_hwnd or fast.get("scope_hwnd"),
            "cache_hit": True,
            "fallback_used": False,
            "path_validation": fast,
        }

    fallback = locate_receipt_body_table(jab, max_rows=max_rows, scope_hwnd=scope_hwnd)
    fallback["cache_hit"] = False
    fallback["fallback_used"] = True
    fallback["path_validation"] = fast
    if fallback.get("best"):
        fallback["best"]["cache_hit"] = False
        fallback["best"]["validated_by_path"] = False
    return fallback


def read_receipt_body_table_by_cached_path(
    jab,
    located,
    max_rows=5,
    scope_hwnd=None,
    required_cols=25,
):
    best = (located or {}).get("best") or {}
    path = best.get("path")
    if not path:
        return {"ok": False, "reason": "未提供收款单明细表 path"}

    cached_window = best.get("window") or {}
    path_hwnd = cached_window.get("hwnd") or scope_hwnd
    context, vm_id, owned, window_info = jab.find_context_by_path_once(
        path,
        class_name=cached_window.get("class_name"),
        scope_hwnd=path_hwnd,
        role="table",
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "path": path,
            "scope_hwnd": path_hwnd,
            "reason": "按 cached path 读取收款单明细表失败",
        }

    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {
                "ok": False,
                "path": path,
                "scope_hwnd": path_hwnd,
                "reason": "cached path 命中控件但 table_info 不可读",
            }
        col_count = int(table_info.columnCount)
        if required_cols is not None and col_count != int(required_cols):
            return {
                "ok": False,
                "path": path,
                "scope_hwnd": path_hwnd,
                "row_count": int(table_info.rowCount),
                "col_count": col_count,
                "reason": f"cached path 命中表格列数不符：期望 {required_cols}，实际 {col_count}",
            }
        return {
            "ok": True,
            "path": path,
            "scope_hwnd": path_hwnd,
            "window": window_info or cached_window,
            "row_count": int(table_info.rowCount),
            "col_count": col_count,
            "rows": read_key_rows(jab, vm_id, context, table_info, max_rows=max_rows),
        }
    finally:
        jab.release_contexts(vm_id, owned)


def find_tables_with_index_paths(jab, scope_hwnd=None):
    result = []
    windows = (
        jab.get_scoped_windows(scope_hwnd, include_children=True)
        if scope_hwnd is not None
        else enum_windows(include_children=True)
    )
    for hwnd, title, class_name, pid, visible in windows:
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

        window = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
        }
        collect_tables_in_tree(
            jab,
            vm_id.value,
            root_context.value,
            depth=0,
            path_indexes=[],
            owned_contexts=[],
            window=window,
            result=result,
        )
    return result


def collect_tables_in_tree(
    jab,
    vm_id,
    context,
    depth,
    path_indexes,
    owned_contexts,
    window,
    result,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if role == "table":
        table_info = jab.get_table_info(vm_id, context)
        if table_info and table_info.rowCount > 0 and table_info.columnCount > 0:
            result.append(
                {
                    "context": context,
                    "vm_id": vm_id,
                    "owned_contexts": list(owned_contexts),
                    "table_info": table_info,
                    "window": window,
                    "path": "0" + "".join(f".{index}" for index in path_indexes),
                }
            )
        return

    if depth >= jab.max_depth:
        return

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        collect_tables_in_tree(
            jab,
            vm_id,
            child,
            depth + 1,
            path_indexes + [index],
            owned_contexts + [child],
            window,
            result,
        )


def read_key_rows(jab, vm_id, table_context, table_info, max_rows=5):
    row_limit = min(table_info.rowCount, max_rows)
    rows = []
    for row in range(row_limit):
        cells = {}
        selected = False
        for col in KEY_COLUMNS:
            if col >= table_info.columnCount:
                continue
            text, is_selected = jab.get_table_cell_text_and_selection(
                vm_id, table_context, row, col
            )
            cells[str(col)] = text
            selected = selected or is_selected
        rows.append({"row_index": row, "selected": selected, "cells": cells})
    return rows


def score_receipt_body_table(table_info, rows, window):
    score = 0
    reasons = []
    if table_info.columnCount == 25:
        score += 3
        reasons.append("cols=25")
    if table_info.rowCount >= 2:
        score += 2
        reasons.append("rows>=2")
    if window.get("class_name") == "SunAwtCanvas":
        score += 1
        reasons.append("window=SunAwtCanvas")

    first = rows[0]["cells"] if rows else {}
    if first.get("0") == "客户":
        score += 2
        reasons.append("col0=客户")
    if first.get("1") in ("货款", "手续费"):
        score += 2
        reasons.append("col1=收款业务类型值")
    if first.get("2") in ("应收款", "手续费"):
        score += 1
        reasons.append("col2=收款性质值")
    if first.get("3") in ("人民币", "美元"):
        score += 1
        reasons.append("col3=币种值")
    if any((row["cells"].get("13") or row["cells"].get("19")) for row in rows):
        score += 1
        reasons.append("客户/订单客户列有值")
    if table_info.rowCount == 1 and first.get("0") != "客户":
        if table_info.columnCount == 25 and window.get("class_name") == "SunAwtCanvas":
            score += 1
            reasons.append("single-row-blank-receipt-body")
        else:
            score -= 3
            reasons.append("single-row-non-body")
    return score, reasons


def table_bounds(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    return [info.x, info.y, info.width, info.height]


def print_text(report):
    best = report["best"]
    if best:
        print(
            "best: "
            f"table={best['table_index']} path={best['path']} "
            f"rows={best['row_count']} cols={best['col_count']} "
            f"score={best['score']} reasons={best['reasons']}"
        )
    else:
        print("best: <none>")

    print("candidates:", len(report["candidates"]))
    for item in report["candidates"]:
        print(
            f"  table={item['table_index']} path={item['path']} "
            f"rows={item['row_count']} cols={item['col_count']} "
            f"score={item['score']} reasons={item['reasons']} "
            f"selected={item['selected_indexes']} bounds={item['bounds']}"
        )
        for row in item["rows"][:3]:
            print(
                f"    row {row['row_index']} selected={row['selected']}: {row['cells']}"
            )


if __name__ == "__main__":
    raise SystemExit(main())
