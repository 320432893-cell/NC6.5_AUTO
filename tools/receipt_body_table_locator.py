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
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        report = locate_receipt_body_table(jab, max_rows=args.max_rows)
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)
    return 0 if report["best"] else 1


def locate_receipt_body_table(jab, max_rows=5):
    jab.ensure_started()
    candidates = []
    tables = find_tables_with_index_paths(jab)
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
    return {"best": best, "candidates": ranked}


def find_tables_with_index_paths(jab):
    result = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
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
