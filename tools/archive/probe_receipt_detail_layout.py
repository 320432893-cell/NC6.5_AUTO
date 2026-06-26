import argparse
import ctypes
import json
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_probe import JOBJECT, enum_windows  # noqa: E402
from tools.receipt_body_table_locator import (  # noqa: E402
    KEY_COLUMNS,
    find_tables_with_index_paths,
    locate_receipt_body_table,
    read_key_rows,
    table_bounds,
)


TARGET_TEXTS = {
    "客户",
    "收款业务类型",
    "收款性质",
    "币种",
    "收款银行账户",
    "科目",
    "贷方原币金额",
    "结算方式",
    "合计",
    "1",
    "2",
}


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Read-only probe for NC receipt detail visual layout. It compares "
            "body table, row-number tables, total-row tables, and nearby text/header bounds."
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--max-rows", type=int, default=3)
    parser.add_argument("--max-controls", type=int, default=180)
    parser.add_argument("--max-depth", type=int, default=10)
    parser.add_argument("--max-visited", type=int, default=1800)
    parser.add_argument("--activate-window-before-read", action="store_true")
    parser.add_argument("--activate-wait", type=float, default=0.2)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        report = probe_layout(
            jab,
            max_rows=args.max_rows,
            max_controls=args.max_controls,
            max_depth=args.max_depth,
            max_visited=args.max_visited,
            activate_window_before_read=args.activate_window_before_read,
            activate_wait=max(0.0, args.activate_wait),
        )
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    else:
        print_text(report)
    return 0 if report.get("best_table") else 1


def probe_layout(
    jab,
    max_rows,
    max_controls,
    max_depth,
    max_visited,
    activate_window_before_read,
    activate_wait,
):
    located = locate_receipt_body_table(jab, max_rows=max_rows)
    best = located.get("best")
    report = {
        "read_only": True,
        "requested": {
            "max_rows": max_rows,
            "max_controls": max_controls,
            "max_depth": max_depth,
            "max_visited": max_visited,
            "activate_window_before_read": activate_window_before_read,
            "activate_wait": activate_wait,
        },
        "foreground_before": safe_foreground(jab),
        "activation": None,
        "foreground_after_activation": None,
        "best_table": summarize_best(best),
        "tables_same_window": [],
        "text_controls_same_window": [],
        "layout_hints": {},
    }
    if not best:
        report["layout_hints"] = {"reason": "未定位到收款单明细表"}
        return report

    if activate_window_before_read:
        report["activation"] = activate_nc_window(best.get("window") or {})
        if activate_wait:
            time.sleep(activate_wait)
        report["foreground_after_activation"] = safe_foreground(jab)
        located = locate_receipt_body_table(jab, max_rows=max_rows)
        best = located.get("best") or best
        report["best_table_after_activation"] = summarize_best(best)

    window = best.get("window") or {}
    report["tables_same_window"] = collect_tables_for_window(
        jab,
        window,
        max_rows=max_rows,
    )
    report["text_controls_same_window"] = collect_text_controls_for_window(
        jab,
        window,
        max_controls=max_controls,
        max_depth=max_depth,
        max_visited=max_visited,
    )
    report["layout_hints"] = infer_layout(report)
    return report


def summarize_best(best):
    if not best:
        return None
    return {
        "table_index": best.get("table_index"),
        "path": best.get("path"),
        "window": best.get("window"),
        "row_count": best.get("row_count"),
        "col_count": best.get("col_count"),
        "score": best.get("score"),
        "reasons": best.get("reasons"),
        "bounds": best.get("bounds"),
        "rows": best.get("rows"),
        "selected_indexes": best.get("selected_indexes"),
    }


def collect_tables_for_window(jab, target_window, max_rows):
    target_hwnd = int((target_window or {}).get("hwnd") or 0)
    tables = []
    for index, table in enumerate(find_tables_with_index_paths(jab, scope_hwnd=target_hwnd)):
        context = table["context"]
        vm_id = table["vm_id"]
        owned = table["owned_contexts"]
        table_info = table["table_info"]
        try:
            bounds = table_bounds(jab, vm_id, context)
            child_count = int(table_info.rowCount) * int(table_info.columnCount)
            rows = read_rows_generic(
                jab,
                vm_id,
                context,
                table_info,
                max_rows=max_rows,
            )
            key_rows = read_key_rows(jab, vm_id, context, table_info, max_rows=max_rows)
            tables.append(
                {
                    "table_index": index,
                    "path": table.get("path"),
                    "window": table.get("window"),
                    "row_count": int(table_info.rowCount),
                    "col_count": int(table_info.columnCount),
                    "bounds": bounds,
                    "valid_bounds": bounds_valid(bounds),
                    "selected_indexes": jab.get_selected_child_indexes(
                        vm_id,
                        context,
                        child_count,
                    ),
                    "rows": rows,
                    "key_rows": key_rows,
                    "classification": classify_table(table_info, rows, key_rows, bounds),
                }
            )
        finally:
            jab.release_contexts(vm_id, owned)
    return sorted(
        tables,
        key=lambda item: (
            item["bounds"][1] if item.get("bounds") else 999999,
            item["bounds"][0] if item.get("bounds") else 999999,
            item.get("path") or "",
        ),
    )


def read_rows_generic(jab, vm_id, table_context, table_info, max_rows):
    rows = []
    row_limit = min(int(table_info.rowCount), int(max_rows))
    col_limit = min(int(table_info.columnCount), 12)
    for row in range(row_limit):
        cells = {}
        selected = False
        for col in range(col_limit):
            text, is_selected = jab.get_table_cell_text_and_selection(
                vm_id,
                table_context,
                row,
                col,
            )
            cells[str(col)] = text
            selected = selected or bool(is_selected)
        rows.append({"row_index": row, "selected": selected, "cells": cells})
    return rows


def classify_table(table_info, rows, key_rows, bounds):
    row_count = int(table_info.rowCount)
    col_count = int(table_info.columnCount)
    first_cells = (rows[0].get("cells") if rows else {}) or {}
    key_first = (key_rows[0].get("cells") if key_rows else {}) or {}
    values = list(first_cells.values()) + list(key_first.values())
    non_empty = [str(value).strip() for value in values if str(value or "").strip()]
    if col_count == 25 and key_first.get("0") == "客户":
        label = "body-data-25col"
    elif col_count == 24 and key_first.get("0") == "客户":
        label = "body-data-24col-mirror"
    elif col_count == 1 and any(value in ("1", "2", "合计") for value in non_empty):
        label = "row-number-or-total"
    elif row_count == 1 and any(value == "合计" for value in non_empty):
        label = "total-row"
    elif row_count == 1 and col_count in (24, 25):
        label = "single-row-detail-or-total"
    else:
        label = "other-table"
    return {
        "label": label,
        "valid_bounds": bounds_valid(bounds),
        "first_non_empty": non_empty[:8],
    }


def collect_text_controls_for_window(
    jab,
    target_window,
    max_controls,
    max_depth,
    max_visited,
):
    target_hwnd = int((target_window or {}).get("hwnd") or 0)
    if not target_hwnd:
        return []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if int(hwnd) != target_hwnd:
            continue
        if not visible or not jab.dll.isJavaWindow(hwnd):
            return []
        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            return []
        window = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
        }
        owned = [root_context.value]
        result = []
        stats = {"visited": 0, "stop_reason": None}
        try:
            collect_text_controls(
                jab,
                vm_id.value,
                root_context.value,
                path="0",
                depth=0,
                owned=owned,
                result=result,
                window=window,
                max_controls=max_controls,
                max_depth=max_depth,
                max_visited=max_visited,
                stats=stats,
            )
        finally:
            jab.release_contexts(vm_id.value, owned)
        result.sort(
            key=lambda item: (
                item["bounds"][1] if item.get("bounds") else 999999,
                item["bounds"][0] if item.get("bounds") else 999999,
                item.get("path") or "",
            )
        )
        return result
    return []


def collect_text_controls(
    jab,
    vm_id,
    context,
    path,
    depth,
    owned,
    result,
    window,
    max_controls,
    max_depth,
    max_visited,
    stats,
):
    if len(result) >= max_controls:
        stats["stop_reason"] = "max_controls reached"
        return
    if stats["visited"] >= max_visited:
        stats["stop_reason"] = "max_visited reached"
        return
    stats["visited"] += 1
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    role = info.role_en_US.strip() or info.role.strip()
    states = info.states_en_US.strip() or info.states.strip()
    text = jab.get_text_context_value(vm_id, context) if info.accessibleText else ""
    bounds = [info.x, info.y, info.width, info.height]
    name = info.name.strip()
    description = info.description.strip()
    combined = " ".join([name, description, text, role, states])
    keep = any(target in combined for target in TARGET_TEXTS)
    if keep:
        result.append(
            {
                "path": path,
                "window": window,
                "role": role,
                "name": name,
                "description": description,
                "states": states,
                "text": text,
                "bounds": bounds,
                "valid_bounds": bounds_valid(bounds),
                "children": info.childrenCount,
            }
        )
    if depth >= max_depth or role.lower() == "table":
        return
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        owned.append(child)
        collect_text_controls(
            jab,
            vm_id,
            child,
            f"{path}.{index}",
            depth + 1,
            owned,
            result,
            window,
            max_controls,
            max_depth,
            max_visited,
            stats,
        )
        if len(result) >= max_controls or stats["visited"] >= max_visited:
            return


def infer_layout(report):
    tables = report.get("tables_same_window") or []
    body_tables = [
        item
        for item in tables
        if (item.get("classification") or {}).get("label") == "body-data-25col"
    ]
    row_number_tables = [
        item
        for item in tables
        if (item.get("classification") or {}).get("label") == "row-number-or-total"
    ]
    total_tables = [
        item
        for item in tables
        if "total" in ((item.get("classification") or {}).get("label") or "")
    ]
    valid_table_bounds = [item for item in tables if item.get("valid_bounds")]
    return {
        "body_table_count": len(body_tables),
        "row_number_or_total_table_count": len(row_number_tables),
        "total_table_count": len(total_tables),
        "valid_table_bounds_count": len(valid_table_bounds),
        "body_tables": [
            summarize_layout_table(item)
            for item in body_tables[:4]
        ],
        "row_number_or_total_tables": [
            summarize_layout_table(item)
            for item in row_number_tables[:8]
        ],
        "total_tables": [
            summarize_layout_table(item)
            for item in total_tables[:8]
        ],
        "valid_bounds_tables": [
            summarize_layout_table(item)
            for item in valid_table_bounds[:8]
        ],
        "candidate_data_row_y": infer_data_row_y(body_tables, row_number_tables),
    }


def summarize_layout_table(item):
    return {
        "table_index": item.get("table_index"),
        "path": item.get("path"),
        "row_count": item.get("row_count"),
        "col_count": item.get("col_count"),
        "bounds": item.get("bounds"),
        "valid_bounds": item.get("valid_bounds"),
        "selected_indexes": item.get("selected_indexes"),
        "classification": item.get("classification"),
    }


def infer_data_row_y(body_tables, row_number_tables):
    row_tables = [
        item
        for item in row_number_tables
        if item.get("valid_bounds")
        and any(
            (row.get("cells") or {}).get("0") == "1"
            for row in item.get("rows") or []
        )
    ]
    body_valid = [item for item in body_tables if item.get("valid_bounds")]
    if row_tables:
        first = row_tables[0]
        x, y, width, height = first["bounds"]
        row_count = max(1, int(first.get("row_count") or 1))
        return {
            "ok": True,
            "source": "row-number-table",
            "table_index": first.get("table_index"),
            "bounds": first.get("bounds"),
            "row0_center_y": round(y + (height / row_count) / 2, 1),
        }
    if body_valid:
        first = body_valid[0]
        x, y, width, height = first["bounds"]
        row_count = max(1, int(first.get("row_count") or 1))
        return {
            "ok": True,
            "source": "body-table",
            "table_index": first.get("table_index"),
            "bounds": first.get("bounds"),
            "row0_center_y": round(y + (height / row_count) / 2, 1),
        }
    return {
        "ok": False,
        "reason": "没有找到有效 bounds 的行号表或 25 列数据表",
    }


def safe_foreground(jab):
    try:
        return jab.get_foreground_window_info()
    except Exception as exc:
        return {"ok": False, "reason": f"{type(exc).__name__}: {exc}"}


def activate_nc_window(window):
    hwnd = int((window or {}).get("hwnd") or 0)
    result = {
        "ok": False,
        "skipped": False,
        "hwnd": hwnd,
        "platform": os.name,
        "reason": None,
    }
    if os.name != "nt":
        result["skipped"] = True
        result["reason"] = "非 Windows Python，跳过窗口激活"
        return result
    if not hwnd:
        result["reason"] = "缺少 NC 表格窗口 hwnd"
        return result
    try:
        user32 = ctypes.windll.user32
        user32.ShowWindow(hwnd, 9)
        time.sleep(0.03)
        result["set_foreground_ok"] = bool(user32.SetForegroundWindow(hwnd))
        result["ok"] = bool(result["set_foreground_ok"])
        if not result["ok"]:
            result["reason"] = "SetForegroundWindow 返回失败"
        return result
    except Exception as exc:
        result["reason"] = f"{type(exc).__name__}: {exc}"
        return result


def bounds_valid(bounds):
    return (
        isinstance(bounds, (list, tuple))
        and len(bounds) == 4
        and int(bounds[0]) >= 0
        and int(bounds[1]) >= 0
        and int(bounds[2]) > 0
        and int(bounds[3]) > 0
    )


def print_text(report):
    print("NC 收款单明细布局只读探测")
    print(f"  foreground_before={report.get('foreground_before')}")
    print(f"  activation={report.get('activation')}")
    print(f"  foreground_after_activation={report.get('foreground_after_activation')}")
    print(f"  best={report.get('best_table')}")
    hints = report.get("layout_hints") or {}
    print(f"  layout_hints={json.dumps(hints, ensure_ascii=False)}")
    print("  tables:")
    for item in report.get("tables_same_window") or []:
        print(
            "    "
            f"table={item.get('table_index')} cls={(item.get('classification') or {}).get('label')} "
            f"rows={item.get('row_count')} cols={item.get('col_count')} "
            f"bounds={item.get('bounds')} selected={item.get('selected_indexes')} "
            f"first={(item.get('classification') or {}).get('first_non_empty')}"
        )
    print("  text controls:")
    for item in (report.get("text_controls_same_window") or [])[:40]:
        print(
            "    "
            f"path={item.get('path')} role={item.get('role')!r} "
            f"name={item.get('name')!r} desc={item.get('description')!r} "
            f"text={item.get('text')!r} bounds={item.get('bounds')}"
        )


if __name__ == "__main__":
    raise SystemExit(main())
