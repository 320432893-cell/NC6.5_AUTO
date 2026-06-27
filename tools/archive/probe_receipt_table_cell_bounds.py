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
from core.jab_probe import AccessibleActions, AccessibleTableCellInfo, JOBJECT  # noqa: E402
from tools.receipt_body_table_locator import (  # noqa: E402
    KEY_COLUMNS,
    locate_receipt_body_table,
)


DEFAULT_COLS = sorted(KEY_COLUMNS)
HEADER_LABELS = {
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
        description=(
            "Probe NC receipt body table real cell bounds through JAB "
            "AccessibleTable APIs. Read-only: no mouse/keyboard input."
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--rows", default="0,1", help="Comma-separated 0-based rows.")
    parser.add_argument(
        "--cols",
        default=",".join(str(col) for col in DEFAULT_COLS),
        help="Comma-separated 0-based columns.",
    )
    parser.add_argument("--max-rows", type=int, default=5)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    rows = parse_indexes(args.rows)
    cols = parse_indexes(args.cols)

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        report = probe(jab, rows=rows, cols=cols, max_rows=args.max_rows)
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)

    return 0 if report.get("scheme_a", {}).get("usable") else 1


def parse_indexes(text):
    indexes = []
    for part in str(text).split(","):
        part = part.strip()
        if not part:
            continue
        indexes.append(int(part))
    return indexes


def probe(jab, rows, cols, max_rows):
    located = locate_receipt_body_table(jab, max_rows=max_rows)
    best = located.get("best")
    report = {
        "read_only": True,
        "requested": {"rows": rows, "cols": cols},
        "located": summarize_located(located),
        "table": None,
        "scheme_a": {"usable": False, "reason": None, "cells": []},
        "scheme_b": {"usable": False, "reason": None, "headers": []},
    }
    if not best:
        report["scheme_a"]["reason"] = "receipt body table not located"
        report["scheme_b"]["reason"] = "receipt body table not located"
        return report

    context, vm_id, owned, window = jab.find_context_by_path_once(
        best["path"],
        class_name=best["window"].get("class_name"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        report["scheme_a"]["reason"] = f"table context not found by path {best['path']}"
        report["scheme_b"]["reason"] = f"table context not found by path {best['path']}"
        return report

    try:
        table_info = jab.get_table_info(vm_id, context)
        table_context = describe_context(jab, vm_id, context)
        if not table_info:
            report["table"] = {
                "path": best["path"],
                "window": window,
                "context": table_context,
            }
            report["scheme_a"]["reason"] = "getAccessibleTableInfo failed"
            report["scheme_b"]["reason"] = "getAccessibleTableInfo failed"
            return report

        report["table"] = {
            "path": best["path"],
            "window": window,
            "row_count": table_info.rowCount,
            "col_count": table_info.columnCount,
            "context": table_context,
        }
        report["scheme_a"] = probe_table_cells(
            jab, vm_id, context, best["path"], table_info, rows, cols
        )
        report["scheme_b"] = probe_column_headers(
            jab, vm_id, context, best["path"], table_context, window, cols
        )
        return report
    finally:
        jab.release_contexts(vm_id, owned)


def summarize_located(located):
    best = located.get("best")
    candidates = []
    for item in located.get("candidates", [])[:5]:
        candidates.append(
            {
                "table_index": item.get("table_index"),
                "path": item.get("path"),
                "window": item.get("window"),
                "row_count": item.get("row_count"),
                "col_count": item.get("col_count"),
                "score": item.get("score"),
                "reasons": item.get("reasons"),
                "bounds": item.get("bounds"),
            }
        )
    return {
        "best": {
            "table_index": best.get("table_index"),
            "path": best.get("path"),
            "window": best.get("window"),
            "row_count": best.get("row_count"),
            "col_count": best.get("col_count"),
            "score": best.get("score"),
            "reasons": best.get("reasons"),
            "bounds": best.get("bounds"),
        }
        if best
        else None,
        "candidate_count": len(located.get("candidates", [])),
        "candidates": candidates,
    }


def probe_table_cells(jab, vm_id, table_context, table_path, table_info, rows, cols):
    result = {
        "usable": False,
        "reason": None,
        "api": "getAccessibleTableCellInfo(row,col)",
        "cells": [],
    }
    if not hasattr(jab.dll, "getAccessibleTableCellInfo"):
        result["reason"] = "JAB DLL does not expose getAccessibleTableCellInfo"
        return result

    success_count = 0
    for row in rows:
        for col in cols:
            cell = probe_one_cell(
                jab, vm_id, table_context, table_path, table_info, row, col
            )
            if cell.get("ok") and cell.get("context", {}).get("bounds"):
                success_count += 1
            result["cells"].append(cell)

    if success_count:
        result["usable"] = True
        result["reason"] = f"{success_count} cells returned accessibleContext+bounds"
    else:
        result["reason"] = "no requested cell returned accessibleContext+bounds"
    return result


def probe_one_cell(jab, vm_id, table_context, table_path, table_info, row, col):
    item = {
        "row": row,
        "col": col,
        "label": HEADER_LABELS.get(col),
        "ok": False,
        "reason": None,
        "cell_path": f"{table_path}::cell[{row},{col}]",
        "cell_info": None,
        "context": None,
    }
    if row < 0 or row >= table_info.rowCount:
        item["reason"] = f"row out of range: {row} >= {table_info.rowCount}"
        return item
    if col < 0 or col >= table_info.columnCount:
        item["reason"] = f"col out of range: {col} >= {table_info.columnCount}"
        return item

    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    item["cell_info"] = {
        "ok": bool(ok),
        "accessible_context": int(cell_info.accessibleContext or 0),
        "index": cell_info.index,
        "row": cell_info.row,
        "column": cell_info.column,
        "row_extent": cell_info.rowExtent,
        "column_extent": cell_info.columnExtent,
        "is_selected": bool(cell_info.isSelected),
    }
    if not ok:
        item["reason"] = "getAccessibleTableCellInfo returned false"
        return item
    if not cell_info.accessibleContext:
        item["reason"] = "cell has no accessibleContext"
        return item

    context = int(cell_info.accessibleContext)
    item["context"] = describe_context(jab, vm_id, context)
    if not item["context"]:
        item["reason"] = "getAccessibleContextInfo failed for cell context"
        return item

    item["ok"] = True
    return item


def describe_context(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    role = info.role_en_US.strip() or info.role.strip()
    states = info.states_en_US.strip() or info.states.strip()
    return {
        "context": int(context),
        "role": role,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "states": states,
        "bounds": [info.x, info.y, info.width, info.height],
        "children": info.childrenCount,
        "accessible_text": bool(info.accessibleText),
        "text": jab.get_text_context_value(vm_id, context),
        "accessible_action": bool(info.accessibleAction),
        "actions": get_action_names(jab, vm_id, context)
        if info.accessibleAction
        else [],
    }


def get_action_names(jab, vm_id, context):
    if not hasattr(jab.dll, "getAccessibleActions"):
        return []
    actions = AccessibleActions()
    if not jab.dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
        return []
    return [
        actions.actionInfo[index].name.strip() for index in range(actions.actionsCount)
    ]


def probe_column_headers(
    jab, vm_id, table_context, table_path, table_context_info, window, cols
):
    labels = {HEADER_LABELS[col] for col in cols if col in HEADER_LABELS}
    result = {
        "usable": False,
        "reason": None,
        "strategy": "scan same Java window for controls whose name/description matches target column labels",
        "headers": [],
    }
    if not labels:
        result["reason"] = "no known labels requested"
        return result

    root_context = get_window_root_context(jab, window)
    if not root_context:
        result["reason"] = "window root context not found"
        return result

    table_bounds = (table_context_info or {}).get("bounds")
    try:
        collect_header_candidates(
            jab,
            root_context["vm_id"],
            root_context["context"],
            path="0",
            depth=0,
            labels=labels,
            table_bounds=table_bounds,
            table_path=table_path,
            result=result["headers"],
        )
    finally:
        jab.release_contexts(root_context["vm_id"], root_context["owned"])

    result["headers"].sort(key=header_sort_key)
    if result["headers"]:
        result["usable"] = True
        result["reason"] = (
            f"{len(result['headers'])} matching header/control candidates"
        )
    else:
        result["reason"] = "no matching header/control candidates found"
    return result


def get_window_root_context(jab, window):
    hwnd = window.get("hwnd")
    if not hwnd:
        return None
    vm_id = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        hwnd,
        ctypes.byref(vm_id),
        ctypes.byref(root_context),
    ):
        return None
    return {
        "vm_id": vm_id.value,
        "context": root_context.value,
        "owned": [root_context.value],
    }


def collect_header_candidates(
    jab,
    vm_id,
    context,
    path,
    depth,
    labels,
    table_bounds,
    table_path,
    result,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    context_texts = {info.name.strip(), info.description.strip()}
    matched = sorted(text for text in context_texts if text in labels)
    if matched:
        bounds = [info.x, info.y, info.width, info.height]
        result.append(
            {
                "matched": matched,
                "path": path,
                "relative_to_table": path_relation(path, table_path),
                "role": info.role_en_US.strip() or info.role.strip(),
                "name": info.name.strip(),
                "description": info.description.strip(),
                "states": info.states_en_US.strip() or info.states.strip(),
                "bounds": bounds,
                "near_table": bounds_near_or_above_table(bounds, table_bounds),
                "children": info.childrenCount,
            }
        )

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if depth >= jab.max_depth or role == "table":
        return

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_path = f"{path}.{index}"
        try:
            collect_header_candidates(
                jab,
                vm_id,
                child,
                child_path,
                depth + 1,
                labels,
                table_bounds,
                table_path,
                result,
            )
        finally:
            jab.release_contexts(vm_id, [child])


def path_relation(path, table_path):
    if path == table_path:
        return "table"
    if path.startswith(f"{table_path}."):
        return "inside_table_subtree"
    parent = ".".join(table_path.split(".")[:-1])
    if parent and path.startswith(f"{parent}."):
        return "same_parent_subtree"
    return "same_window"


def bounds_near_or_above_table(bounds, table_bounds):
    if not bounds or not table_bounds:
        return None
    x, y, width, height = bounds
    tx, ty, tw, th = table_bounds
    if width <= 0 or height <= 0:
        return False
    horizontal_overlap = min(x + width, tx + tw) - max(x, tx)
    if horizontal_overlap <= 0:
        return False
    close_above = ty - 120 <= y + height <= ty + 10
    inside_table_y = ty <= y <= ty + th
    return close_above or inside_table_y


def header_sort_key(item):
    bounds = item.get("bounds") or [0, 0, 0, 0]
    near = item.get("near_table")
    return (0 if near else 1, bounds[1], bounds[0], item.get("path") or "")


def print_text(report):
    print("read_only:", report.get("read_only"))
    table = report.get("table")
    if table:
        print(
            "table: "
            f"path={table.get('path')} rows={table.get('row_count')} "
            f"cols={table.get('col_count')} window={table.get('window')}"
        )
        print(f"table context: {table.get('context')}")
    else:
        print("table: <none>")
        print(f"located: {json.dumps(report.get('located'), ensure_ascii=False)}")

    scheme_a = report.get("scheme_a", {})
    print(f"\n方案 A usable={scheme_a.get('usable')} reason={scheme_a.get('reason')}")
    for cell in scheme_a.get("cells", []):
        context = cell.get("context") or {}
        print(
            f"  row={cell.get('row')} col={cell.get('col')} "
            f"label={cell.get('label')!r} ok={cell.get('ok')} "
            f"reason={cell.get('reason')} path={cell.get('cell_path')} "
            f"bounds={context.get('bounds')} role={context.get('role')!r} "
            f"name={context.get('name')!r} desc={context.get('description')!r} "
            f"text={context.get('text')!r} context={context.get('context')}"
        )

    scheme_b = report.get("scheme_b", {})
    print(f"\n方案 B usable={scheme_b.get('usable')} reason={scheme_b.get('reason')}")
    for header in scheme_b.get("headers", []):
        print(
            f"  matched={header.get('matched')} path={header.get('path')} "
            f"relation={header.get('relative_to_table')} near_table={header.get('near_table')} "
            f"bounds={header.get('bounds')} role={header.get('role')!r} "
            f"name={header.get('name')!r} desc={header.get('description')!r}"
        )


if __name__ == "__main__":
    raise SystemExit(main())
