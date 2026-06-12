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
from tools.jab_probe import AccessibleActions, AccessibleTableCellInfo  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description="Dump NC receipt table cell contexts.")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--row", type=int, default=0)
    parser.add_argument("--cols", type=int, default=25)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        jab.ensure_started()
        report = dump_row(jab, args.row, args.cols)
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)
    return 0 if report.get("table") else 1


def dump_row(jab, row, col_limit):
    located = locate_receipt_body_table(jab, max_rows=3)
    best = located.get("best")
    if not best:
        return {"table": None, "located": located}

    context, vm_id, owned, window_info = jab.find_context_by_path_once(
        best["path"],
        class_name=best["window"].get("class_name"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {"table": best, "error": "table context not found"}

    try:
        table_info = jab.get_table_info(vm_id, context)
        cells = []
        for col in range(min(col_limit, table_info.columnCount)):
            cells.append(dump_cell(jab, vm_id, context, row, col))
        return {
            "table": {
                "path": best["path"],
                "window": window_info,
                "row_count": table_info.rowCount,
                "col_count": table_info.columnCount,
            },
            "cells": cells,
        }
    finally:
        jab.release_contexts(vm_id, owned)


def dump_cell(jab, vm_id, table_context, row, col):
    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    item = {
        "row": row,
        "col": col,
        "ok": bool(ok),
        "index": cell_info.index,
        "rowExtent": cell_info.rowExtent,
        "columnExtent": cell_info.columnExtent,
        "isSelected": bool(cell_info.isSelected),
        "context": None,
    }
    if not ok or not cell_info.accessibleContext:
        return item

    info = jab.get_context_info(vm_id, cell_info.accessibleContext)
    if not info:
        return item

    item["context"] = {
        "role": info.role_en_US.strip() or info.role.strip(),
        "name": info.name.strip(),
        "description": info.description.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "bounds": [info.x, info.y, info.width, info.height],
        "children": info.childrenCount,
        "accessibleText": bool(info.accessibleText),
        "textValue": jab.get_text_context_value(vm_id, cell_info.accessibleContext),
        "accessibleAction": bool(info.accessibleAction),
        "actions": get_action_names(jab, vm_id, cell_info.accessibleContext),
    }
    return item


def get_action_names(jab, vm_id, context):
    if not hasattr(jab.dll, "getAccessibleActions"):
        return []
    actions = AccessibleActions()
    if not jab.dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
        return []
    return [
        actions.actionInfo[index].name.strip() for index in range(actions.actionsCount)
    ]


def print_text(report):
    print(json.dumps(report.get("table"), ensure_ascii=False))
    for cell in report.get("cells", []):
        context = cell.get("context") or {}
        print(
            f"col={cell['col']} idx={cell['index']} selected={cell['isSelected']} "
            f"role={context.get('role')!r} name={context.get('name')!r} "
            f"desc={context.get('description')!r} text={context.get('textValue')!r} "
            f"bounds={context.get('bounds')} actions={context.get('actions')}"
        )


if __name__ == "__main__":
    raise SystemExit(main())
