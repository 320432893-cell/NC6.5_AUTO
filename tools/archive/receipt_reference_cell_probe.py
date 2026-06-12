import argparse
import ctypes
import json
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_probe import (  # noqa: E402
    AccessibleActions,
    AccessibleTableCellInfo,
    JOBJECT,
    enum_windows,
)
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402


INTERESTING_ROLES = {
    "combo box",
    "list",
    "list item",
    "menu",
    "menu item",
    "page tab",
    "push button",
    "table",
    "text",
    "toggle button",
    "tree",
}


def build_parser():
    parser = argparse.ArgumentParser(
        description="Probe popup/reference controls opened by an NC receipt body cell."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--row", type=int, required=True)
    parser.add_argument("--col", type=int, required=True)
    parser.add_argument(
        "--key",
        action="append",
        choices=("enter", "f2", "space", "tab", "down"),
        default=[],
        help="Semantic key used to trigger cell editing; may be passed multiple times.",
    )
    parser.add_argument("--wait", type=float, default=0.8)
    parser.add_argument("--max-depth", type=int, default=8)
    parser.add_argument("--max-children", type=int, default=120)
    parser.add_argument("--json", action="store_true")
    parser.add_argument(
        "--all-windows",
        action="store_true",
        help="Print all Java/AWT windows after triggering, not only new/changed ones.",
    )
    return parser


def main():
    args = build_parser().parse_args()
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        before = collect_java_windows(
            jab,
            max_depth=args.max_depth,
            max_children=args.max_children,
        )
        selection = select_receipt_body_cell(jab, args.row, args.col)
        key_results = []
        if selection.get("ok"):
            key_results = press_keys(jab, args.key, args.wait)
        time.sleep(args.wait)
        after = collect_java_windows(
            jab,
            max_depth=args.max_depth,
            max_children=args.max_children,
        )
        tables = jab.read_all_table_cells(max_rows=5, max_cols=25)
    finally:
        jab.hide_blank_awt_windows_enabled = False
        jab.close()

    report = {
        "target": {"row": args.row, "col": args.col, "keys": args.key},
        "selection": selection,
        "key_results": key_results,
        "new_or_changed_windows": diff_windows(before, after),
        "windows_after": after if args.all_windows else [],
        "tables_after": tables,
    }
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)
    return 0 if selection.get("ok") else 1


def select_receipt_body_cell(jab, row, col):
    if not jab.has_selection_api():
        return {"ok": False, "reason": "selection API unavailable"}

    located = locate_receipt_body_table(jab, max_rows=4)
    best = located.get("best") or fallback_receipt_body_candidate(located)
    if not best:
        return {
            "ok": False,
            "reason": "receipt body table not found",
            "located": located,
        }

    result = jab.find_context_by_path_once(
        best["path"],
        class_name=best["window"].get("class_name"),
        require_showing=False,
        require_valid_bounds=False,
    )
    table_context, vm_id, owned, window_info = result
    if not table_context:
        return {"ok": False, "reason": "table path no longer resolves", "best": best}

    cell_context = None
    try:
        table_info = jab.get_table_info(vm_id, table_context)
        if not table_info:
            return {"ok": False, "reason": "table info unavailable", "best": best}
        if row < 0 or row >= table_info.rowCount:
            return {
                "ok": False,
                "reason": "row out of range",
                "rows": table_info.rowCount,
            }
        if col < 0 or col >= table_info.columnCount:
            return {
                "ok": False,
                "reason": "col out of range",
                "cols": table_info.columnCount,
            }

        child_index = row * table_info.columnCount + col
        before_text = jab.get_table_cell_text(vm_id, table_context, row, col)
        before_selected = jab.get_selected_child_indexes(
            vm_id,
            table_context,
            table_info.rowCount * table_info.columnCount,
        )
        activate_hwnd(window_info.get("hwnd"))
        jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
        jab.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)
        if hasattr(jab.dll, "requestFocus"):
            jab.dll.requestFocus(vm_id, table_context)
            cell_context = get_table_cell_context(jab, vm_id, table_context, row, col)
            if cell_context:
                jab.dll.requestFocus(vm_id, cell_context)
        time.sleep(0.2)
        after_selected = jab.get_selected_child_indexes(
            vm_id,
            table_context,
            table_info.rowCount * table_info.columnCount,
        )
        after_text = jab.get_table_cell_text(vm_id, table_context, row, col)
        return {
            "ok": child_index in after_selected,
            "table": {
                "path": best["path"],
                "row_count": table_info.rowCount,
                "col_count": table_info.columnCount,
                "window": window_info,
            },
            "child_index": child_index,
            "cell_text_before": before_text,
            "cell_text_after_selection": after_text,
            "selected_before": before_selected,
            "selected_after": after_selected,
        }
    finally:
        if cell_context:
            jab.release_contexts(vm_id, [cell_context])
        jab.release_contexts(vm_id, owned)


def fallback_receipt_body_candidate(located):
    for candidate in located.get("candidates", []):
        if candidate.get("col_count") == 25 and candidate.get("row_count", 0) >= 1:
            return candidate
    return None


def activate_hwnd(hwnd):
    if not hwnd or os.name != "nt":
        return False
    user32 = ctypes.windll.user32
    user32.ShowWindow(hwnd, 9)
    return bool(user32.SetForegroundWindow(hwnd))


def get_table_cell_context(jab, vm_id, table_context, row, col):
    if not hasattr(jab.dll, "getAccessibleTableCellInfo"):
        return None
    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    if not ok or not cell_info.accessibleContext:
        return None
    return cell_info.accessibleContext


def press_keys(jab, keys, wait):
    if not keys:
        return []
    if os.name != "nt":
        return [
            {"key": key, "ok": False, "reason": "requires Windows Python"}
            for key in keys
        ]

    results = []
    for key in keys:
        jab.press_key(key, wait=wait)
        results.append({"key": key, "ok": True})
    return results


def collect_java_windows(jab, max_depth=8, max_children=120):
    windows = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if not class_name.startswith(("SunAwt", "Yonyou")):
            continue
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
            "root": summarize_context(jab, vm_id.value, root_context.value, "0"),
            "controls": [],
        }
        collect_controls(
            jab,
            vm_id.value,
            root_context.value,
            "0",
            window["controls"],
            depth=0,
            max_depth=max_depth,
            max_children=max_children,
        )
        windows.append(window)
    return windows


def collect_controls(
    jab, vm_id, context, path, controls, depth, max_depth, max_children
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    item = summarize_info(jab, vm_id, context, info, path)
    role = item["role"].lower()
    if should_keep_control(item, role):
        controls.append(item)

    if depth >= max_depth or role == "table":
        return

    for index in range(min(info.childrenCount, max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        collect_controls(
            jab,
            vm_id,
            child,
            f"{path}.{index}",
            controls,
            depth + 1,
            max_depth,
            max_children,
        )
        jab.release_contexts(vm_id, [child])


def summarize_context(jab, vm_id, context, path):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    return summarize_info(jab, vm_id, context, info, path)


def summarize_info(jab, vm_id, context, info, path):
    role = info.role_en_US.strip() or info.role.strip()
    states = info.states_en_US.strip() or info.states.strip()
    item = {
        "path": path,
        "role": role,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "states": states,
        "children": info.childrenCount,
        "bounds": [info.x, info.y, info.width, info.height],
        "accessibleAction": bool(info.accessibleAction),
        "accessibleSelection": bool(info.accessibleSelection),
        "accessibleText": bool(info.accessibleText),
        "actions": [],
        "text_value": None,
        "table": None,
    }
    if info.accessibleAction:
        item["actions"] = get_action_names(jab, vm_id, context)
    if info.accessibleText or role.lower() == "text":
        item["text_value"] = jab.get_text_context_value(vm_id, context)
    if role.lower() == "table":
        table_info = jab.get_table_info(vm_id, context)
        if table_info:
            item["table"] = {
                "rows": table_info.rowCount,
                "cols": table_info.columnCount,
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


def should_keep_control(item, role):
    if role in INTERESTING_ROLES:
        return True
    if (
        item["accessibleAction"]
        or item["accessibleText"]
        or item["accessibleSelection"]
    ):
        return True
    if item["name"] or item["description"]:
        return True
    return False


def diff_windows(before, after):
    before_by_key = {window_key(item): window_signature(item) for item in before}
    changed = []
    for item in after:
        key = window_key(item)
        signature = window_signature(item)
        if key not in before_by_key or before_by_key[key] != signature:
            changed.append(item)
    return changed


def window_key(window):
    return (
        window.get("hwnd"),
        window.get("class_name"),
        window.get("title"),
    )


def window_signature(window):
    controls = []
    for item in window.get("controls", []):
        controls.append(
            (
                item.get("path"),
                item.get("role"),
                item.get("name"),
                item.get("description"),
                item.get("states"),
                tuple(item.get("actions") or []),
                item.get("text_value"),
                json.dumps(item.get("table"), sort_keys=True),
            )
        )
    return (
        json.dumps(window.get("root"), ensure_ascii=False, sort_keys=True),
        tuple(controls),
    )


def print_text(report):
    target = report["target"]
    print(f"target: row={target['row']} col={target['col']} keys={target['keys']}")
    print("selection:", json.dumps(report["selection"], ensure_ascii=False))
    print("key_results:", json.dumps(report["key_results"], ensure_ascii=False))
    print("new_or_changed_windows:", len(report["new_or_changed_windows"]))
    for window in report["new_or_changed_windows"]:
        print(
            f"  window hwnd={window['hwnd']} class={window['class_name']!r} "
            f"title={window['title']!r} visible={window['visible']} "
            f"root={json.dumps(window['root'], ensure_ascii=False)}"
        )
        for item in window["controls"][:80]:
            print(
                f"    path={item['path']} role={item['role']!r} "
                f"name={item['name']!r} desc={item['description']!r} "
                f"states={item['states']!r} actions={item['actions']} "
                f"text={item['text_value']!r} table={item['table']} "
                f"bounds={item['bounds']}"
            )
    print("tables_after:", len(report["tables_after"]))
    for table in report["tables_after"]:
        print(
            f"  table={table['table_index']} win={table['window_title']!r}/"
            f"{table['window_class']!r} rows={table['row_count']} cols={table['col_count']}"
        )
        for row in table["rows"][:3]:
            print(
                f"    row {row['row_index']} selected={row['selected']}: {row['cells']}"
            )


if __name__ == "__main__":
    raise SystemExit(main())
