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
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402
from core.jab_probe import AccessibleTableCellInfo, JOBJECT, enum_windows  # noqa: E402


def build_parser():
    parser = argparse.ArgumentParser(
        description="Probe NC receipt body table cell selection/editability via JAB."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--table-index", type=int, default=0)
    parser.add_argument("--locate-body-table", action="store_true")
    parser.add_argument("--row", type=int, required=True)
    parser.add_argument("--col", type=int, required=True)
    parser.add_argument("--window-title", default=None)
    parser.add_argument("--window-class", default="SunAwtCanvas")
    parser.add_argument("--key", choices=("enter", "f2"), default=None)
    parser.add_argument("--paste-text", default=None)
    parser.add_argument("--type-text", default=None)
    parser.add_argument("--set-cell-text", default=None)
    parser.add_argument(
        "--commit-key",
        choices=("enter", "tab", "none"),
        default="enter",
        help="Key pressed after --paste-text. Use 'none' to skip.",
    )
    parser.add_argument(
        "--focus-target",
        choices=("cell", "table", "cell-then-table"),
        default="cell-then-table",
    )
    parser.add_argument("--wait", type=float, default=0.5)
    parser.add_argument("--json", action="store_true")
    return parser


def main():
    args = build_parser().parse_args()
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        jab.ensure_started()
        before_texts = collect_text_controls(jab, args.window_title, args.window_class)
        selected = select_cell(
            jab,
            args.table_index,
            args.row,
            args.col,
            args.window_title,
            locate_body_table=args.locate_body_table,
            paste_text=args.paste_text,
            type_text=args.type_text,
            set_cell_text=args.set_cell_text,
            commit_key=args.commit_key,
            focus_target=args.focus_target,
            wait=args.wait,
        )
        if args.key:
            jab.press_key(args.key, wait=args.wait)
        elif args.paste_text is None:
            time.sleep(args.wait)
        after_texts = collect_text_controls(jab, args.window_title, args.window_class)
        tables = jab.read_all_table_cells(max_rows=5, max_cols=25)
        report = {
            "target": {
                "table_index": args.table_index,
                "row": args.row,
                "col": args.col,
                "key": args.key,
                "paste_text": args.paste_text,
                "type_text": args.type_text,
                "set_cell_text": args.set_cell_text,
                "commit_key": args.commit_key if args.paste_text is not None else None,
                "focus_target": args.focus_target,
            },
            "selection": selected,
            "new_or_changed_text_controls": diff_text_controls(
                before_texts, after_texts
            ),
            "text_controls_after": after_texts[:80],
            "tables": tables,
        }
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)
    return 0 if report["selection"].get("ok") else 1


def select_cell(
    jab,
    table_index,
    row,
    col,
    window_title=None,
    locate_body_table=False,
    paste_text=None,
    type_text=None,
    set_cell_text=None,
    commit_key="enter",
    focus_target="cell-then-table",
    wait=0.5,
):
    if not jab.has_selection_api():
        return {"ok": False, "reason": "selection API unavailable"}

    tables = list(iter_target_tables(jab, locate_body_table))
    for current_index, (
        table_context,
        vm_id,
        owned,
        table_info,
        window_info,
    ) in enumerate(tables):
        try:
            if not locate_body_table and current_index != table_index:
                continue
            if window_title is not None and window_info.get("title") != window_title:
                return {
                    "ok": False,
                    "reason": "window title mismatch",
                    "window": window_info,
                }
            if row < 0 or row >= table_info.rowCount:
                return {
                    "ok": False,
                    "reason": "row out of range",
                    "row_count": table_info.rowCount,
                }
            if col < 0 or col >= table_info.columnCount:
                return {
                    "ok": False,
                    "reason": "col out of range",
                    "col_count": table_info.columnCount,
                }

            child_index = row * table_info.columnCount + col
            before = jab.get_selected_child_indexes(
                vm_id, table_context, table_info.rowCount * table_info.columnCount
            )
            before_text = jab.get_table_cell_text(vm_id, table_context, row, col)
            jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
            jab.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)
            after = jab.get_selected_child_indexes(
                vm_id, table_context, table_info.rowCount * table_info.columnCount
            )
            after_text = jab.get_table_cell_text(vm_id, table_context, row, col)
            edit = None
            if set_cell_text is not None and child_index in after:
                edit = set_selected_cell_text_direct(
                    jab,
                    vm_id,
                    table_context,
                    row,
                    col,
                    set_cell_text,
                )
                after_text = jab.get_table_cell_text(vm_id, table_context, row, col)
            elif (
                paste_text is not None or type_text is not None
            ) and child_index in after:
                edit = input_into_selected_cell(
                    jab,
                    vm_id,
                    table_context,
                    table_info,
                    row,
                    col,
                    window_info,
                    paste_text,
                    type_text,
                    commit_key,
                    focus_target,
                    wait,
                )
                after_text = jab.get_table_cell_text(vm_id, table_context, row, col)
            return {
                "ok": child_index in after,
                "table_index": current_index,
                "window": window_info,
                "row_count": table_info.rowCount,
                "col_count": table_info.columnCount,
                "child_index": child_index,
                "selected_before": before,
                "selected_after": after,
                "cell_text_before": before_text,
                "cell_text_after": after_text,
                "edit": edit,
            }
        finally:
            jab.release_contexts(vm_id, owned)

    return {"ok": False, "reason": "table not found", "table_count": len(tables)}


def set_selected_cell_text_direct(jab, vm_id, table_context, row, col, text):
    cell_context = get_table_cell_context(jab, vm_id, table_context, row, col)
    if not cell_context:
        return {"ok": False, "reason": "cell context unavailable"}
    try:
        before = jab.get_text_context_value(vm_id, cell_context)
        ok = jab.set_text_context(vm_id, cell_context, text)
        after = jab.get_text_context_value(vm_id, cell_context)
        return {"ok": bool(ok), "text_before": before, "text_after": after}
    finally:
        jab.release_contexts(vm_id, [cell_context])


def iter_target_tables(jab, locate_body_table=False):
    if not locate_body_table:
        yield from jab.find_tables_once()
        return

    report = locate_receipt_body_table(jab, max_rows=3)
    best = report.get("best")
    if not best:
        return
    result = jab.find_context_by_path_once(
        best["path"],
        class_name=best["window"].get("class_name"),
        require_showing=False,
        require_valid_bounds=False,
    )
    context, vm_id, owned, window_info = result
    if not context:
        return
    table_info = jab.get_table_info(vm_id, context)
    if not table_info:
        jab.release_contexts(vm_id, owned)
        return
    yield context, vm_id, owned, table_info, window_info


def input_into_selected_cell(
    jab,
    vm_id,
    table_context,
    table_info,
    row,
    col,
    window_info,
    paste_text,
    type_text,
    commit_key,
    focus_target,
    wait,
):
    if os.name != "nt":
        return {"ok": False, "reason": "paste probe must run under Windows Python"}

    focus_results = []
    activate_main = jab.activate_window_by_title(
        "Yonyou UClient", class_name="YonyouUWnd", timeout=2
    )
    activate_result = activate_hwnd(window_info.get("hwnd"))
    cell_context = get_table_cell_context(jab, vm_id, table_context, row, col)
    try:
        if focus_target in ("cell", "cell-then-table") and cell_context:
            focus_results.append(
                {
                    "target": "cell",
                    "ok": bool(jab.dll.requestFocus(vm_id, cell_context)),
                }
            )
            time.sleep(0.1)
        if focus_target in ("table", "cell-then-table"):
            focus_results.append(
                {
                    "target": "table",
                    "ok": bool(jab.dll.requestFocus(vm_id, table_context)),
                }
            )
            time.sleep(0.1)

        foreground_before = jab.get_foreground_window_info()
        old_clipboard = safe_clipboard_read(jab)
        if type_text is not None:
            jab.type_text(type_text, interval=0.01)
        else:
            jab.clipboard_copy(paste_text)
            jab.clipboard_paste(wait=0.0)
        time.sleep(wait)
        if commit_key and commit_key != "none":
            jab.press_key(commit_key, wait=0.0)
            time.sleep(wait)
        foreground_after = jab.get_foreground_window_info()
        restore_clipboard(jab, old_clipboard)
        return {
            "ok": True,
            "activate_main": activate_main,
            "activate_window": activate_result,
            "focus": focus_results,
            "foreground_before": foreground_before,
            "foreground_after": foreground_after,
        }
    finally:
        if cell_context:
            jab.release_contexts(vm_id, [cell_context])


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


def activate_hwnd(hwnd):
    if not hwnd or os.name != "nt":
        return False
    user32 = ctypes.windll.user32
    user32.ShowWindow(hwnd, 9)
    return bool(user32.SetForegroundWindow(hwnd))


def safe_clipboard_read(jab):
    try:
        return jab.clipboard_read()
    except Exception:
        return None


def restore_clipboard(jab, value):
    if value is None:
        return
    try:
        jab.clipboard_copy(value)
    except Exception:
        pass


def collect_text_controls(jab, window_title=None, window_class=None):
    controls = []
    seen = set()
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if window_title is not None and title != window_title:
            continue
        if window_class is not None and class_name != window_class:
            continue
        if not visible or not jab.dll.isJavaWindow(hwnd):
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vm_id), ctypes.byref(root_context)
        ):
            continue
        collect_text_controls_in_tree(
            jab,
            vm_id.value,
            root_context.value,
            [],
            {"hwnd": int(hwnd), "title": title, "class": class_name, "pid": pid},
            controls,
            seen,
            0,
        )
    return controls


def collect_text_controls_in_tree(
    jab, vm_id, context, path, window, controls, seen, depth
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    role = info.role_en_US.strip() or info.role.strip()
    role_l = role.lower()
    states = info.states_en_US.strip() or info.states.strip()
    states_l = states.lower()
    path_text = "0" + "".join(f".{index}" for index in path)

    if info.accessibleText or role_l == "text":
        value = jab.get_text_context_value(vm_id, context)
        item = {
            "path": path_text,
            "role": role,
            "name": info.name.strip(),
            "description": info.description.strip(),
            "states": states,
            "showing": "visible" in states_l and "showing" in states_l,
            "accessibleText": bool(info.accessibleText),
            "value": value,
            "bounds": [info.x, info.y, info.width, info.height],
            "window": window,
        }
        key = (
            item["path"],
            item["role"],
            item["name"],
            item["description"],
            item["value"],
            window["hwnd"],
        )
        if key not in seen:
            seen.add(key)
            controls.append(item)

    if depth >= jab.max_depth or role_l == "table":
        return

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        collect_text_controls_in_tree(
            jab, vm_id, child, path + [index], window, controls, seen, depth + 1
        )
        jab.release_contexts(vm_id, [child])


def diff_text_controls(before, after):
    before_by_path = {item["path"]: item for item in before}
    changed = []
    for item in after:
        previous = before_by_path.get(item["path"])
        if previous is None or comparable_text(item) != comparable_text(previous):
            changed.append(item)
    return changed[:80]


def comparable_text(item):
    return (
        item.get("role"),
        item.get("name"),
        item.get("description"),
        item.get("states"),
        item.get("value"),
        item.get("bounds"),
    )


def print_text(report):
    target = report["target"]
    selection = report["selection"]
    print(
        "target: "
        f"table={target['table_index']} row={target['row']} col={target['col']} "
        f"key={target['key']}"
    )
    print(f"selection: {json.dumps(selection, ensure_ascii=False)}")
    print("new_or_changed_text_controls:", len(report["new_or_changed_text_controls"]))
    for item in report["new_or_changed_text_controls"][:20]:
        print(
            f"  path={item['path']} role={item['role']!r} name={item['name']!r} "
            f"desc={item['description']!r} showing={item['showing']} "
            f"value={item['value']!r} bounds={item['bounds']} "
            f"win={item['window']['title']!r}/{item['window']['class']!r}"
        )
    print("tables:", len(report["tables"]))
    for table in report["tables"]:
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
