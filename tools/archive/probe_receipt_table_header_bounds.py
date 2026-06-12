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
from tools.jab_probe import AccessibleTableCellInfo, JOBJECT, enum_windows  # noqa: E402
from tools.receipt_body_table_locator import (  # noqa: E402
    locate_receipt_body_table,
)


TARGET_COLUMNS = {
    1: "收款业务类型",
    3: "币种",
    4: "收款银行账户",
    5: "科目",
    7: "贷方原币金额",
    11: "结算方式",
}


INTERESTING_ROLES = {
    "label",
    "text",
    "table cell",
    "table column header",
    "column header",
    "panel",
    "unknown",
}


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Read-only probe for NC receipt detail table header/semantic bounds. "
            "No keyboard, mouse, JAB actions, or text writes are sent."
        )
    )
    parser.add_argument("--config", default=str(ROOT / "config.json"))
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--max-rows", type=int, default=3)
    parser.add_argument(
        "--scan-window",
        action="store_true",
        help="Also scan the Java window control tree. Disabled by default.",
    )
    parser.add_argument("--max-controls", type=int, default=120)
    parser.add_argument("--max-depth", type=int, default=8)
    parser.add_argument("--max-children", type=int, default=80)
    parser.add_argument("--max-visited", type=int, default=1200)
    parser.add_argument(
        "--include-all-text",
        action="store_true",
        help="Include every visible text-like control, not only target matches.",
    )
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        report = probe_header_bounds(
            jab,
            max_rows=args.max_rows,
            scan_window=args.scan_window,
            max_controls=args.max_controls,
            max_depth=args.max_depth,
            max_children=args.max_children,
            max_visited=args.max_visited,
            include_all_text=args.include_all_text,
        )
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    else:
        print_text(report)
    return 0 if report.get("best_table") else 1


def probe_header_bounds(
    jab,
    max_rows=3,
    scan_window=False,
    max_controls=120,
    max_depth=8,
    max_children=80,
    max_visited=1200,
    include_all_text=False,
):
    located = locate_receipt_body_table(jab, max_rows=max_rows)
    best = located.get("best")
    report = {
        "readonly": True,
        "note": "This probe only reads JAB context/table/text information.",
        "targets": TARGET_COLUMNS,
        "best_table": best,
        "table_candidates": summarize_candidates(located.get("candidates") or []),
        "target_cell_bounds": {},
        "window_scan": {
            "enabled": scan_window,
            "max_depth": max_depth,
            "max_children": max_children,
            "max_visited": max_visited,
            "visited": 0,
            "kept": 0,
            "stop_reason": None,
        },
        "text_candidates": [],
        "target_matches": {},
        "column_inference": {},
        "judgement": {},
    }
    if not best:
        report["judgement"] = {
            "scheme_b_feasible": False,
            "reason": "No receipt body table candidate was found.",
        }
        return report

    table_context, vm_id, owned, window_info = jab.find_context_by_path_once(
        best["path"],
        class_name=best["window"].get("class_name"),
        scope_hwnd=best["window"].get("hwnd"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not table_context:
        report["judgement"] = {
            "scheme_b_feasible": False,
            "reason": "Best table path could not be resolved again.",
        }
        return report

    try:
        table_info = jab.get_table_info(vm_id, table_context)
        table_bounds = context_bounds(jab, vm_id, table_context)
        report["best_table"] = {
            **best,
            "resolved_window": window_info,
            "resolved_bounds": table_bounds,
            "resolved_row_count": table_info.rowCount if table_info else None,
            "resolved_col_count": table_info.columnCount if table_info else None,
        }
        report["target_cell_bounds"] = read_target_cell_bounds(
            jab, vm_id, table_context, table_info
        )
    finally:
        jab.release_contexts(vm_id, owned)

    if scan_window:
        controls, scan_stats = collect_text_controls_for_window(
            jab,
            best["window"],
            max_controls=max_controls,
            max_depth=max_depth,
            max_children=max_children,
            max_visited=max_visited,
        )
        report["window_scan"].update(scan_stats)
    else:
        controls = []
        report["window_scan"]["stop_reason"] = "disabled; pass --scan-window to enable"
    controls = annotate_controls(controls, report["best_table"].get("resolved_bounds"))
    target_matches = match_targets(controls)
    report["target_matches"] = target_matches
    report["column_inference"] = infer_columns_from_matches(
        target_matches,
        report["target_cell_bounds"],
        report["best_table"].get("resolved_bounds"),
    )
    report["text_candidates"] = select_report_controls(
        controls,
        include_all_text=include_all_text,
    )
    report["judgement"] = judge_scheme_b(report)
    return report


def summarize_candidates(candidates):
    result = []
    for item in candidates:
        result.append(
            {
                "table_index": item.get("table_index"),
                "path": item.get("path"),
                "window": item.get("window"),
                "row_count": item.get("row_count"),
                "col_count": item.get("col_count"),
                "score": item.get("score"),
                "reasons": item.get("reasons"),
                "bounds": item.get("bounds"),
                "selected_indexes": item.get("selected_indexes"),
                "rows": item.get("rows"),
            }
        )
    return result


def read_target_cell_bounds(jab, vm_id, table_context, table_info):
    result = {}
    if not table_info or not hasattr(jab.dll, "getAccessibleTableCellInfo"):
        return result
    for col, label in TARGET_COLUMNS.items():
        if col >= table_info.columnCount or table_info.rowCount <= 0:
            result[str(col)] = {"label": label, "ok": False}
            continue
        result[str(col)] = {
            "label": label,
            "row0": read_cell_context(jab, vm_id, table_context, 0, col),
            "row1": read_cell_context(jab, vm_id, table_context, 1, col)
            if table_info.rowCount > 1
            else None,
        }
    return result


def read_cell_context(jab, vm_id, table_context, row, col):
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
    item["context"] = describe_context(jab, vm_id, cell_info.accessibleContext)
    return item


def collect_text_controls_for_window(
    jab,
    table_window,
    max_controls=120,
    max_depth=8,
    max_children=80,
    max_visited=1200,
):
    stats = {
        "visited": 0,
        "kept": 0,
        "stop_reason": None,
    }
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if int(hwnd) != int(table_window.get("hwnd")):
            continue
        if not jab.dll.isJavaWindow(hwnd):
            stats["stop_reason"] = "target hwnd is not a Java window"
            return [], stats
        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            stats["stop_reason"] = "getAccessibleContextFromHWND failed"
            return [], stats
        window = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
        }
        result = []
        owned = [root_context.value]
        try:
            collect_text_controls_in_tree(
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
                max_children=max_children,
                max_visited=max_visited,
                stats=stats,
            )
        finally:
            jab.release_contexts(vm_id.value, owned)
        stats["kept"] = len(result)
        return result, stats
    stats["stop_reason"] = "target hwnd not enumerated"
    return [], stats


def collect_text_controls_in_tree(
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
    max_children,
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

    role = normalize_role(info)
    text_value = (
        jab.get_text_context_value(vm_id, context) if info.accessibleText else None
    )
    item = {
        "path": path,
        "window": window,
        **describe_info(info, text_value=text_value),
    }
    if should_keep_control(item, role):
        result.append(item)

    if depth >= max_depth:
        stats["stop_reason"] = stats["stop_reason"] or "max_depth reached"
        return
    if role == "table":
        return

    for index in range(min(info.childrenCount, max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        owned.append(child)
        collect_text_controls_in_tree(
            jab,
            vm_id,
            child,
            path=f"{path}.{index}",
            depth=depth + 1,
            owned=owned,
            result=result,
            window=window,
            max_controls=max_controls,
            max_depth=max_depth,
            max_children=max_children,
            max_visited=max_visited,
            stats=stats,
        )
        if len(result) >= max_controls:
            stats["stop_reason"] = "max_controls reached"
            return
        if stats["visited"] >= max_visited:
            stats["stop_reason"] = "max_visited reached"
            return


def should_keep_control(item, role):
    combined = normalize_text(
        " ".join(
            str(item.get(key) or "")
            for key in ("name", "description", "text", "role", "states")
        )
    )
    if any(normalize_text(label) in combined for label in TARGET_COLUMNS.values()):
        return True
    if role in INTERESTING_ROLES and (
        item.get("name") or item.get("description") or item.get("text")
    ):
        return True
    return False


def annotate_controls(controls, table_bounds):
    annotated = []
    table_rect = rect_from_bounds(table_bounds)
    for item in controls:
        rect = rect_from_bounds(item.get("bounds"))
        relation = relate_to_table(rect, table_rect)
        annotated.append({**item, "table_relation": relation})
    annotated.sort(
        key=lambda control: (
            -int(control["table_relation"].get("near_table_header_zone", False)),
            control["bounds"][1] if control.get("bounds") else 999999,
            control["bounds"][0] if control.get("bounds") else 999999,
            control.get("path") or "",
        )
    )
    return annotated


def match_targets(controls):
    result = {str(col): [] for col in TARGET_COLUMNS}
    for col, label in TARGET_COLUMNS.items():
        needle = normalize_text(label)
        for control in controls:
            haystack = normalize_text(
                " ".join(
                    str(control.get(key) or "")
                    for key in ("name", "description", "text")
                )
            )
            if needle not in haystack:
                continue
            relation = control.get("table_relation") or {}
            score = 0
            if relation.get("near_table_header_zone"):
                score += 5
            if relation.get("horizontally_overlaps_table"):
                score += 2
            if relation.get("above_or_same_top"):
                score += 2
            role = normalize_text(control.get("role"))
            if "header" in role or "label" in role:
                score += 1
            result[str(col)].append({**control, "match_score": score})
        result[str(col)].sort(
            key=lambda item: (
                -item.get("match_score", 0),
                item["bounds"][1] if item.get("bounds") else 999999,
                item["bounds"][0] if item.get("bounds") else 999999,
            )
        )
    return result


def infer_columns_from_matches(target_matches, target_cell_bounds, table_bounds):
    result = {}
    table_rect = rect_from_bounds(table_bounds)
    for col, label in TARGET_COLUMNS.items():
        matches = target_matches.get(str(col)) or []
        best = matches[0] if matches else None
        header_range = x_range_from_bounds(best.get("bounds")) if best else None
        cell_range = cell_x_range(target_cell_bounds.get(str(col)))
        inferred = header_range or cell_range
        source = "header" if header_range else "cell" if cell_range else None
        confidence = "none"
        if header_range and cell_range:
            confidence = (
                "high" if x_ranges_overlap(header_range, cell_range) else "conflict"
            )
        elif header_range:
            relation = best.get("table_relation") or {}
            confidence = "medium" if relation.get("near_table_header_zone") else "low"
        elif cell_range:
            confidence = "cell-only"
        result[str(col)] = {
            "label": label,
            "source": source,
            "confidence": confidence,
            "header_x_range": header_range,
            "cell_x_range": cell_range,
            "inferred_x_range": clamp_x_range(inferred, table_rect),
            "best_header_path": best.get("path") if best else None,
            "best_header_bounds": best.get("bounds") if best else None,
            "best_header_role": best.get("role") if best else None,
            "best_header_name": best.get("name") if best else None,
            "best_header_description": best.get("description") if best else None,
            "best_header_text": best.get("text") if best else None,
            "match_count": len(matches),
        }
    return result


def judge_scheme_b(report):
    inferred = report.get("column_inference") or {}
    header_hits = [
        item
        for item in inferred.values()
        if item.get("source") == "header"
        and item.get("confidence") in ("high", "medium")
    ]
    conflicts = [
        item for item in inferred.values() if item.get("confidence") == "conflict"
    ]
    if conflicts:
        return {
            "scheme_b_feasible": False,
            "reason": "Some header-derived x ranges conflict with row/col cell bounds.",
            "header_hit_count": len(header_hits),
            "conflict_count": len(conflicts),
        }
    if len(header_hits) >= 4:
        return {
            "scheme_b_feasible": True,
            "reason": "Most target columns have visible header/semantic controls near the table.",
            "header_hit_count": len(header_hits),
            "conflict_count": 0,
        }
    return {
        "scheme_b_feasible": False,
        "reason": "Too few target headers were found near the receipt detail table.",
        "header_hit_count": len(header_hits),
        "conflict_count": 0,
    }


def select_report_controls(controls, include_all_text=False):
    if include_all_text:
        return controls
    selected = []
    for control in controls:
        combined = normalize_text(
            " ".join(
                str(control.get(key) or "") for key in ("name", "description", "text")
            )
        )
        relation = control.get("table_relation") or {}
        if any(normalize_text(label) in combined for label in TARGET_COLUMNS.values()):
            selected.append(control)
        elif relation.get("near_table_header_zone") and len(selected) < 120:
            selected.append(control)
    return selected[:160]


def describe_context(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    text_value = (
        jab.get_text_context_value(vm_id, context) if info.accessibleText else None
    )
    return describe_info(info, text_value=text_value)


def describe_info(info, text_value=None):
    return {
        "role": normalize_role(info),
        "name": info.name.strip(),
        "description": info.description.strip(),
        "text": text_value,
        "states": info.states_en_US.strip() or info.states.strip(),
        "bounds": [info.x, info.y, info.width, info.height],
        "children": info.childrenCount,
        "accessibleText": bool(info.accessibleText),
        "accessibleAction": bool(info.accessibleAction),
        "accessibleSelection": bool(info.accessibleSelection),
    }


def context_bounds(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    return [info.x, info.y, info.width, info.height]


def normalize_role(info):
    if isinstance(info, str):
        return info.strip()
    return info.role_en_US.strip() or info.role.strip()


def normalize_text(value):
    return str(value or "").strip().replace(" ", "").replace("　", "").lower()


def rect_from_bounds(bounds):
    if not bounds or len(bounds) != 4:
        return None
    x, y, width, height = bounds
    if width <= 0 or height <= 0:
        return None
    return {
        "left": x,
        "top": y,
        "right": x + width,
        "bottom": y + height,
        "width": width,
        "height": height,
        "mid_x": x + width / 2,
        "mid_y": y + height / 2,
    }


def relate_to_table(rect, table_rect):
    if not rect or not table_rect:
        return {
            "has_valid_bounds": bool(rect),
            "near_table_header_zone": False,
        }
    horizontal_overlap = max(
        0,
        min(rect["right"], table_rect["right"]) - max(rect["left"], table_rect["left"]),
    )
    vertical_gap_to_top = table_rect["top"] - rect["bottom"]
    same_top_distance = abs(rect["top"] - table_rect["top"])
    near_header = (
        horizontal_overlap > 0
        and -8 <= vertical_gap_to_top <= 80
        or horizontal_overlap > 0
        and same_top_distance <= 28
    )
    return {
        "has_valid_bounds": True,
        "horizontally_overlaps_table": horizontal_overlap > 0,
        "horizontal_overlap": horizontal_overlap,
        "vertical_gap_to_table_top": vertical_gap_to_top,
        "same_top_distance": same_top_distance,
        "above_or_same_top": rect["mid_y"] <= table_rect["mid_y"],
        "near_table_header_zone": bool(near_header),
    }


def x_range_from_bounds(bounds):
    rect = rect_from_bounds(bounds)
    if not rect:
        return None
    return [rect["left"], rect["right"]]


def cell_x_range(cell_entry):
    if not cell_entry:
        return None
    for key in ("row0", "row1"):
        cell = cell_entry.get(key)
        context = cell.get("context") if cell else None
        x_range = x_range_from_bounds(context.get("bounds") if context else None)
        if x_range:
            return x_range
    return None


def x_ranges_overlap(left, right):
    if not left or not right:
        return False
    return min(left[1], right[1]) > max(left[0], right[0])


def clamp_x_range(x_range, table_rect):
    if not x_range or not table_rect:
        return x_range
    return [max(x_range[0], table_rect["left"]), min(x_range[1], table_rect["right"])]


def print_text(report):
    print("NC 收款单明细表列标题 bounds 只读探测")
    print("只读保证：不发送鼠标、键盘、JAB action 或 setTextContents。")
    print()

    best = report.get("best_table")
    if not best:
        print("best table: <none>")
        print(f"判断：{report.get('judgement')}")
        return

    print("best table:")
    window = best.get("window") or {}
    print(
        f"  path={best.get('path')} rows={best.get('resolved_row_count') or best.get('row_count')} "
        f"cols={best.get('resolved_col_count') or best.get('col_count')} "
        f"score={best.get('score')} reasons={best.get('reasons')}"
    )
    print(
        f"  window hwnd={window.get('hwnd')} class={window.get('class_name')} "
        f"title={window.get('title') or '<无标题>'}"
    )
    print(f"  bounds={best.get('resolved_bounds') or best.get('bounds')}")
    print()

    scan = report.get("window_scan") or {}
    print(
        "window scan: "
        f"enabled={scan.get('enabled')} visited={scan.get('visited')} "
        f"kept={scan.get('kept')} max_depth={scan.get('max_depth')} "
        f"max_children={scan.get('max_children')} "
        f"max_visited={scan.get('max_visited')} "
        f"stop={scan.get('stop_reason')}"
    )
    print()

    print("table candidates:")
    for item in report.get("table_candidates") or []:
        print(
            f"  table={item.get('table_index')} path={item.get('path')} "
            f"rows={item.get('row_count')} cols={item.get('col_count')} "
            f"score={item.get('score')} bounds={item.get('bounds')} "
            f"reasons={item.get('reasons')}"
        )
    print()

    print("target columns:")
    inference = report.get("column_inference") or {}
    target_cell_bounds = report.get("target_cell_bounds") or {}
    for col, label in TARGET_COLUMNS.items():
        item = inference.get(str(col)) or {}
        print(
            f"  col {col} {label}: source={item.get('source')} "
            f"confidence={item.get('confidence')} "
            f"inferred_x={item.get('inferred_x_range')} "
            f"header_x={item.get('header_x_range')} cell_x={item.get('cell_x_range')}"
        )
        if item.get("best_header_path"):
            print(
                f"    header path={item.get('best_header_path')} "
                f"role={item.get('best_header_role')!r} "
                f"name={item.get('best_header_name')!r} "
                f"desc={item.get('best_header_description')!r} "
                f"text={item.get('best_header_text')!r} "
                f"bounds={item.get('best_header_bounds')} "
                f"matches={item.get('match_count')}"
            )
        cell_entry = target_cell_bounds.get(str(col)) or {}
        for key in ("row0", "row1"):
            cell = cell_entry.get(key)
            if not cell:
                continue
            context = cell.get("context") or {}
            print(
                f"    {key} cell ok={cell.get('ok')} idx={cell.get('index')} "
                f"role={context.get('role')!r} name={context.get('name')!r} "
                f"desc={context.get('description')!r} text={context.get('text')!r} "
                f"bounds={context.get('bounds')}"
            )
    print()

    print("matched header/text candidates:")
    for control in report.get("text_candidates") or []:
        relation = control.get("table_relation") or {}
        print(
            f"  path={control.get('path')} role={control.get('role')!r} "
            f"name={control.get('name')!r} desc={control.get('description')!r} "
            f"text={control.get('text')!r} bounds={control.get('bounds')} "
            f"nearHeader={relation.get('near_table_header_zone')} "
            f"gapTop={relation.get('vertical_gap_to_table_top')}"
        )
    print()
    print(f"判断：{report.get('judgement')}")


if __name__ == "__main__":
    raise SystemExit(main())
