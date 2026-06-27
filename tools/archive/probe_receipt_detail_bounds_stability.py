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
from core.jab_probe import AccessibleTableCellInfo  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402


COLUMN_LABELS = {
    1: "收款业务类型",
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
            "Read-only stability probe for NC receipt detail table cell bounds. "
            "It can optionally select/requestFocus cells, but never types, pastes, "
            "clicks, or writes values."
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--rows", default="0", help="Comma-separated 0-based rows.")
    parser.add_argument(
        "--cols",
        default="4,5,7,11",
        help="Comma-separated 0-based columns.",
    )
    parser.add_argument("--repeat", type=int, default=5)
    parser.add_argument("--interval", type=float, default=0.25)
    parser.add_argument("--max-rows", type=int, default=3)
    parser.add_argument("--scope-hwnd", type=int, default=None)
    parser.add_argument(
        "--focus",
        action="store_true",
        help="For each target cell, use table selection API and requestFocus, then reread bounds.",
    )
    parser.add_argument("--focus-wait", type=float, default=0.08)
    parser.add_argument(
        "--locate-each-repeat",
        action="store_true",
        help="Run semantic table location on every repeat. Slower, but detects path/window changes.",
    )
    parser.add_argument(
        "--activate-window-before-read",
        action="store_true",
        help="Activate the located NC table window before every sample, then reacquire context and bounds.",
    )
    parser.add_argument("--activate-wait", type=float, default=0.2)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    rows = parse_indexes(args.rows)
    cols = parse_indexes(args.cols)
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        report = run_probe(
            jab,
            rows=rows,
            cols=cols,
            repeat=max(1, args.repeat),
            interval=max(0.0, args.interval),
            max_rows=args.max_rows,
            scope_hwnd=args.scope_hwnd,
            focus=args.focus,
            focus_wait=max(0.0, args.focus_wait),
            locate_each_repeat=args.locate_each_repeat,
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


def parse_indexes(text):
    result = []
    for part in str(text or "").split(","):
        part = part.strip()
        if part:
            result.append(int(part))
    return result


def run_probe(
    jab,
    rows,
    cols,
    repeat,
    interval,
    max_rows,
    scope_hwnd,
    focus,
    focus_wait,
    locate_each_repeat,
    activate_window_before_read,
    activate_wait,
):
    started_at = time.time()
    located = locate_receipt_body_table(
        jab,
        max_rows=max_rows,
        scope_hwnd=scope_hwnd,
    )
    best = located.get("best")
    report = {
        "read_only": True,
        "started_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(started_at)),
        "requested": {
            "rows": rows,
            "cols": cols,
            "repeat": repeat,
            "interval": interval,
            "focus": focus,
            "focus_wait": focus_wait,
            "locate_each_repeat": locate_each_repeat,
            "activate_window_before_read": activate_window_before_read,
            "activate_wait": activate_wait,
            "scope_hwnd": scope_hwnd,
        },
        "best_table": summarize_best(best),
        "candidate_count": len(located.get("candidates") or []),
        "samples": [],
        "summary": {},
    }
    if not best:
        report["summary"] = {"reason": "未定位到收款单明细表"}
        return report

    for index in range(repeat):
        if index > 0 and interval:
            time.sleep(interval)
        current_best = best
        if locate_each_repeat:
            current_located = locate_receipt_body_table(
                jab,
                max_rows=max_rows,
                scope_hwnd=scope_hwnd,
            )
            current_best = current_located.get("best")
        sample = read_sample(
            jab,
            current_best,
            rows=rows,
            cols=cols,
            focus=focus,
            focus_wait=focus_wait,
            activate_window_before_read=activate_window_before_read,
            activate_wait=activate_wait,
            sample_index=index + 1,
            started_at=started_at,
        )
        report["samples"].append(sample)

    report["summary"] = summarize_samples(report["samples"], rows, cols)
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
    }


def read_sample(
    jab,
    best,
    rows,
    cols,
    focus,
    focus_wait,
    activate_window_before_read,
    activate_wait,
    sample_index,
    started_at,
):
    sample = {
        "sample": sample_index,
        "offset_seconds": round(time.time() - started_at, 3),
        "ok": False,
        "foreground": safe_foreground(jab),
        "activation": None,
        "table": None,
        "cells": [],
        "reason": None,
    }
    if not best:
        sample["reason"] = "本次未定位到明细表"
        return sample

    if activate_window_before_read:
        sample["activation"] = activate_nc_window(best.get("window") or {})
        if activate_wait:
            time.sleep(activate_wait)
        sample["foreground_after_activation"] = safe_foreground(jab)

    window = best.get("window") or {}
    context, vm_id, owned, window_info = jab.find_context_by_path_once(
        best.get("path"),
        class_name=window.get("class_name"),
        scope_hwnd=window.get("hwnd"),
        role="table",
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        sample["reason"] = "按 path 重新取得明细表 context 失败"
        return sample

    try:
        table_info = jab.get_table_info(vm_id, context)
        table_context = describe_context(jab, vm_id, context)
        sample["table"] = {
            "path": best.get("path"),
            "window": window_info,
            "row_count": table_info.rowCount if table_info else None,
            "col_count": table_info.columnCount if table_info else None,
            "context": table_context,
            "selected_indexes": jab.get_selected_child_indexes(
                vm_id,
                context,
                table_info.rowCount * table_info.columnCount if table_info else 0,
            )
            if table_info
            else [],
        }
        if not table_info:
            sample["reason"] = "getAccessibleTableInfo 失败"
            return sample

        for row in rows:
            for col in cols:
                item = read_cell(jab, vm_id, context, table_info, row, col)
                if focus:
                    item["after_focus"] = focus_and_reread_cell(
                        jab,
                        vm_id,
                        context,
                        table_info,
                        row,
                        col,
                        focus_wait,
                    )
                item["estimated_from_table_bounds"] = estimate_cell_center(
                    (table_context or {}).get("bounds"),
                    table_info.rowCount,
                    table_info.columnCount,
                    row,
                    col,
                )
                sample["cells"].append(item)
        sample["ok"] = True
        return sample
    finally:
        jab.release_contexts(vm_id, owned)


def read_cell(jab, vm_id, table_context, table_info, row, col):
    item = {
        "row": row,
        "col": col,
        "label": COLUMN_LABELS.get(col),
        "ok": False,
        "reason": None,
        "cell_info": None,
        "context": None,
    }
    if row < 0 or row >= table_info.rowCount:
        item["reason"] = f"row out of range: {row} / {table_info.rowCount}"
        return item
    if col < 0 or col >= table_info.columnCount:
        item["reason"] = f"col out of range: {col} / {table_info.columnCount}"
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
    item["context"] = describe_context(jab, vm_id, cell_info.accessibleContext)
    item["ok"] = True
    return item


def focus_and_reread_cell(jab, vm_id, table_context, table_info, row, col, focus_wait):
    child_index = row * table_info.columnCount + col
    result = {
        "ok": False,
        "child_index": child_index,
        "selection": None,
        "request_focus": [],
        "cell": None,
        "reason": None,
    }
    if not jab.has_selection_api():
        result["reason"] = "JAB selection API 不可用"
        return result

    try:
        jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
        add_ok = bool(jab.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index))
        if focus_wait:
            time.sleep(focus_wait)
        selected = jab.get_selected_child_indexes(
            vm_id,
            table_context,
            table_info.rowCount * table_info.columnCount,
        )
        result["selection"] = {
            "add_ok": add_ok,
            "selected_indexes": selected,
            "target_selected": child_index in selected,
        }
        if hasattr(jab.dll, "requestFocus"):
            result["request_focus"].append(
                {"target": "table", "ok": bool(jab.dll.requestFocus(vm_id, table_context))}
            )
            if focus_wait:
                time.sleep(focus_wait)
        reread = read_cell(jab, vm_id, table_context, table_info, row, col)
        cell_context = ((reread.get("cell_info") or {}).get("accessible_context") or 0)
        if cell_context and hasattr(jab.dll, "requestFocus"):
            result["request_focus"].append(
                {"target": "cell", "ok": bool(jab.dll.requestFocus(vm_id, cell_context))}
            )
            if focus_wait:
                time.sleep(focus_wait)
            reread = read_cell(jab, vm_id, table_context, table_info, row, col)
        result["cell"] = reread
        result["ok"] = bool(reread.get("ok"))
        return result
    except Exception as exc:
        result["reason"] = f"{type(exc).__name__}: {exc}"
        return result


def describe_context(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    role = info.role_en_US.strip() or info.role.strip()
    states = info.states_en_US.strip() or info.states.strip()
    bounds = [info.x, info.y, info.width, info.height]
    return {
        "context": int(context),
        "role": role,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "states": states,
        "bounds": bounds,
        "valid_bounds": bounds_valid(bounds),
        "children": info.childrenCount,
        "accessible_text": bool(info.accessibleText),
        "text": jab.get_text_context_value(vm_id, context),
        "accessible_action": bool(info.accessibleAction),
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


def estimate_cell_center(table_bounds, row_count, col_count, row, col):
    if not bounds_valid(table_bounds) or row_count <= 0 or col_count <= 0:
        return {"ok": False, "reason": f"表格 bounds 不可用：{table_bounds}"}
    x, y, width, height = [int(value) for value in table_bounds]
    cell_width = width / float(col_count)
    cell_height = height / float(row_count)
    return {
        "ok": True,
        "table_bounds": list(table_bounds),
        "point": [
            round(x + (col + 0.5) * cell_width, 1),
            round(y + (row + 0.5) * cell_height, 1),
        ],
        "cell_width": round(cell_width, 3),
        "cell_height": round(cell_height, 3),
    }


def summarize_samples(samples, rows, cols):
    summary = {
        "target_cells": {},
        "valid_table_bounds_samples": 0,
        "sample_count": len(samples),
    }
    for sample in samples:
        table_bounds = (
            ((sample.get("table") or {}).get("context") or {}).get("bounds")
        )
        if bounds_valid(table_bounds):
            summary["valid_table_bounds_samples"] += 1
        for cell in sample.get("cells") or []:
            key = f"{cell.get('row')},{cell.get('col')}"
            entry = summary["target_cells"].setdefault(
                key,
                {
                    "row": cell.get("row"),
                    "col": cell.get("col"),
                    "label": cell.get("label"),
                    "read_count": 0,
                    "valid_before_focus": 0,
                    "valid_after_focus": 0,
                    "bounds_seen": [],
                    "states_seen": [],
                    "texts_seen": [],
                },
            )
            entry["read_count"] += 1
            context = cell.get("context") or {}
            remember(entry["bounds_seen"], context.get("bounds"))
            remember(entry["states_seen"], context.get("states"))
            remember(entry["texts_seen"], context.get("text") or context.get("description"))
            if bounds_valid(context.get("bounds")):
                entry["valid_before_focus"] += 1
            after_context = (
                (((cell.get("after_focus") or {}).get("cell") or {}).get("context"))
                or {}
            )
            if bounds_valid(after_context.get("bounds")):
                entry["valid_after_focus"] += 1
                remember(entry["bounds_seen"], after_context.get("bounds"))

    for row in rows:
        for col in cols:
            summary["target_cells"].setdefault(
                f"{row},{col}",
                {
                    "row": row,
                    "col": col,
                    "label": COLUMN_LABELS.get(col),
                    "read_count": 0,
                    "valid_before_focus": 0,
                    "valid_after_focus": 0,
                    "bounds_seen": [],
                    "states_seen": [],
                    "texts_seen": [],
                },
            )
    return summary


def remember(items, value):
    if value is None:
        return
    if value not in items:
        items.append(value)


def print_text(report):
    print("NC 收款单明细单元格 bounds 稳定性探测")
    print(f"  read_only={report.get('read_only')} requested={report.get('requested')}")
    best = report.get("best_table")
    if not best:
        print(f"  未定位到明细表：{report.get('summary')}")
        return
    print(
        "  best: "
        f"path={best.get('path')} rows={best.get('row_count')} "
        f"cols={best.get('col_count')} bounds={best.get('bounds')} "
        f"window={best.get('window')}"
    )
    summary = report.get("summary") or {}
    print(
        "  table bounds valid samples: "
        f"{summary.get('valid_table_bounds_samples')}/{summary.get('sample_count')}"
    )
    for key, item in sorted((summary.get("target_cells") or {}).items()):
        print(
            "  cell "
            f"{key} {item.get('label') or ''}: reads={item.get('read_count')} "
            f"valid_before={item.get('valid_before_focus')} "
            f"valid_after_focus={item.get('valid_after_focus')} "
            f"bounds_seen={item.get('bounds_seen')} "
            f"states_seen={item.get('states_seen')} "
            f"texts_seen={item.get('texts_seen')}"
        )


if __name__ == "__main__":
    raise SystemExit(main())
