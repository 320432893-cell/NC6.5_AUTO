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
from tools.receipt_body_table_locator import (  # noqa: E402
    locate_receipt_body_table,
    table_bounds,
)


MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Probe NC receipt detail row click x-ranges. It activates NC, single-clicks "
            "points across the visible detail row, and reads selected table index. "
            "No text input, paste, Enter, or double-click is sent."
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--max-rows", type=int, default=3)
    parser.add_argument("--activate-wait", type=float, default=0.2)
    parser.add_argument("--click-wait", type=float, default=0.12)
    parser.add_argument("--step", type=int, default=25)
    parser.add_argument("--x-start", type=int, default=None)
    parser.add_argument("--x-end", type=int, default=None)
    parser.add_argument("--y", type=int, default=None)
    parser.add_argument(
        "--restore-cursor",
        action="store_true",
        help="Move mouse cursor back to its original position after scanning.",
    )
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        report = run_probe(
            jab,
            max_rows=args.max_rows,
            activate_wait=max(0.0, args.activate_wait),
            click_wait=max(0.0, args.click_wait),
            step=max(1, args.step),
            x_start=args.x_start,
            x_end=args.x_end,
            y=args.y,
            restore_cursor=args.restore_cursor,
        )
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    else:
        print_text(report)
    return 0 if report.get("ok") else 1


def run_probe(
    jab,
    max_rows,
    activate_wait,
    click_wait,
    step,
    x_start,
    x_end,
    y,
    restore_cursor,
):
    report = {
        "ok": False,
        "read_write_note": "Mouse single-click only; no keyboard/text input is sent.",
        "requested": {
            "max_rows": max_rows,
            "activate_wait": activate_wait,
            "click_wait": click_wait,
            "step": step,
            "x_start": x_start,
            "x_end": x_end,
            "y": y,
            "restore_cursor": restore_cursor,
        },
        "platform": os.name,
        "foreground_before": safe_foreground(jab),
        "activation": None,
        "foreground_after_activation": None,
        "table": None,
        "scan": [],
        "selected_index_ranges": [],
        "reason": None,
    }
    if os.name != "nt":
        report["reason"] = "必须在 Windows Python 下运行"
        return report

    located = locate_receipt_body_table(jab, max_rows=max_rows)
    best = located.get("best")
    if not best:
        report["reason"] = "未定位到收款单明细表"
        return report

    report["activation"] = activate_nc_window(best.get("window") or {})
    if activate_wait:
        time.sleep(activate_wait)
    report["foreground_after_activation"] = safe_foreground(jab)

    located = locate_receipt_body_table(jab, max_rows=max_rows)
    best = located.get("best") or best
    table = describe_table(jab, best)
    report["table"] = table
    bounds = table.get("bounds")
    if not bounds_valid(bounds):
        report["reason"] = f"明细表 bounds 不可用：{bounds}"
        return report

    screen_width = int(ctypes.windll.user32.GetSystemMetrics(0))
    screen_height = int(ctypes.windll.user32.GetSystemMetrics(1))
    left, top, width, height = [int(value) for value in bounds]
    click_y = int(y) if y is not None else int(top + height // 2)
    start = int(x_start) if x_start is not None else max(0, left)
    end = int(x_end) if x_end is not None else min(screen_width - 1, left + width - 1)
    if start > end:
        report["reason"] = f"x 扫描范围无效：{start}>{end}"
        return report
    if click_y < 0 or click_y >= screen_height:
        report["reason"] = f"y 坐标超出屏幕：{click_y}"
        return report

    original_cursor = get_cursor_pos()
    try:
        for x in range(start, end + 1, step):
            point = {"x": int(x), "y": int(click_y)}
            click = single_click(point["x"], point["y"])
            if click_wait:
                time.sleep(click_wait)
            selection = read_selection(jab, best)
            report["scan"].append(
                {
                    "point": point,
                    "click": click,
                    "selection": selection,
                    "selected_index": first_selected_index(selection),
                    "selected_col": selected_col(selection),
                }
            )
    finally:
        if restore_cursor and original_cursor:
            ctypes.windll.user32.SetCursorPos(original_cursor[0], original_cursor[1])

    report["selected_index_ranges"] = group_selected_ranges(report["scan"])
    report["ok"] = True
    return report


def describe_table(jab, best):
    result = {
        "path": best.get("path"),
        "window": best.get("window"),
        "row_count": best.get("row_count"),
        "col_count": best.get("col_count"),
        "bounds": best.get("bounds"),
        "resolved_bounds": None,
        "selected_indexes": [],
    }
    window = best.get("window") or {}
    context, vm_id, owned, _window_info = jab.find_context_by_path_once(
        best.get("path"),
        class_name=window.get("class_name"),
        scope_hwnd=window.get("hwnd"),
        role="table",
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return result
    try:
        table_info = jab.get_table_info(vm_id, context)
        resolved = table_bounds(jab, vm_id, context)
        result["resolved_bounds"] = resolved
        result["bounds"] = resolved or result.get("bounds")
        if table_info:
            result["row_count"] = int(table_info.rowCount)
            result["col_count"] = int(table_info.columnCount)
            result["selected_indexes"] = jab.get_selected_child_indexes(
                vm_id,
                context,
                int(table_info.rowCount) * int(table_info.columnCount),
            )
    finally:
        jab.release_contexts(vm_id, owned)
    return result


def read_selection(jab, best):
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
        return {"ok": False, "reason": "按 path 重新取得明细表 context 失败"}
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {"ok": False, "reason": "getAccessibleTableInfo 失败"}
        selected = jab.get_selected_child_indexes(
            vm_id,
            context,
            int(table_info.rowCount) * int(table_info.columnCount),
        )
        return {
            "ok": True,
            "window": window_info,
            "row_count": int(table_info.rowCount),
            "col_count": int(table_info.columnCount),
            "selected_indexes": selected,
        }
    finally:
        jab.release_contexts(vm_id, owned)


def first_selected_index(selection):
    selected = (selection or {}).get("selected_indexes") or []
    return int(selected[0]) if selected else None


def selected_col(selection):
    selected = first_selected_index(selection)
    col_count = int((selection or {}).get("col_count") or 0)
    if selected is None or col_count <= 0:
        return None
    return int(selected % col_count)


def group_selected_ranges(scan):
    ranges = []
    current = None
    for item in scan:
        selected = item.get("selected_col")
        x = int((item.get("point") or {}).get("x"))
        if current and current.get("selected_col") == selected:
            current["x_end"] = x
            current["points"] += 1
            continue
        if current:
            ranges.append(current)
        current = {
            "selected_col": selected,
            "selected_index": item.get("selected_index"),
            "x_start": x,
            "x_end": x,
            "points": 1,
        }
    if current:
        ranges.append(current)
    return ranges


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


def single_click(x, y):
    try:
        user32 = ctypes.windll.user32
        user32.SetCursorPos(int(x), int(y))
        time.sleep(0.02)
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.03)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        return {"ok": True}
    except Exception as exc:
        return {"ok": False, "reason": f"{type(exc).__name__}: {exc}"}


def get_cursor_pos():
    if os.name != "nt":
        return None
    point = ctypes.wintypes.POINT()
    if not ctypes.windll.user32.GetCursorPos(ctypes.byref(point)):
        return None
    return [int(point.x), int(point.y)]


def safe_foreground(jab):
    try:
        return jab.get_foreground_window_info()
    except Exception as exc:
        return {"ok": False, "reason": f"{type(exc).__name__}: {exc}"}


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
    print("NC 收款单明细点击 x 扫描探测")
    print(f"  ok={report.get('ok')} reason={report.get('reason')}")
    print(f"  table={report.get('table')}")
    print("  selected index ranges:")
    for item in report.get("selected_index_ranges") or []:
        print(
            "    "
            f"x={item.get('x_start')}-{item.get('x_end')} "
            f"selected_col={item.get('selected_col')} "
            f"selected_index={item.get('selected_index')} "
            f"points={item.get('points')}"
        )
    print("  first scan points:")
    for item in (report.get("scan") or [])[:20]:
        print(
            "    "
            f"x={(item.get('point') or {}).get('x')} "
            f"y={(item.get('point') or {}).get('y')} "
            f"selected_col={item.get('selected_col')} "
            f"selected={((item.get('selection') or {}).get('selected_indexes'))}"
        )


if __name__ == "__main__":
    raise SystemExit(main())
