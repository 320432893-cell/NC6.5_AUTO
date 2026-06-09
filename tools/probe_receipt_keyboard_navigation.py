import argparse
import ctypes
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_probe import AccessibleTableCellInfo  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402


VK_KEYS = {
    "Tab": 0x09,
    "Enter": 0x0D,
    "Escape": 0x1B,
    "Left": 0x25,
    "Up": 0x26,
    "Right": 0x27,
    "Down": 0x28,
    "F2": 0x71,
}


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Probe NC receipt detail-table keyboard navigation. "
            "No mouse, text input, save, JAB action, or full-window tree scan."
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument(
        "--keys",
        default="Tab,Tab,Tab",
        help="Comma-separated probe keys: Tab,Enter,Escape,Left,Right,Up,Down,F2.",
    )
    parser.add_argument(
        "--key-target", choices=("foreground", "table"), default="foreground"
    )
    parser.add_argument("--delay", type=float, default=0.45)
    parser.add_argument("--max-rows", type=int, default=3)
    parser.add_argument(
        "--focus-cell",
        default="0,0",
        help="0-based row,col to select/focus before probing keys.",
    )
    parser.add_argument(
        "--no-auto-focus",
        action="store_true",
        help="Do not activate NC or select/focus a table cell before sending keys.",
    )
    parser.add_argument("--json", action="store_true")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only read current selected cells; do not send any key.",
    )
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        report = probe_navigation(
            jab,
            keys=parse_keys(args.keys),
            key_target=args.key_target,
            delay=args.delay,
            max_rows=args.max_rows,
            focus_cell=parse_focus_cell(args.focus_cell),
            auto_focus=not args.no_auto_focus,
            dry_run=args.dry_run,
        )
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    else:
        print_text(report)
    return 0 if report.get("ok") else 1


def parse_keys(text):
    keys = []
    for part in str(text).split(","):
        key = part.strip()
        if not key:
            continue
        if key not in VK_KEYS:
            raise ValueError(f"unsupported key: {key}")
        keys.append(key)
    return keys


def parse_focus_cell(text):
    row_text, col_text = str(text).split(",", maxsplit=1)
    return int(row_text.strip()), int(col_text.strip())


def probe_navigation(
    jab,
    keys,
    key_target,
    delay,
    max_rows,
    focus_cell,
    auto_focus,
    dry_run,
):
    located = locate_receipt_body_table(jab, max_rows=max_rows)
    best = located.get("best")
    report = {
        "ok": False,
        "read_only": dry_run,
        "dry_run": dry_run,
        "keys": keys,
        "key_target": key_target,
        "auto_focus": auto_focus,
        "focus_cell": {"row": focus_cell[0], "col": focus_cell[1]},
        "located": summarize_located(located),
        "table": None,
        "focus": None,
        "steps": [],
        "analysis": {},
    }
    if not best:
        report["reason"] = "receipt body table not located"
        return report

    before = read_table_selection(jab, best)
    report["table"] = before.get("table")
    report["steps"].append({"step": "initial", "selection": before})
    if not before.get("ok"):
        report["reason"] = before.get("reason")
        return report
    if dry_run:
        report["ok"] = True
        report["analysis"] = analyze_steps(report["steps"])
        return report

    if auto_focus:
        focus = focus_table_cell(jab, best, focus_cell[0], focus_cell[1])
        report["focus"] = focus
        focused = read_table_selection(jab, best)
        report["steps"].append({"step": "after_focus", "selection": focused})
        if not focus.get("ok"):
            report["reason"] = focus.get("reason")
            report["analysis"] = analyze_steps(report["steps"])
            return report
        if not selection_is_exactly(focused, focus_cell[0], focus_cell[1]):
            report["reason"] = (
                "focused cell did not read back as the only selected cell"
            )
            report["analysis"] = analyze_steps(report["steps"])
            return report

    for index, key_name in enumerate(keys, start=1):
        send = guarded_press_key(best.get("window"), key_name, key_target)
        step = {"step": index, "key": key_name, "send": send}
        if not send.get("ok"):
            step["selection"] = read_table_selection(jab, best)
            report["steps"].append(step)
            report["reason"] = send.get("reason")
            report["analysis"] = analyze_steps(report["steps"])
            return report
        time.sleep(max(delay, 0))
        step["selection"] = read_table_selection(jab, best)
        report["steps"].append(step)
        if not step["selection"].get("ok"):
            report["reason"] = step["selection"].get("reason")
            report["analysis"] = analyze_steps(report["steps"])
            return report

    report["ok"] = True
    report["analysis"] = analyze_steps(report["steps"])
    return report


def focus_table_cell(jab, best, row, col):
    context, vm_id, owned, window_info = jab.find_context_by_path_once(
        best["path"],
        class_name=best["window"].get("class_name"),
        scope_hwnd=best["window"].get("hwnd"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {"ok": False, "reason": "table context not found by path"}
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {"ok": False, "reason": "getAccessibleTableInfo failed"}
        if row < 0 or row >= table_info.rowCount:
            return {
                "ok": False,
                "reason": f"focus row out of range: {row} >= {table_info.rowCount}",
            }
        if col < 0 or col >= table_info.columnCount:
            return {
                "ok": False,
                "reason": f"focus col out of range: {col} >= {table_info.columnCount}",
            }
        activate = activate_window(window_info.get("hwnd"))
        selection = select_child(jab, vm_id, context, table_info, row, col)
        cell_context = get_cell_context(jab, vm_id, context, row, col)
        focus_results = []
        if hasattr(jab.dll, "requestFocus"):
            focus_results.append(
                {"target": "table", "ok": bool(jab.dll.requestFocus(vm_id, context))}
            )
            time.sleep(0.1)
            if cell_context:
                focus_results.append(
                    {
                        "target": "cell",
                        "ok": bool(jab.dll.requestFocus(vm_id, cell_context)),
                    }
                )
                time.sleep(0.1)
        return {
            "ok": bool(activate.get("ok")) and bool(selection.get("ok")),
            "activate": activate,
            "selection": selection,
            "request_focus": focus_results,
            "window": window_info,
            "target": {"row": row, "col": col},
        }
    finally:
        jab.release_contexts(vm_id, owned)


def activate_window(hwnd):
    if sys.platform != "win32" or not hwnd:
        return {"ok": False, "reason": "must run under Windows Python with a hwnd"}
    user32 = ctypes.windll.user32
    hwnd = int(hwnd)
    root = user32.GetAncestor(hwnd, 2) or hwnd
    SW_RESTORE = 9
    user32.ShowWindow(root, SW_RESTORE)
    user32.BringWindowToTop(root)
    ok = bool(user32.SetForegroundWindow(root))
    time.sleep(0.25)
    foreground = read_window_info(user32.GetForegroundWindow())
    return {
        "ok": ok,
        "hwnd": hwnd,
        "root_hwnd": int(root),
        "foreground": foreground,
    }


def select_child(jab, vm_id, table_context, table_info, row, col):
    if not (
        hasattr(jab.dll, "clearAccessibleSelectionFromContext")
        and hasattr(jab.dll, "addAccessibleSelectionFromContext")
    ):
        return {"ok": False, "reason": "selection API unavailable"}
    child_index = row * table_info.columnCount + col
    jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
    jab.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)
    time.sleep(0.15)
    selected = jab.get_selected_child_indexes(
        vm_id, table_context, table_info.rowCount * table_info.columnCount
    )
    return {
        "ok": child_index in selected,
        "child_index": child_index,
        "selected_indexes": selected,
    }


def get_cell_context(jab, vm_id, table_context, row, col):
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


def selection_is_exactly(selection, row, col):
    selected = (selection.get("selected") if selection else None) or []
    return (
        len(selected) == 1
        and selected[0].get("row") == row
        and selected[0].get("col") == col
    )


def read_table_selection(jab, best):
    context, vm_id, owned, window_info = jab.find_context_by_path_once(
        best["path"],
        class_name=best["window"].get("class_name"),
        scope_hwnd=best["window"].get("hwnd"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "reason": "table context not found by path",
            "table": {"path": best.get("path"), "window": best.get("window")},
        }
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {
                "ok": False,
                "reason": "getAccessibleTableInfo failed",
                "table": {"path": best.get("path"), "window": window_info},
            }
        selected = []
        cells = []
        for row in range(table_info.rowCount):
            for col in range(table_info.columnCount):
                cell = read_one_cell(jab, vm_id, context, row, col)
                if col in {1, 3, 4, 5, 7, 11}:
                    cells.append(cell)
                if cell.get("is_selected"):
                    selected.append(cell)
        return {
            "ok": True,
            "table": {
                "path": best.get("path"),
                "window": window_info,
                "row_count": table_info.rowCount,
                "col_count": table_info.columnCount,
                "bounds": context_bounds(jab, vm_id, context),
            },
            "selected": selected,
            "key_cells": cells,
        }
    finally:
        jab.release_contexts(vm_id, owned)


def read_one_cell(jab, vm_id, table_context, row, col):
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
        "is_selected": bool(cell_info.isSelected),
        "text": None,
        "bounds": None,
    }
    if not ok or not cell_info.accessibleContext:
        return item
    info = jab.get_context_info(vm_id, cell_info.accessibleContext)
    if not info:
        return item
    item["text"] = info.name.strip() or info.description.strip()
    item["bounds"] = [info.x, info.y, info.width, info.height]
    return item


def context_bounds(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    return [info.x, info.y, info.width, info.height]


def guarded_press_key(table_window, key_name, key_target):
    if sys.platform != "win32":
        return {"ok": False, "reason": "must run under Windows Python"}
    if key_name not in VK_KEYS:
        return {"ok": False, "reason": f"unsupported key: {key_name}"}
    user32 = ctypes.windll.user32
    table_info = read_window_info(int((table_window or {}).get("hwnd") or 0))
    foreground_info = read_window_info(user32.GetForegroundWindow())
    if not table_info or not foreground_info:
        return {
            "ok": False,
            "reason": "failed to read foreground or table window",
            "table_window": table_info,
            "foreground": foreground_info,
        }
    same_root = (
        foreground_info.get("hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("root_hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("hwnd") == table_info.get("hwnd")
    )
    if not same_root:
        return {
            "ok": False,
            "reason": f"foreground is not the located NC receipt window; skipped {key_name}",
            "table_window": table_info,
            "foreground": foreground_info,
        }
    target_hwnd = (
        table_info["hwnd"] if key_target == "table" else foreground_info["hwnd"]
    )
    try:
        if key_target == "table":
            send_ok = post_virtual_key(target_hwnd, VK_KEYS[key_name])
            mode = f"PostMessage({key_name})"
        else:
            send_virtual_key(VK_KEYS[key_name])
            send_ok = True
            mode = f"SendInput({key_name})"
    except Exception as exc:
        return {
            "ok": False,
            "reason": f"{type(exc).__name__}: {exc}",
            "mode": f"{key_target}({key_name})",
            "key": key_name,
            "target_hwnd": target_hwnd,
            "table_window": table_info,
            "foreground": foreground_info,
        }
    return {
        "ok": bool(send_ok),
        "mode": mode,
        "key": key_name,
        "key_target": key_target,
        "target_hwnd": target_hwnd,
        "table_window": table_info,
        "foreground": foreground_info,
    }


def read_window_info(hwnd):
    if not hwnd or sys.platform != "win32":
        return None
    user32 = ctypes.windll.user32
    if not user32.IsWindow(hwnd):
        return None
    length = user32.GetWindowTextLengthW(hwnd)
    title_buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, title_buffer, length + 1)
    class_buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_buffer, 256)
    root = user32.GetAncestor(hwnd, 2)
    return {
        "hwnd": int(hwnd),
        "title": title_buffer.value,
        "class_name": class_buffer.value,
        "visible": bool(user32.IsWindowVisible(hwnd)),
        "root_hwnd": int(root) if root else None,
    }


def send_virtual_key(vk):
    down = INPUT()
    down.type = 1
    down.ki = KEYBDINPUT(vk, 0, 0, 0, None)
    send_input(down)
    up = INPUT()
    up.type = 1
    up.ki = KEYBDINPUT(vk, 0, 0x0002, 0, None)
    send_input(up)


def post_virtual_key(hwnd, vk):
    user32 = ctypes.windll.user32
    WM_KEYDOWN = 0x0100
    WM_KEYUP = 0x0101
    down_ok = bool(user32.PostMessageW(hwnd, WM_KEYDOWN, vk, 0))
    up_ok = bool(user32.PostMessageW(hwnd, WM_KEYUP, vk, 0))
    return down_ok and up_ok


def send_input(inp):
    ctypes.windll.kernel32.SetLastError(0)
    sent = ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
    if sent != 1:
        error_code = ctypes.windll.kernel32.GetLastError()
        raise RuntimeError(f"SendInput failed, error={error_code}")


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class INPUT(ctypes.Structure):
    class _INPUT_UNION(ctypes.Union):
        _fields_ = [
            ("mi", MOUSEINPUT),
            ("ki", KEYBDINPUT),
        ]

    _anonymous_ = ("union",)
    _fields_ = [
        ("type", ctypes.c_ulong),
        ("union", _INPUT_UNION),
    ]


def summarize_located(located):
    best = located.get("best")
    return {
        "best": {
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
        "candidate_count": len(located.get("candidates") or []),
    }


def analyze_steps(steps):
    selected_points = []
    for step in steps:
        selected = ((step.get("selection") or {}).get("selected")) or []
        selected_points.append(
            {
                "step": step.get("step"),
                "key": step.get("key"),
                "selected": [
                    {
                        "row": item.get("row"),
                        "col": item.get("col"),
                        "text": item.get("text"),
                    }
                    for item in selected
                ],
            }
        )
    deltas = []
    previous = None
    for point in selected_points:
        current = point["selected"][0] if len(point["selected"]) == 1 else None
        if previous and current:
            deltas.append(
                {
                    "key": point.get("key"),
                    "from": previous,
                    "to": current,
                    "delta_row": current["row"] - previous["row"],
                    "delta_col": current["col"] - previous["col"],
                }
            )
        previous = current
    return {
        "selected_points": selected_points,
        "deltas": deltas,
        "single_selected_every_step": all(
            len(point["selected"]) == 1 for point in selected_points
        ),
    }


def print_text(report):
    print("NC 收款单明细表键盘导航试探")
    print("保证：不点鼠标、不输入文本、不保存、不扫描全窗口控件树。")
    print(f"ok={report.get('ok')} dry_run={report.get('dry_run')}")
    if report.get("reason"):
        print(f"reason={report.get('reason')}")
    table = report.get("table") or {}
    if table:
        print(
            f"table rows={table.get('row_count')} cols={table.get('col_count')} "
            f"bounds={table.get('bounds')} window={table.get('window')}"
        )
    for step in report.get("steps") or []:
        selection = step.get("selection") or {}
        selected = selection.get("selected") or []
        print(f"step={step.get('step')} key={step.get('key')} send={step.get('send')}")
        print(
            "  selected="
            + ", ".join(
                f"r{item.get('row')}c{item.get('col')}={item.get('text')!r}"
                for item in selected
            )
        )
    print(f"analysis={json.dumps(report.get('analysis'), ensure_ascii=False)}")


if __name__ == "__main__":
    raise SystemExit(main())
