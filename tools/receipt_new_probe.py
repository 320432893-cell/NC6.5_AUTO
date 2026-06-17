# 职责：收款单"新增→自制"开单引擎主入口——CLI/run 编排 + 前台/控件查找(被 monkeypatch)
# 不做什么：报告/win32/菜单/窗口枚举/按钮逻辑已拆到 tools.receipt_new_*、win32_window_utils
# 允许依赖层：core JAB、tools.jab_probe、tools.receipt_query_guard、tools.receipt_new_*
# 谁不应该 import：core 层模块不应 import 本入口

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
from tools.jab_probe import AccessibleActions, JOBJECT, enum_windows  # noqa: E402, F401
from tools.receipt_query_guard import (  # noqa: E402
    ReceiptPageGuardError,
    guard_receipt_parent_page,
)
from tools.receipt_new_report import (  # noqa: E402, F401
    get_action_names,
    print_text,
    summarize_action_report,
    summarize_candidates,
    summarize_context,
    summarize_control,
    summarize_info,
    summarize_report,
    summarize_target,
)
from tools.win32_window_utils import (  # noqa: E402, F401
    close_popup_hwnd,
    describe_hwnd,
    has_valid_bounds,
    is_visible_sun_awt_popup,
    root_hwnd,
    window_class_name,
    window_text,
)
from tools.receipt_new_menu import (  # noqa: E402, F401
    choose_click_action,
    choose_self_made_menu_item,
    detect_self_made_entry_state,
    do_action_by_window_path,
    is_current_visible_control,
    normalize_entry_state_names,
)
from tools.receipt_new_window_scan import (  # noqa: E402, F401
    annotate_foreground_root,
    annotate_foreground_root_for_targets,
    collect_controls,
    collect_entry_context_snapshot,
    collect_new_visible_popup_windows,
    collect_receipt_new_windows,
    collect_receipt_new_windows_compat,
    diff_windows,
    find_named_controls_in_windows,
    find_new_visible_popup,
    keep_control,
    resolve_current_canvas_header_anchor,
    window_key,
    window_signature,
)
from tools.receipt_new_button import (  # noqa: E402, F401
    filter_usable_new_buttons,
    find_new_buttons,
    foreground_nc_guard,
    new_button_priority,
    open_new_menu,
    open_new_menu_with_ctrl_n,
    open_new_menu_with_known_buttons,
    send_ctrl_n,
    trigger_button_async,
    wait_for_self_made_popup,
)


def elapsed(start):
    return round(time.perf_counter() - start, 3)



def measure(timings, name, func, *args, **kwargs):
    start = time.perf_counter()
    result = func(*args, **kwargs)
    timings.append({"name": name, "seconds": elapsed(start)})
    return result



def build_parser():
    parser = argparse.ArgumentParser(
        description="Open and inspect the NC receipt New/Self-made entry menu."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument(
        "--method",
        choices=("probe-button", "button", "action-path"),
        default="probe-button",
        help="How to open the New menu.",
    )
    parser.add_argument("--path", default=None)
    parser.add_argument("--title", default=None)
    parser.add_argument("--class-name", default=None)
    parser.add_argument("--name", default="新增")
    parser.add_argument("--role", default=None)
    parser.add_argument("--action", default=None)
    parser.add_argument("--return-timeout", type=float, default=0.2)
    parser.add_argument("--wait", type=float, default=0.8)
    parser.add_argument("--choose-self-made", action="store_true")
    parser.add_argument(
        "--self-made-index",
        type=int,
        default=None,
        help="Fallback menu item index when the menu item has no readable name.",
    )
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--summary", action="store_true")
    return parser



def main():
    args = build_parser().parse_args()
    if args.method == "action-path" and not args.path:
        raise SystemExit("--path is required with --method action-path")
    if args.choose_self_made and args.self_made_index is None:
        args.self_made_index = 0

    report = run(args)
    if args.summary:
        print(json.dumps(summarize_report(report), ensure_ascii=False, indent=2))
    elif args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)
    if args.choose_self_made:
        return (
            0
            if report.get("open", {}).get("ok")
            and report.get("choose_self_made", {}).get("ok")
            else 1
        )
    return 0 if report.get("open", {}).get("ok") else 1



def run(args, jab=None, before=None, buttons=None):
    timings = []
    owns_jab = jab is None
    cfg = None
    if owns_jab:
        cfg = load_config(args.config)
        jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        if owns_jab:
            measure(timings, "new-probe.jab.ensure-started", jab.ensure_started)
        else:
            timings.append({"name": "new-probe.jab.ensure-started", "seconds": 0.0})
        parent_guard = measure(
            timings,
            "new-probe.receipt-parent-guard",
            guard_receipt_new_parent_page,
            jab,
            cfg or getattr(jab, "config", {}) or {},
        )
        if not parent_guard.get("ok"):
            return {
                "foreground": foreground_info(),
                "matches": [],
                "buttons": [],
                "usable_buttons": [],
                "open": {
                    "ok": False,
                    "method": "receipt-parent-guard",
                    "reason": parent_guard.get("reason")
                    or "新增前未确认当前是收款单录入父页",
                },
                "receipt_parent_guard": parent_guard,
                "tracked_popup": None,
                "popup_cleanup": None,
                "new_or_changed_after_open": [],
                "windows_after_open": [],
                "choose_self_made": None,
                "windows_after_choose": None,
                "entry_state": {"ok": False, "reason": "新增前页面守卫失败"},
                "timings": timings,
            }
        if before is None:
            before = measure(
                timings, "new-probe.collect-before", collect_receipt_new_windows, jab
            )
        else:
            timings.append({"name": "new-probe.collect-before", "seconds": 0.0})
        foreground = measure(timings, "new-probe.foreground", foreground_info)
        annotate_foreground_root(before, foreground)
        matches = measure(
            timings,
            "new-probe.find-matches",
            find_named_controls_in_windows,
            before,
            args.name,
            args.role,
            args.class_name,
            require_action=False,
        )
        if buttons is None:
            buttons = measure(
                timings,
                "new-probe.find-buttons",
                find_named_controls_in_windows,
                before,
                args.name,
                args.role,
                args.class_name,
                require_action=True,
            )
        else:
            timings.append({"name": "new-probe.find-buttons", "seconds": 0.0})
        annotate_foreground_root_for_targets(buttons, foreground)
        usable_buttons = filter_usable_new_buttons(buttons, foreground)
        usable_buttons.sort(key=new_button_priority)
        initial_entry_state = measure(
            timings,
            "new-probe.detect-initial-entry-state",
            detect_self_made_entry_state,
            before,
        )
        if args.choose_self_made and initial_entry_state.get("ok"):
            open_report = {
                "ok": True,
                "method": "already-self-made-entry",
                "reason": "self-made entry state already visible before opening New",
            }
            popup_wait = {
                "ok": False,
                "reason": "self-made entry state already visible; popup not needed",
                "popup": None,
                "windows": before,
            }
        else:
            open_report = measure(
                timings,
                "new-probe.open-new-menu",
                open_new_menu_with_known_buttons,
                jab,
                args,
                usable_buttons,
                buttons,
                foreground,
            )
            popup_wait = measure(
                timings,
                "new-probe.wait-for-popup",
                wait_for_self_made_popup,
                jab,
                before,
                args.wait,
            )
        after_open = popup_wait.get("windows") or []
        tracked_popup = popup_wait.get("popup")
        choose_report = None
        after_choose = None
        entry_wait = None
        popup_cleanup = None
        if args.choose_self_made and initial_entry_state.get("ok"):
            choose_report = {
                "ok": True,
                "method": "already-self-made-entry",
                "reason": "self-made entry state already visible before opening New",
            }
            entry_wait = {
                "ok": True,
                "confirmed": True,
                "state": initial_entry_state,
                "windows": before,
            }
            after_choose = before
        elif args.choose_self_made and open_report.get("ok"):
            choose_report = measure(
                timings,
                "new-probe.choose-self-made",
                choose_self_made_menu_item,
                jab,
                after_open,
                args.self_made_index,
                popup_hwnd=tracked_popup.get("hwnd") if tracked_popup else None,
            )
            if choose_report.get("ok"):
                if entry_wait is None:
                    entry_wait = measure(
                        timings,
                        "new-probe.entry-context-snapshot",
                        collect_entry_context_snapshot,
                        jab,
                    )
                    after_choose = entry_wait.get("windows") or []
            if (
                tracked_popup
                and choose_report.get("ok")
                and (entry_wait or {}).get("ok")
            ):
                popup_cleanup = measure(
                    timings,
                    "new-probe.popup-cleanup",
                    close_popup_hwnd,
                    tracked_popup["hwnd"],
                )
        elif args.choose_self_made:
            choose_report = {
                "ok": False,
                "reason": "new menu was not opened; self-made selection skipped",
            }
        entry_state = (
            entry_wait.get("state")
            if entry_wait
            else measure(
                timings,
                "new-probe.detect-entry-state",
                detect_self_made_entry_state,
                after_choose or after_open,
            )
        )
    finally:
        jab.hide_blank_awt_windows_enabled = False
        if owns_jab:
            jab.close()

    report = {
        "foreground": foreground,
        "matches": matches,
        "buttons": buttons,
        "usable_buttons": usable_buttons,
        "open": open_report,
        "tracked_popup": tracked_popup,
        "popup_cleanup": popup_cleanup,
        "new_or_changed_after_open": diff_windows(before, after_open),
        "windows_after_open": after_open,
        "choose_self_made": choose_report,
        "windows_after_choose": after_choose,
        "entry_state": entry_state,
        "receipt_parent_guard": parent_guard,
        "timings": timings,
    }
    return report



def guard_receipt_new_parent_page(jab, config):
    query_cfg = ((config or {}).get("receipt_entry") or {}).get("query") or {}
    try:
        report = guard_receipt_parent_page(jab, config or {}, query_cfg)
    except ReceiptPageGuardError as exc:
        return {
            "ok": False,
            "enabled": True,
            "reason": str(exc),
        }
    return {
        **report,
        "purpose": "before-new-self-made",
    }



def foreground_info():
    if os.name != "nt" or not hasattr(ctypes, "windll"):
        return {}
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return {}
    root = root_hwnd(hwnd)
    return {
        "hwnd": int(hwnd),
        "root": root,
        "class_name": window_class_name(hwnd),
        "title": window_text(hwnd),
        "root_class_name": window_class_name(root),
        "root_title": window_text(root),
    }



def find_named_controls(
    jab,
    name_query="新增",
    role=None,
    class_name=None,
    require_action=True,
):
    results = []
    name_query = str(name_query or "").lower()
    role = role.lower() if role else None
    for window in collect_receipt_new_windows(jab, max_depth=25, max_children=1000):
        if class_name and window.get("class_name") != class_name:
            continue
        if not window.get("is_java"):
            continue
        for control in window.get("all_controls", []):
            control_role = control.get("role", "").lower()
            if role and control_role != role:
                continue
            text = f"{control.get('name', '')} {control.get('description', '')}".lower()
            if name_query and name_query not in text:
                continue
            if require_action and not control.get("accessibleAction"):
                continue
            results.append(
                {
                    "window": {
                        key: window.get(key)
                        for key in ("hwnd", "title", "class_name", "visible")
                    },
                    "control": control,
                }
            )
    return results
