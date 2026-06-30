import argparse
import ctypes
import json
import os
import sys
import threading
import time
from ctypes import wintypes
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from core.jab_probe import AccessibleActions, JOBJECT, enum_windows  # noqa: E402
from core.receipt_query_guard import (  # noqa: E402
    ReceiptPageGuardError,
    guard_receipt_parent_page,
)


SELF_MADE_NAMES = {"自制"}
ENTRY_STATE_NAMES = {"保存(Ctrl+S)", "暂存", "取消(Ctrl+Q)"}


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
        choices=("probe-button", "button"),
        default="probe-button",
        help="How to open the New menu.",
    )
    parser.add_argument("--class-name", default=None)
    parser.add_argument("--name", default="新增")
    parser.add_argument("--role", default=None)
    parser.add_argument("--action", default=None)
    parser.add_argument("--return-timeout", type=float, default=0.2)
    parser.add_argument("--wait", type=float, default=0.8)
    parser.add_argument("--choose-self-made", action="store_true")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--summary", action="store_true")
    return parser


def main():
    args = build_parser().parse_args()

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
            and report.get("entry_state", {}).get("ok")
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
                timings, "new-probe.collect-before", collect_receipt_new_baseline, jab
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
                popup_hwnd=tracked_popup.get("hwnd") if tracked_popup else None,
            )
            if tracked_popup and choose_report.get("ok"):
                popup_cleanup = measure(
                    timings,
                    "new-probe.popup-cleanup",
                    close_popup_hwnd,
                    tracked_popup["hwnd"],
                )
            if choose_report.get("ok"):
                entry_wait = measure(
                    timings,
                    "new-probe.wait-entry-ready",
                    wait_self_made_entry_ready,
                    jab,
                    popup_hwnd=tracked_popup.get("hwnd") if tracked_popup else None,
                    timeout=max(float(args.wait or 0), 0.8),
                    entry_button_target=open_report.get("target"),
                )
                after_choose = entry_wait.get("windows") or []
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
        "entry_context_snapshot": summarize_entry_context_snapshot(entry_wait),
        "entry_state": entry_state,
        "receipt_parent_guard": parent_guard,
        "timings": timings,
    }
    return report


def summarize_entry_context_snapshot(entry_wait):
    if not entry_wait:
        return None
    anchor = entry_wait.get("anchor") or {}
    quick_anchor = entry_wait.get("quick_anchor") or {}
    return {
        "ok": bool(entry_wait.get("ok")),
        "confirmed": bool(entry_wait.get("confirmed")),
        "method": entry_wait.get("method"),
        "entry_ready_source": "edit-button",
        "anchor_ok": bool(anchor.get("ok")),
        "anchor_reason": anchor.get("reason"),
        "header_scope_anchor_ok": bool(anchor.get("ok")),
        "scope_hwnd": anchor.get("scope_hwnd"),
        "dynamic_index": anchor.get("dynamic_index"),
        "quick_anchor_ok": bool(quick_anchor.get("ok")),
        "quick_anchor_reason": quick_anchor.get("reason"),
        "quick_anchor_candidate_count": len(quick_anchor.get("candidates") or []),
    }


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


def open_new_menu_with_known_buttons(
    jab, args, buttons, all_buttons=None, foreground=None
):
    if args.method == "button" and buttons:
        target = buttons[0]
        action_report = trigger_button_async(
            jab,
            target["window"]["hwnd"],
            target["control"]["path"],
            action_name=args.action,
            return_timeout=args.return_timeout,
            target=target,
        )
        if action_report.get("ok"):
            return action_report
        return {
            "ok": False,
            "method": "button",
            "reason": "新增按钮 action 失败；正式收款流程不回退 Ctrl+N",
            "button_action": action_report,
        }
    if args.method == "button" and all_buttons:
        return {
            "ok": False,
            "method": "button",
            "reason": "找到新增候选，但没有前台收款单窗口内可用的新增按钮；正式流程不回退 Ctrl+N",
            "button_reason": "new button candidates were found, but none were usable in the foreground NC window",
            "rejected_count": len(all_buttons),
            "rejected": summarize_candidates(all_buttons[:20]),
        }
    if args.method == "button":
        return {
            "ok": False,
            "method": "button",
            "reason": "未找到新增按钮；正式收款流程不回退 Ctrl+N",
            "button_reason": "new button not found",
            "rejected_count": 0,
            "rejected": [],
        }
    return open_new_menu(jab, args)


def open_new_menu(jab, args):
    if args.method == "probe-button":
        return {"ok": True, "method": "probe-button"}
    return {
        "ok": False,
        "method": args.method,
        "reason": "unsupported new-menu method",
    }


def wait_for_self_made_popup(jab, before, timeout=0.8, interval=0.08):
    start = time.perf_counter()
    deadline = time.perf_counter() + max(float(timeout or 0), 0)
    attempts = 0
    last_windows = []
    while True:
        attempts += 1
        last_windows = collect_new_visible_popup_windows(jab, before)
        popup = find_new_visible_popup(before, last_windows)
        if popup:
            return {
                "ok": True,
                "attempts": attempts,
                "wait_seconds": elapsed(start),
                "popup": popup,
                "windows": last_windows,
            }
        remaining = deadline - time.perf_counter()
        if remaining <= 0:
            break
        time.sleep(min(interval, remaining))
    full_windows = collect_receipt_new_windows(jab)
    popup = find_new_visible_popup(before, full_windows)
    return {
        "ok": bool(popup),
        "attempts": attempts,
        "wait_seconds": max(float(timeout or 0), 0),
        "popup": popup,
        "windows": full_windows if popup else last_windows,
    }


def wait_self_made_entry_ready(
    jab,
    popup_hwnd=None,
    timeout=0.8,
    interval=0.08,
    entry_button_target=None,
):
    start = time.perf_counter()
    deadline = start + max(float(timeout or 0), 0.0)
    interval = max(float(interval or 0.08), 0.02)
    attempts = []
    last = None
    while True:
        last = collect_entry_context_snapshot(
            jab,
            popup_hwnd=popup_hwnd,
            entry_button_target=entry_button_target,
        )
        attempts.append(summarize_entry_context_snapshot(last))
        state = last.get("state") or {}
        if last.get("confirmed") and state.get("ok") and not last.get("popup_visible"):
            return {
                **last,
                "ok": True,
                "wait_seconds": elapsed(start),
                "attempts": attempts,
            }
        remaining = deadline - time.perf_counter()
        if remaining <= 0:
            return {
                **(last or {}),
                "ok": False,
                "wait_seconds": elapsed(start),
                "attempts": attempts,
                "reason": ("自制菜单已点击，但未确认 popup 关闭且进入新建编辑态"),
            }
        time.sleep(min(interval, remaining))


def collect_entry_context_snapshot(jab, popup_hwnd=None, entry_button_target=None):
    foreground = foreground_info()
    canvas_window = find_foreground_canvas_window_light(jab, foreground)
    popup_state = describe_popup_visibility(popup_hwnd)
    popup_visible = bool(popup_state.get("visible")) if popup_hwnd else False
    edit_state = detect_entry_state_ready_light(
        jab,
        foreground,
        entry_button_target=entry_button_target,
    )
    parent_new_state = {
        "ok": False,
        "skipped": True,
        "method": "skip-parent-new-scan",
        "reason": "改用保存/暂存/取消任一编辑态按钮确认新建录入态",
    }
    entry_ready = bool(not popup_visible and edit_state.get("ok"))
    if popup_visible:
        state_reason = "自制菜单 popup 仍可见，暂不写入"
    elif not edit_state.get("ok"):
        state_reason = "未确认保存/暂存/取消任一编辑态按钮"
    else:
        state_reason = "popup 已关闭，编辑态按钮已出现"
    canvas_hit = []
    windows = []
    if canvas_window.get("ok"):
        window = canvas_window.get("window") or {}
        control = {
            "path": None,
            "role": "canvas",
            "name": "current SunAwtCanvas",
            "description": "当前前台收款单 canvas；表头 dynamic_index 由 header 阶段解析",
        }
        canvas_hit.append({"window": window, "control": control})
        windows.append(
            {
                **window,
                "is_java": True,
                "controls": [control],
                "all_controls": [control],
            }
        )
    state = {
        "ok": entry_ready,
        "edit_buttons_ok": bool(edit_state.get("ok")),
        "partial_ok": bool(edit_state.get("ok")),
        "names": edit_state.get("names") or [],
        "hits": canvas_hit + (edit_state.get("hits") or []),
        "reason": state_reason,
        "parent_new_state": parent_new_state,
        "edit_button_state": edit_state,
        "canvas_state": canvas_window,
    }
    return {
        "ok": True,
        "confirmed": entry_ready,
        "method": "edit-button-ready",
        "state": state,
        "windows": windows,
        "anchor": {},
        "quick_anchor": {},
        "foreground": foreground,
        "popup": popup_state,
        "popup_visible": popup_visible,
    }


def find_foreground_canvas_window_light(jab, foreground=None):
    foreground = foreground or foreground_info()
    fg_root = (foreground or {}).get("root")
    if os.name != "nt" or not hasattr(ctypes, "WinDLL") or not fg_root:
        return {
            "ok": False,
            "method": "foreground-canvas-window-light",
            "reason": "foreground root not available",
            "window": None,
        }
    candidates = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if class_name != "SunAwtCanvas" or not visible:
            continue
        if root_hwnd(hwnd) != fg_root:
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue
        window = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
            "root_hwnd": int(fg_root),
            "is_foreground_root": True,
        }
        candidates.append(window)
        return {
            "ok": True,
            "method": "foreground-canvas-window-light",
            "window": window,
            "candidate_count": len(candidates),
        }
    return {
        "ok": False,
        "method": "foreground-canvas-window-light",
        "reason": "foreground SunAwtCanvas not found",
        "foreground": foreground,
        "candidate_count": len(candidates),
        "window": None,
    }


def detect_entry_state_ready_light(jab, foreground=None, entry_button_target=None):
    foreground = foreground or foreground_info()
    direct = detect_entry_state_from_button_target(jab, entry_button_target, foreground)
    if direct.get("ok"):
        return direct
    fg_root = (foreground or {}).get("root")
    if os.name != "nt" or not hasattr(ctypes, "WinDLL") or not fg_root:
        return {
            "ok": False,
            "method": "light-entry-button-scan",
            "reason": "foreground root not available",
            "names": [],
            "hits": [],
        }
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if class_name != "SunAwtFrame" or not visible:
            continue
        if root_hwnd(hwnd) != fg_root:
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue
        window_info = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "visible": visible,
            "root_hwnd": int(fg_root),
            "is_foreground_root": True,
        }
        state = find_first_entry_state_button_in_window(jab, int(hwnd), window_info)
        if state.get("ok"):
            return state
    return {
        "ok": False,
        "method": "light-entry-button-scan",
        "reason": "保存/暂存/取消编辑态按钮未出现",
        "direct_button_state": direct,
        "names": [],
        "hits": [],
    }


def detect_entry_state_from_button_target(jab, target, foreground=None):
    target = target or {}
    window = target.get("window") or {}
    control = target.get("control") or {}
    hwnd = window.get("hwnd")
    path = control.get("path")
    if not hwnd or not path:
        return {
            "ok": False,
            "method": "direct-entry-button-path",
            "reason": "missing original new-button hwnd/path",
            "names": [],
            "hits": [],
        }
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        str(path),
        scope_hwnd=int(hwnd),
        require_showing=True,
        require_valid_bounds=True,
    )
    if not context:
        return {
            "ok": False,
            "method": "direct-entry-button-path",
            "reason": "original new-button path not found",
            "path": str(path),
            "window": window,
            "names": [],
            "hits": [],
        }
    try:
        info = jab.get_context_info(vm_id, context)
        current = summarize_info(jab, vm_id, context, info, str(path)) if info else {}
    finally:
        jab.release_contexts(vm_id, owned_contexts)
    names = sorted(normalize_entry_state_names(current))
    if not names or not is_current_visible_control(current):
        return {
            "ok": False,
            "method": "direct-entry-button-path",
            "reason": "original new-button path is not an edit-state button",
            "path": str(path),
            "window": window_info or window,
            "control": current,
            "names": names,
            "hits": [],
        }
    annotate_foreground_root_for_targets(
        [{"window": window_info or window}], foreground
    )
    return {
        "ok": True,
        "partial_ok": True,
        "method": "direct-entry-button-path",
        "names": names,
        "hits": [
            {
                "window": {
                    key: (window_info or window).get(key)
                    for key in (
                        "hwnd",
                        "title",
                        "class_name",
                        "class",
                        "visible",
                        "root_hwnd",
                        "is_foreground_root",
                    )
                },
                "control": current,
            }
        ],
    }


def find_first_entry_state_button_in_window(jab, hwnd, window_info):
    vm_id = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        int(hwnd),
        ctypes.byref(vm_id),
        ctypes.byref(root_context),
    ):
        return {
            "ok": False,
            "method": "light-entry-button-scan",
            "reason": "getAccessibleContextFromHWND failed",
            "names": [],
            "hits": [],
        }
    try:
        hit = find_first_entry_state_button_context(
            jab,
            vm_id.value,
            root_context.value,
            path="0",
            depth=0,
            max_depth=25,
            max_children=1000,
        )
    finally:
        jab.release_contexts(vm_id.value, [root_context.value])
    if not hit:
        return {
            "ok": False,
            "method": "light-entry-button-scan",
            "reason": "保存/暂存/取消编辑态按钮未出现",
            "names": [],
            "hits": [],
        }
    return {
        "ok": True,
        "partial_ok": True,
        "method": "light-entry-button-scan",
        "names": hit["names"],
        "hits": [
            {
                "window": window_info,
                "control": hit["control"],
            }
        ],
    }


def find_first_entry_state_button_context(
    jab,
    vm_id,
    context,
    path,
    depth,
    max_depth,
    max_children,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    control = summarize_info(jab, vm_id, context, info, path)
    names = sorted(normalize_entry_state_names(control))
    if names and is_current_visible_control(control):
        return {"names": names, "control": control}
    if depth >= max_depth:
        return None
    for index in range(min(info.childrenCount, max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        try:
            hit = find_first_entry_state_button_context(
                jab,
                vm_id,
                child,
                f"{path}.{index}",
                depth + 1,
                max_depth,
                max_children,
            )
            if hit:
                return hit
        finally:
            jab.release_contexts(vm_id, [child])
    return None


def describe_popup_visibility(hwnd):
    if os.name != "nt" or not hasattr(ctypes, "windll") or not hwnd:
        return {"ok": False, "reason": "missing hwnd", "hwnd": hwnd}
    return describe_hwnd(ctypes.windll.user32, wintypes.HWND(int(hwnd)))


def collect_new_visible_popup_windows(jab, before, max_depth=8, max_children=120):
    if os.name != "nt" or not hasattr(ctypes, "WinDLL"):
        return collect_receipt_new_windows(jab)

    before_signatures = {window_key(item): window_signature(item) for item in before}
    windows = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if class_name != "SunAwtWindow" or not visible:
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
            "is_java": True,
            "root": summarize_context(jab, vm_id.value, root_context.value, "0"),
            "controls": [],
            "all_controls": [],
        }
        key = window_key(window)
        if not is_visible_sun_awt_popup(window):
            windows.append(window)
            continue
        collect_controls(
            jab,
            vm_id.value,
            root_context.value,
            path="0",
            controls=window["controls"],
            all_controls=window["all_controls"],
            depth=0,
            max_depth=max_depth,
            max_children=max_children,
        )
        if key in before_signatures and before_signatures[key] == window_signature(
            window
        ):
            continue
        windows.append(window)
    return windows


def collect_receipt_new_baseline(jab):
    """Collect enough state for New-button lookup and later popup diffing."""
    foreground = foreground_info()
    windows = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if not class_name.startswith(("SunAwt", "Yonyou")):
            continue
        is_java = bool(jab.dll.isJavaWindow(hwnd))
        root = root_hwnd(hwnd)
        window = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
            "is_java": is_java,
            "root_hwnd": root,
            "is_foreground_root": bool(foreground and root == foreground.get("root")),
            "root": None,
            "controls": [],
            "all_controls": [],
        }
        windows.append(window)
        if not is_java or not visible:
            continue
        if class_name == "SunAwtFrame" and window["is_foreground_root"]:
            collect_window_controls_limited(
                jab, window, max_depth=25, max_children=1000
            )
        elif class_name == "SunAwtWindow":
            collect_window_controls_limited(jab, window, max_depth=8, max_children=120)
    return windows


def collect_window_controls_limited(jab, window, max_depth=8, max_children=120):
    hwnd = window.get("hwnd")
    if not hwnd:
        return window
    vm_id = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        int(hwnd),
        ctypes.byref(vm_id),
        ctypes.byref(root_context),
    ):
        return window
    window["root"] = summarize_context(jab, vm_id.value, root_context.value, "0")
    collect_controls(
        jab,
        vm_id.value,
        root_context.value,
        path="0",
        controls=window["controls"],
        all_controls=window["all_controls"],
        depth=0,
        max_depth=max_depth,
        max_children=max_children,
    )
    return window


def find_named_controls_in_windows(
    windows,
    name_query="新增",
    role=None,
    class_name=None,
    require_action=True,
):
    results = []
    name_query = str(name_query or "").lower()
    role = role.lower() if role else None
    for window in windows or []:
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


def new_button_priority(item):
    control = item.get("control") or {}
    states = control.get("states", "")
    bounds = control.get("bounds") or []
    is_foreground_root = bool((item.get("window") or {}).get("is_foreground_root"))
    has_valid_size = has_valid_bounds(bounds)
    is_showing = "showing" in states.lower()
    desc = control.get("description") or ""
    is_plain_new = desc == "新增(Ctrl+N)"
    return (
        not is_foreground_root,
        not is_showing,
        not has_valid_size,
        not is_plain_new,
    )


def filter_usable_new_buttons(buttons, foreground=None):
    buttons = [
        item
        for item in buttons or []
        if is_current_visible_control(item.get("control") or {})
    ]
    if not buttons:
        return []
    if foreground and foreground.get("root"):
        foreground_buttons = [
            item
            for item in buttons
            if (item.get("window") or {}).get("is_foreground_root")
        ]
        if foreground_buttons:
            return foreground_buttons
    return buttons


def annotate_foreground_root_for_targets(targets, foreground):
    for item in targets or []:
        window = item.get("window") or {}
        window["root_hwnd"] = root_hwnd(window.get("hwnd"))
        window["is_foreground_root"] = bool(
            foreground
            and foreground.get("root")
            and window.get("root_hwnd") == foreground.get("root")
        )


def annotate_foreground_root(windows, foreground):
    for window in windows or []:
        window["root_hwnd"] = root_hwnd(window.get("hwnd"))
        window["is_foreground_root"] = bool(
            foreground
            and foreground.get("root")
            and window.get("root_hwnd") == foreground.get("root")
        )


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


def root_hwnd(hwnd):
    if os.name != "nt" or not hasattr(ctypes, "windll") or not hwnd:
        return 0
    root = ctypes.windll.user32.GetAncestor(wintypes.HWND(int(hwnd)), 2)
    return int(root or 0)


def window_text(hwnd):
    if os.name != "nt" or not hasattr(ctypes, "windll") or not hwnd:
        return ""
    user32 = ctypes.windll.user32
    hwnd_obj = wintypes.HWND(int(hwnd))
    length = user32.GetWindowTextLengthW(hwnd_obj)
    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd_obj, buffer, length + 1)
    return buffer.value


def window_class_name(hwnd):
    if os.name != "nt" or not hasattr(ctypes, "windll") or not hwnd:
        return ""
    buffer = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(wintypes.HWND(int(hwnd)), buffer, 256)
    return buffer.value


def trigger_button_async(
    jab, hwnd, path, action_name=None, return_timeout=0.2, target=None
):
    result = jab.find_context_by_path_once(
        path,
        scope_hwnd=hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    context, vm_id, owned, _window_info = result
    if not context:
        return {
            "ok": False,
            "method": "button",
            "reason": "button path not found",
            "target": target,
        }

    status = {"returned": False, "ok": None, "exception": None}

    def run_action():
        try:
            status["ok"] = jab.do_action(
                vm_id,
                context,
                action_name=action_name,
                cleanup_blank_awt=False,
            )
        except Exception as exc:
            status["exception"] = repr(exc)
        finally:
            status["returned"] = True

    thread = threading.Thread(target=run_action, daemon=True)
    thread.start()
    thread.join(return_timeout)
    returned = not thread.is_alive()
    if returned:
        jab.release_contexts(vm_id, owned)
    return {
        "ok": True if not returned else bool(status["ok"]),
        "method": "button",
        "path": path,
        "target": target,
        "action_returned_within_timeout": returned,
        "action_status": status,
    }


def collect_receipt_new_windows(jab, max_depth=25, max_children=1000):
    windows = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if not class_name.startswith(("SunAwt", "Yonyou")):
            continue
        is_java = bool(jab.dll.isJavaWindow(hwnd))
        window = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
            "is_java": is_java,
            "root": None,
            "controls": [],
            "all_controls": [],
        }
        windows.append(window)
        if not is_java:
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            continue
        window["root"] = summarize_context(jab, vm_id.value, root_context.value, "0")
        collect_controls(
            jab,
            vm_id.value,
            root_context.value,
            path="0",
            controls=window["controls"],
            all_controls=window["all_controls"],
            depth=0,
            max_depth=max_depth,
            max_children=max_children,
        )
    return windows


def collect_controls(
    jab,
    vm_id,
    context,
    path,
    controls,
    all_controls,
    depth,
    max_depth,
    max_children,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    item = summarize_info(jab, vm_id, context, info, path)
    all_controls.append(item)
    if keep_control(item):
        controls.append(item)
    if depth >= max_depth:
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
            all_controls,
            depth + 1,
            max_depth,
            max_children,
        )
        jab.release_contexts(vm_id, [child])


def keep_control(item):
    role = item["role"].lower()
    if role in {
        "menu item",
        "menu",
        "push button",
        "text",
        "page tab",
        "page tab list",
    }:
        return True
    texts = {item["name"], item["description"]}
    if texts & SELF_MADE_NAMES or normalize_entry_state_names(item) & ENTRY_STATE_NAMES:
        return True
    if item["accessibleAction"]:
        return True
    return False


def summarize_context(jab, vm_id, context, path):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    return summarize_info(jab, vm_id, context, info, path)


def summarize_info(jab, vm_id, context, info, path):
    role = info.role_en_US.strip() or info.role.strip()
    item = {
        "path": path,
        "role": role,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "bounds": [info.x, info.y, info.width, info.height],
        "children": info.childrenCount,
        "accessibleAction": bool(info.accessibleAction),
        "actions": [],
    }
    if info.accessibleAction:
        item["actions"] = get_action_names(jab, vm_id, context)
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


def choose_self_made_menu_item(jab, windows, popup_hwnd=None):
    direct = choose_self_made_menu_item_direct(jab, popup_hwnd)
    if direct.get("ok"):
        return direct

    candidates = []
    for window in windows:
        if not window.get("is_java"):
            continue
        if not window.get("visible"):
            continue
        if popup_hwnd is not None and window.get("hwnd") != popup_hwnd:
            continue
        for control in window.get("all_controls", []):
            if not is_current_visible_control(control):
                continue
            if not control.get("accessibleAction"):
                continue
            if (
                control["role"].lower() == "menu item"
                or control.get("name") in SELF_MADE_NAMES
            ):
                candidates.append({"window": window, "control": control})

    named = [
        item for item in candidates if item["control"].get("name") in SELF_MADE_NAMES
    ]
    if named:
        target = named[0]
    else:
        return {
            "ok": False,
            "reason": "未找到可见命名为【自制】的菜单项；正式收款流程不按序号兜底",
            "candidate_count": len(candidates),
            "candidates": summarize_candidates(candidates),
        }

    ok = do_action_by_window_path(
        jab,
        target["window"]["hwnd"],
        target["control"]["path"],
        action_name=choose_click_action(target["control"].get("actions", [])),
    )
    return {
        "ok": bool(ok),
        "method": "menu-scan",
        "target": {
            "window": {
                key: target["window"].get(key)
                for key in ("hwnd", "title", "class_name", "visible")
            },
            "control": target["control"],
        },
        "candidate_count": len(candidates),
        "candidates": summarize_candidates(candidates),
    }


def choose_self_made_menu_item_direct(jab, popup_hwnd=None):
    if not popup_hwnd:
        return {"ok": False, "reason": "missing popup hwnd"}
    popup_state = describe_popup_visibility(popup_hwnd)
    if not popup_state.get("exists") or not popup_state.get("visible"):
        return {"ok": False, "reason": "popup not visible", "popup": popup_state}
    path = "0.0.1.0.0.0"
    result = jab.find_context_by_path_once(
        path,
        scope_hwnd=popup_hwnd,
        require_showing=True,
        require_valid_bounds=True,
    )
    context, vm_id, owned, _window_info = result
    if not context:
        return {
            "ok": False,
            "method": "tracked-popup-direct",
            "reason": "tracked popup self-made path not found",
            "path": path,
            "popup": popup_state,
        }
    try:
        info = jab.get_context_info(vm_id, context)
        control = summarize_info(jab, vm_id, context, info, path) if info else {}
        if control.get("name") not in SELF_MADE_NAMES:
            return {
                "ok": False,
                "method": "tracked-popup-direct",
                "reason": "tracked popup direct path is not 自制",
                "path": path,
                "control": control,
                "popup": popup_state,
            }
        ok = jab.do_action(
            vm_id,
            context,
            action_name=choose_click_action(control.get("actions", [])),
            cleanup_blank_awt=False,
        )
    finally:
        jab.release_contexts(vm_id, owned)
    return {
        "ok": bool(ok),
        "method": "tracked-popup-direct",
        "path": path,
        "target": {
            "window": {
                "hwnd": int(popup_hwnd),
                "class_name": popup_state.get("class_name"),
                "title": popup_state.get("title"),
                "visible": popup_state.get("visible"),
            },
            "control": control,
        },
        "candidate_count": 1,
        "candidates": [
            {
                "window": {
                    "hwnd": int(popup_hwnd),
                    "class_name": popup_state.get("class_name"),
                    "title": popup_state.get("title"),
                    "visible": popup_state.get("visible"),
                },
                "control": control,
            }
        ],
    }


def find_new_visible_popup(before, after):
    before_signatures = {window_key(item): window_signature(item) for item in before}
    candidates = []
    for item in after:
        key = window_key(item)
        if key in before_signatures and before_signatures[key] == window_signature(
            item
        ):
            continue
        if not is_visible_sun_awt_popup(item):
            continue
        menu_names = {
            control.get("name")
            for control in item.get("all_controls", [])
            if control.get("role", "").lower() == "menu item"
        }
        if SELF_MADE_NAMES & menu_names or "应收单" in menu_names:
            candidates.append(item)
    candidates.sort(key=lambda item: (item.get("root") or {}).get("bounds", [0, 0])[1])
    if not candidates:
        return None
    item = candidates[0]
    return {
        "hwnd": item.get("hwnd"),
        "class_name": item.get("class_name"),
        "title": item.get("title"),
        "visible": item.get("visible"),
        "root": item.get("root"),
        "menu_items": [
            summarize_control(control)
            for control in item.get("all_controls", [])
            if control.get("role", "").lower() == "menu item"
        ],
    }


def is_visible_sun_awt_popup(window):
    if window.get("class_name") != "SunAwtWindow" or not window.get("visible"):
        return False
    root = window.get("root") or {}
    bounds = root.get("bounds") or []
    if len(bounds) != 4:
        return False
    _x, _y, width, height = bounds
    if width <= 0 or height <= 0:
        return False
    return width <= 500 and height <= 500


def close_popup_hwnd(hwnd):
    if os.name != "nt":
        return {"ok": False, "reason": "Windows only", "hwnd": hwnd}
    if not hwnd:
        return {"ok": False, "reason": "missing hwnd"}
    user32 = ctypes.windll.user32
    hwnd_obj = wintypes.HWND(int(hwnd))
    before = describe_hwnd(user32, hwnd_obj)
    if not before.get("exists"):
        return {"ok": True, "reason": "already gone", "before": before}
    if before.get("class_name") != "SunAwtWindow":
        return {"ok": False, "reason": "class mismatch", "before": before}
    user32.ShowWindow(hwnd_obj, 0)
    user32.SetWindowPos(
        hwnd_obj, 0, -32000, -32000, 0, 0, 0x0001 | 0x0010 | 0x0080 | 0x0200
    )
    user32.PostMessageW(hwnd_obj, 0x0010, 0, 0)
    return {"ok": True, "before": before, "after": describe_hwnd(user32, hwnd_obj)}


def describe_hwnd(user32, hwnd):
    if not user32.IsWindow(hwnd):
        return {"exists": False}

    class Rect(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    title_len = user32.GetWindowTextLengthW(hwnd)
    title = ctypes.create_unicode_buffer(title_len + 1)
    user32.GetWindowTextW(hwnd, title, title_len + 1)
    class_name = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_name, 256)
    rect = Rect()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return {
        "exists": True,
        "hwnd": int(hwnd.value),
        "visible": bool(user32.IsWindowVisible(hwnd)),
        "class_name": class_name.value,
        "title": title.value,
        "rect": [rect.left, rect.top, rect.right, rect.bottom],
        "width": rect.right - rect.left,
        "height": rect.bottom - rect.top,
    }


def has_valid_bounds(bounds):
    return (
        isinstance(bounds, list)
        and len(bounds) == 4
        and bounds[0] >= 0
        and bounds[1] >= 0
        and bounds[2] > 0
        and bounds[3] > 0
    )


def is_current_visible_control(control):
    states = control.get("states", "")
    bounds = control.get("bounds") or []
    normalized_states = states.lower()
    if "visible" not in normalized_states or "showing" not in normalized_states:
        return False
    return (
        isinstance(bounds, list)
        and len(bounds) == 4
        and bounds[2] > 0
        and bounds[3] > 0
    )


def choose_click_action(actions):
    if not actions:
        return None
    for preferred in ("单击", "click", "press"):
        if preferred in actions:
            return preferred
    return actions[0]


def do_action_by_window_path(jab, hwnd, path, action_name=None):
    result = jab.find_context_by_path_once(
        path,
        scope_hwnd=hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    context, vm_id, owned, _window_info = result
    if not context:
        return False
    try:
        return jab.do_action(
            vm_id,
            context,
            action_name=action_name,
            cleanup_blank_awt=False,
        )
    finally:
        jab.release_contexts(vm_id, owned)


def summarize_candidates(candidates):
    result = []
    for index, item in enumerate(candidates):
        result.append(
            {
                "index": index,
                "window": {
                    key: item["window"].get(key)
                    for key in (
                        "hwnd",
                        "class_name",
                        "title",
                        "visible",
                        "root_hwnd",
                        "is_foreground_root",
                    )
                },
                "control": item["control"],
            }
        )
    return result


def summarize_report(report):
    return {
        "foreground": report.get("foreground"),
        "matches": [summarize_target(item) for item in report.get("matches", [])[:20]],
        "buttons": [summarize_target(item) for item in report.get("buttons", [])[:20]],
        "usable_buttons": [
            summarize_target(item) for item in report.get("usable_buttons", [])[:20]
        ],
        "open": summarize_action_report(report.get("open")),
        "tracked_popup": report.get("tracked_popup"),
        "popup_cleanup": report.get("popup_cleanup"),
        "changed_windows": [
            {
                "hwnd": item.get("hwnd"),
                "class_name": item.get("class_name"),
                "title": item.get("title"),
                "visible": item.get("visible"),
                "root": item.get("root"),
                "controls": [
                    summarize_control(control)
                    for control in item.get("controls", [])[:30]
                ],
            }
            for item in report.get("new_or_changed_after_open", [])[:10]
        ],
        "choose_self_made": summarize_action_report(report.get("choose_self_made")),
        "entry_context_snapshot": report.get("entry_context_snapshot"),
        "entry_state": report.get("entry_state"),
        "timings": report.get("timings") or [],
    }


def summarize_action_report(action_report):
    if not isinstance(action_report, dict):
        return action_report
    result = {
        key: value
        for key, value in action_report.items()
        if key
        in {
            "ok",
            "method",
            "reason",
            "path",
            "candidate_count",
            "action_returned_within_timeout",
            "action_status",
            "rejected_count",
        }
    }
    if "target" in action_report:
        result["target"] = summarize_target(action_report["target"])
    if "candidates" in action_report:
        result["candidates"] = [
            summarize_target(item) for item in action_report.get("candidates", [])[:20]
        ]
    return result


def summarize_target(item):
    if not isinstance(item, dict):
        return item
    return {
        "window": {
            key: item.get("window", {}).get(key)
            for key in (
                "hwnd",
                "class_name",
                "title",
                "visible",
                "root_hwnd",
                "is_foreground_root",
            )
        },
        "control": summarize_control(item.get("control", {})),
    }


def summarize_control(control):
    return {
        key: control.get(key)
        for key in (
            "path",
            "role",
            "name",
            "description",
            "states",
            "bounds",
            "accessibleAction",
            "actions",
        )
    }


def detect_self_made_entry_state(windows):
    names = set()
    hits = []
    for window in windows or []:
        for control in window.get("controls", []):
            matched_names = normalize_entry_state_names(control)
            if matched_names:
                names.update(matched_names)
                hits.append(
                    {
                        "window": {
                            key: window.get(key)
                            for key in ("hwnd", "class_name", "title", "visible")
                        },
                        "control": control,
                    }
                )
    return {
        "ok": bool(names),
        "partial_ok": bool(names),
        "names": sorted(names),
        "hits": hits,
    }


def normalize_entry_state_names(control):
    texts = {
        str(control.get("name") or "").strip(),
        str(control.get("description") or "").strip(),
    }
    matched = set()
    if "暂存" in texts:
        matched.add("暂存")
    if "保存(Ctrl+S)" in texts or "保存" in texts:
        matched.add("保存(Ctrl+S)")
    if "取消(Ctrl+Q)" in texts or "取消" in texts:
        matched.add("取消(Ctrl+Q)")
    return matched


def diff_windows(before, after):
    before_signatures = {window_key(item): window_signature(item) for item in before}
    changed = []
    for item in after:
        key = window_key(item)
        sig = window_signature(item)
        if key not in before_signatures or before_signatures[key] != sig:
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
            )
        )
    return (
        json.dumps(window.get("root"), ensure_ascii=False, sort_keys=True),
        tuple(controls),
    )


def print_text(report):
    print("open:", json.dumps(report["open"], ensure_ascii=False))
    print("new_or_changed_after_open:", len(report["new_or_changed_after_open"]))
    for window in report["new_or_changed_after_open"]:
        print(
            f"  window hwnd={window['hwnd']} class={window['class_name']!r} "
            f"title={window['title']!r} visible={window['visible']} root={window['root']}"
        )
        for control in window["controls"][:80]:
            print(
                f"    path={control['path']} role={control['role']!r} "
                f"name={control['name']!r} desc={control['description']!r} "
                f"states={control['states']!r} actions={control['actions']} "
                f"bounds={control['bounds']}"
            )
    print(
        "choose_self_made:", json.dumps(report["choose_self_made"], ensure_ascii=False)
    )
    print("entry_state:", json.dumps(report["entry_state"], ensure_ascii=False))


if __name__ == "__main__":
    raise SystemExit(main())
