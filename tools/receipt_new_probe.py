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
from tools.jab_probe import AccessibleActions, JOBJECT, enum_windows  # noqa: E402


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
            and report.get("entry_state", {}).get("ok")
            else 1
        )
    return 0 if report.get("open", {}).get("ok") else 1


def run(args, jab=None, before=None, buttons=None):
    timings = []
    owns_jab = jab is None
    if owns_jab:
        cfg = load_config(args.config)
        jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        if owns_jab:
            measure(timings, "new-probe.jab.ensure-started", jab.ensure_started)
        else:
            timings.append({"name": "new-probe.jab.ensure-started", "seconds": 0.0})
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
        residue_cleanup = None
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
            pre_choose_entry = None
            if not tracked_popup:
                pre_choose_entry = measure(
                    timings,
                    "new-probe.pre-choose-entry-state",
                    wait_for_self_made_entry_state,
                    jab,
                    0.35,
                    0.04,
                )
            if pre_choose_entry and (
                pre_choose_entry.get("ok") or pre_choose_entry.get("partial_ok")
            ):
                choose_report = {
                    "ok": True,
                    "method": "entry-state-already-open",
                    "reason": "self-made entry state detected without visible popup",
                }
                entry_wait = pre_choose_entry
                after_choose = entry_wait.get("windows") or []
            else:
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
                        "new-probe.wait-for-entry-state",
                        quick_check_self_made_entry_state,
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
            if (
                choose_report.get("ok")
                and entry_wait
                and (entry_wait.get("ok") or entry_wait.get("partial_ok"))
            ):
                residue_cleanup = measure(
                    timings,
                    "new-probe.residue-cleanup",
                    cleanup_awt_popup_residue,
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
        "residue_cleanup": residue_cleanup,
        "new_or_changed_after_open": diff_windows(before, after_open),
        "windows_after_open": after_open,
        "choose_self_made": choose_report,
        "windows_after_choose": after_choose,
        "entry_state": entry_state,
        "timings": timings,
    }
    return report


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
        fallback = open_new_menu_with_ctrl_n(foreground)
        fallback["button_action"] = action_report
        return fallback
    if args.method == "button" and all_buttons:
        fallback = open_new_menu_with_ctrl_n(foreground)
        fallback.update(
            {
                "button_reason": "new button candidates were found, but none were usable in the foreground NC window",
                "rejected_count": len(all_buttons),
                "rejected": summarize_candidates(all_buttons[:20]),
            }
        )
        return fallback
    if args.method == "button":
        fallback = open_new_menu_with_ctrl_n(foreground)
        fallback.update(
            {
                "button_reason": "new button not found",
                "rejected_count": 0,
                "rejected": [],
            }
        )
        return fallback
    return open_new_menu(jab, args)


def open_new_menu_with_ctrl_n(foreground):
    guard = foreground_nc_guard(foreground)
    if not guard.get("ok"):
        return {
            "ok": False,
            "method": "ctrl+n-fallback",
            "reason": "foreground is not NC; Ctrl+N not sent",
            "foreground_guard": guard,
        }
    try:
        send_ctrl_n()
    except Exception as exc:
        return {
            "ok": False,
            "method": "ctrl+n-fallback",
            "reason": f"Ctrl+N send failed: {exc!r}",
            "foreground_guard": guard,
        }
    return {
        "ok": True,
        "method": "ctrl+n-fallback",
        "foreground_guard": guard,
    }


def foreground_nc_guard(foreground):
    if os.name != "nt":
        return {"ok": False, "reason": "Windows only", "foreground": foreground}
    if not foreground:
        return {
            "ok": False,
            "reason": "missing foreground window",
            "foreground": foreground,
        }
    ok = bool(
        foreground.get("class_name") == "YonyouUWnd"
        or foreground.get("root_class_name") == "YonyouUWnd"
    )
    return {
        "ok": ok,
        "reason": None if ok else "foreground root is not YonyouUWnd",
        "foreground": foreground,
    }


def send_ctrl_n():
    user32 = ctypes.windll.user32
    vk_control = 0x11
    vk_n = 0x4E
    try:
        user32.keybd_event(vk_control, 0, 0, 0)
        time.sleep(0.02)
        user32.keybd_event(vk_n, 0, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(vk_n, 0, 2, 0)
    finally:
        user32.keybd_event(vk_control, 0, 2, 0)


def open_new_menu(jab, args):
    if args.method == "probe-button":
        return {"ok": True, "method": "probe-button"}
    if args.method == "button":
        foreground = foreground_info()
        buttons = find_new_buttons(jab, args.name, args.role, args.class_name)
        if not buttons:
            fallback = open_new_menu_with_ctrl_n(foreground)
            fallback.update({"button_reason": "new button not found"})
            return fallback
        return open_new_menu_with_known_buttons(
            jab, args, buttons, all_buttons=buttons, foreground=foreground
        )

    ok = jab.do_action_by_path(
        args.path,
        title=args.title,
        class_name=args.class_name,
        name=args.name,
        role=args.role,
        action_name=args.action,
        wait=0,
        timeout=2.0,
        require_showing=False,
        require_valid_bounds=False,
        cleanup_blank_awt=False,
    )
    return {"ok": bool(ok), "method": "action-path", "path": args.path}


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


def wait_for_self_made_entry_state(jab, timeout=0.45, interval=0.04):
    start = time.perf_counter()
    deadline = time.perf_counter() + max(float(timeout or 0), 0)
    attempts = 0
    last_windows = []
    last_state = None
    while True:
        attempts += 1
        last_windows = collect_receipt_new_windows_compat(
            jab, max_depth=12, max_children=320
        )
        last_state = detect_self_made_entry_state(last_windows)
        if last_state.get("ok") or last_state.get("partial_ok"):
            return {
                "ok": True,
                "partial_ok": bool(last_state.get("partial_ok"))
                and not bool(last_state.get("ok")),
                "attempts": attempts,
                "wait_seconds": elapsed(start),
                "state": last_state,
                "windows": last_windows,
            }
        remaining = deadline - time.perf_counter()
        if remaining <= 0:
            break
        time.sleep(min(interval, remaining))
    full_windows = collect_receipt_new_windows_compat(
        jab, max_depth=18, max_children=520
    )
    full_state = detect_self_made_entry_state(full_windows)
    if full_state.get("ok") or full_state.get("partial_ok"):
        return {
            "ok": True,
            "partial_ok": bool(full_state.get("partial_ok"))
            and not bool(full_state.get("ok")),
            "attempts": attempts,
            "wait_seconds": elapsed(start),
            "state": full_state,
            "windows": full_windows,
        }
    return {
        "ok": False,
        "attempts": attempts,
        "wait_seconds": max(float(timeout or 0), 0),
        "state": last_state or {"ok": False, "names": [], "hits": []},
        "windows": last_windows,
    }


def quick_check_self_made_entry_state(jab):
    windows = collect_receipt_new_windows_compat(jab, max_depth=12, max_children=320)
    state = detect_self_made_entry_state(windows)
    confirmed = bool(state.get("ok") or state.get("partial_ok"))
    if not confirmed:
        windows = collect_receipt_new_windows_compat(
            jab, max_depth=18, max_children=520
        )
        state = detect_self_made_entry_state(windows)
        confirmed = bool(state.get("ok") or state.get("partial_ok"))
    if not confirmed:
        state = {
            **state,
            "partial_ok": True,
            "reason": "quick check did not see entry buttons; trusting successful self-made action and deferring to header fill",
        }
    return {
        "ok": True,
        "confirmed": confirmed,
        "state": state,
        "windows": windows,
    }


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


def collect_receipt_new_windows_compat(jab, **kwargs):
    try:
        return collect_receipt_new_windows(jab, **kwargs)
    except TypeError as exc:
        if "unexpected keyword argument" not in str(exc):
            raise
        return collect_receipt_new_windows(jab)


def find_new_buttons(jab, name_query="新增", role=None, class_name=None):
    buttons = find_named_controls(
        jab,
        name_query=name_query,
        role=role,
        class_name=class_name,
        require_action=True,
    )
    foreground = foreground_info()
    annotate_foreground_root_for_targets(buttons, foreground)
    buttons = filter_usable_new_buttons(buttons, foreground)
    buttons.sort(key=new_button_priority)
    return buttons


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


def choose_self_made_menu_item(jab, windows, fallback_index, popup_hwnd=None):
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
    elif fallback_index is not None and 0 <= fallback_index < len(candidates):
        target = candidates[fallback_index]
    else:
        return {
            "ok": False,
            "reason": "self-made menu item not found",
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


def cleanup_awt_popup_residue():
    if os.name != "nt":
        return {"ok": False, "reason": "Windows only", "targets": []}
    user32 = ctypes.windll.user32
    targets = []
    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _lparam):
        hwnd_obj = wintypes.HWND(int(hwnd))
        item = describe_hwnd(user32, hwnd_obj)
        if (
            item.get("exists")
            and item.get("class_name") == "SunAwtWindow"
            and item.get("title") == ""
            and 0 < item.get("width", 0) <= 250
            and 0 < item.get("height", 0) <= 250
        ):
            targets.append(item)
        return True

    user32.EnumWindows(enum_proc(callback), 0)
    for item in targets:
        hwnd_obj = wintypes.HWND(int(item["hwnd"]))
        user32.EnableWindow(hwnd_obj, True)
        user32.ShowWindow(hwnd_obj, 0)
        user32.SetWindowPos(
            hwnd_obj, 0, -32000, -32000, 0, 0, 0x0001 | 0x0010 | 0x0080 | 0x0200
        )
        user32.PostMessageW(hwnd_obj, 0x0010, 0, 0)
        item["after"] = describe_hwnd(user32, hwnd_obj)
    if targets:
        user32.RedrawWindow(
            user32.GetDesktopWindow(),
            None,
            0,
            0x0001 | 0x0004 | 0x0080 | 0x0100,
        )
    return {"ok": True, "targets": targets}


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
        "residue_cleanup": report.get("residue_cleanup"),
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
        "ok": ENTRY_STATE_NAMES.issubset(names),
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
