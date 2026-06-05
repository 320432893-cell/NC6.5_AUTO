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


def run(args):
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        before = collect_receipt_new_windows(jab)
        matches = find_named_controls(
            jab, args.name, args.role, args.class_name, require_action=False
        )
        buttons = find_named_controls(
            jab, args.name, args.role, args.class_name, require_action=True
        )
        open_report = open_new_menu(jab, args)
        time.sleep(args.wait)
        after_open = collect_receipt_new_windows(jab)
        tracked_popup = find_new_visible_popup(before, after_open)
        choose_report = None
        after_choose = None
        popup_cleanup = None
        residue_cleanup = None
        if args.choose_self_made and open_report.get("ok"):
            choose_report = choose_self_made_menu_item(
                jab,
                after_open,
                args.self_made_index,
                popup_hwnd=tracked_popup.get("hwnd") if tracked_popup else None,
            )
            time.sleep(args.wait)
            if tracked_popup and choose_report.get("ok"):
                popup_cleanup = close_popup_hwnd(tracked_popup["hwnd"])
                time.sleep(0.1)
            if choose_report.get("ok"):
                residue_cleanup = cleanup_awt_popup_residue()
                time.sleep(0.1)
            after_choose = collect_receipt_new_windows(jab)
        elif args.choose_self_made:
            choose_report = {
                "ok": False,
                "reason": "new menu was not opened; self-made selection skipped",
            }
        entry_state = detect_self_made_entry_state(after_choose or after_open)
    finally:
        jab.hide_blank_awt_windows_enabled = False
        jab.close()

    report = {
        "matches": matches,
        "buttons": buttons,
        "open": open_report,
        "tracked_popup": tracked_popup,
        "popup_cleanup": popup_cleanup,
        "residue_cleanup": residue_cleanup,
        "new_or_changed_after_open": diff_windows(before, after_open),
        "windows_after_open": after_open,
        "choose_self_made": choose_report,
        "windows_after_choose": after_choose,
        "entry_state": entry_state,
    }
    return report


def open_new_menu(jab, args):
    if args.method == "probe-button":
        return {"ok": True, "method": "probe-button"}
    if args.method == "button":
        buttons = find_new_buttons(jab, args.name, args.role, args.class_name)
        if not buttons:
            return {"ok": False, "method": "button", "reason": "new button not found"}
        target = buttons[0]
        return trigger_button_async(
            jab,
            target["window"]["hwnd"],
            target["control"]["path"],
            action_name=args.action,
            return_timeout=args.return_timeout,
            target=target,
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


def find_new_buttons(jab, name_query="新增", role=None, class_name=None):
    buttons = find_named_controls(
        jab,
        name_query=name_query,
        role=role,
        class_name=class_name,
        require_action=True,
    )
    buttons.sort(key=new_button_priority)
    return buttons


def new_button_priority(item):
    control = item.get("control") or {}
    states = control.get("states", "")
    bounds = control.get("bounds") or []
    has_valid_size = len(bounds) == 4 and bounds[2] > 0 and bounds[3] > 0
    is_showing = "showing" in states
    desc = control.get("description") or ""
    is_plain_new = desc == "新增(Ctrl+N)"
    return (not is_showing, not has_valid_size, not is_plain_new)


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
    before_keys = {window_key(item) for item in before}
    candidates = []
    for item in after:
        if window_key(item) in before_keys:
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


def is_current_visible_control(control):
    states = control.get("states", "")
    bounds = control.get("bounds") or []
    if "visible" not in states or "showing" not in states:
        return False
    if len(bounds) != 4:
        return False
    _x, _y, width, height = bounds
    return width > 0 and height > 0


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
                    for key in ("hwnd", "class_name", "title", "visible")
                },
                "control": item["control"],
            }
        )
    return result


def summarize_report(report):
    return {
        "matches": [summarize_target(item) for item in report.get("matches", [])[:20]],
        "buttons": [summarize_target(item) for item in report.get("buttons", [])[:20]],
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
            for key in ("hwnd", "class_name", "title", "visible")
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
