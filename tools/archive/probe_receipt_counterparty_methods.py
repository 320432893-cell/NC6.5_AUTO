import argparse
import ctypes
import json
import os
import sys
import time
from ctypes import wintypes
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.jab_popup import (  # noqa: E402
    close_popup_hwnd,
    collect_new_visible_popup_windows,
    collect_visible_popup_windows,
)
from core.utils import load_config  # noqa: E402
from tools.archive.probe_receipt_counterparty_popup_tree import (  # noqa: E402
    find_counterparty_combo_near_label,
    resolve_current_scope,
    strip_handles,
    summarize_windows,
)
from tools.jab_probe import JOBJECT  # noqa: E402
from tools.receipt_full_flow_entry import (  # noqa: E402
    find_counterparty_combo,
    read_counterparty_combo_state,
)

COUNTERPARTY_EXPECTED = "客户"


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Probe receipt header counterparty dropdown open/commit methods."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--scope-hwnd", type=int, default=None)
    parser.add_argument("--wait", type=float, default=0.45)
    parser.add_argument("--poll", type=float, default=0.08)
    parser.add_argument("--max-depth", type=int, default=8)
    parser.add_argument("--max-children", type=int, default=120)
    parser.add_argument(
        "--commit",
        action="store_true",
        help=(
            "After each open method, test selection lifecycle. This may change "
            "往来对象 to 客户."
        ),
    )
    parser.add_argument(
        "--methods",
        default=(
            "toggle,esc-toggle,double-toggle,focus-alt-down,focus-f4,"
            "focus-space,focus-enter,home-enter,esc-home-enter,"
            "activate-enter,activate-home-enter,activate-esc,"
            "embedded-select-enter,embedded-action-enter,"
            "embedded-select-activate-enter,snapshot"
        ),
        help=(
            "Comma-separated methods: toggle,esc-toggle,double-toggle,"
            "focus-alt-down,focus-f4,focus-space,focus-enter,home-enter,"
            "esc-home-enter,activate-enter,activate-home-enter,activate-esc,"
            "embedded-select-enter,embedded-action-enter,"
            "embedded-select-activate-enter,snapshot"
        ),
    )
    parser.add_argument("--json", action="store_true")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    report = {
        "ok": False,
        "commit": bool(args.commit),
        "methods_requested": parse_methods(args.methods),
        "scope": None,
        "path_found": None,
        "near_label_found": None,
        "target": None,
        "method_results": [],
    }
    try:
        jab.ensure_started()
        cleanup_all_visible_popups(jab)
        scope = resolve_current_scope(jab, args.scope_hwnd)
        report["scope"] = scope
        if not scope.get("ok"):
            report["reason"] = scope.get("reason") or "receipt header scope not found"
            return finish(report, args)

        dynamic_index = scope.get("dynamic_index")
        scope_hwnd = scope.get("hwnd")
        near_label_found = None
        found = find_counterparty_combo(jab, dynamic_index, scope_hwnd=scope_hwnd)
        if not found.get("ok"):
            near_label_found = find_counterparty_combo_near_label(
                jab, scope_hwnd=scope_hwnd
            )
            found = near_label_found
        report["path_found"] = None
        report["near_label_found"] = strip_handles(near_label_found)
        if not found.get("ok"):
            report["reason"] = found.get("reason") or "counterparty combo not found"
            return finish(report, args)

        report["target"] = {
            "method": found.get("method") or found.get("source"),
            "path": found.get("path"),
            "window": found.get("window"),
            "root_hwnd": root_hwnd((found.get("window") or {}).get("hwnd")),
            "target": found.get("target"),
        }
        report["initial_combo_snapshot"] = snapshot_combo_tree(
            jab, found["vm_id"], found["context"], depth=6
        )
        try:
            for method in parse_methods(args.methods):
                result = probe_method(
                    jab,
                    found,
                    method,
                    wait=args.wait,
                    poll=args.poll,
                    max_depth=args.max_depth,
                    max_children=args.max_children,
                    commit=args.commit,
                )
                report["method_results"].append(result)
                cleanup_all_visible_popups(jab)
                time.sleep(0.12)
            report["ok"] = any(item.get("opens_counterparty_popup") for item in report["method_results"])
            if not report["ok"]:
                report["reason"] = "no method opened a detectable counterparty popup"
        finally:
            jab.release_contexts(found["vm_id"], found["owned_contexts"])
    finally:
        jab.close()
    return finish(report, args)


def parse_methods(text):
    return [item.strip() for item in str(text or "").split(",") if item.strip()]


def probe_method(
    jab,
    found,
    method,
    wait,
    poll,
    max_depth,
    max_children,
    commit=False,
):
    vm_id = found["vm_id"]
    context = found["context"]
    window_hwnd = ((found.get("window") or {}).get("hwnd"))
    window_root = root_hwnd(window_hwnd) or window_hwnd
    before_state = read_counterparty_combo_state(jab, vm_id, context)
    before_tree = snapshot_combo_tree(jab, vm_id, context, depth=6)
    before_embedded = probe_embedded_counterparty_tree(jab, vm_id, context)
    before_popups = collect_visible_popup_windows(jab)
    action = run_open_method(jab, vm_id, context, method, window_root=window_root)
    popup_probe = wait_for_any_counterparty_popup(
        jab, before_popups, timeout=wait, interval=poll
    )
    visible_after_open = collect_visible_popup_windows(jab)
    after_open_state = read_counterparty_combo_state(jab, vm_id, context)
    after_open_tree = snapshot_combo_tree(jab, vm_id, context, depth=6)
    after_open_embedded = probe_embedded_counterparty_tree(jab, vm_id, context)
    commit_result = None
    if commit:
        commit_result = commit_method(jab, vm_id, context, method, window_root=window_root)
        time.sleep(0.12)
    visible_after_commit = collect_visible_popup_windows(jab)
    after_commit_state = read_counterparty_combo_state(jab, vm_id, context)
    after_commit_tree = snapshot_combo_tree(jab, vm_id, context, depth=6)
    after_commit_embedded = probe_embedded_counterparty_tree(jab, vm_id, context)
    cleanup = cleanup_popups(visible_after_open + visible_after_commit)
    return {
        "method": method,
        "before": before_state,
        "before_tree": before_tree,
        "before_embedded": before_embedded,
        "action": action,
        "popup_probe": popup_probe,
        "embedded_popup_probe": find_counterparty_in_tree(after_open_tree),
        "after_open_embedded": after_open_embedded,
        "opens_counterparty_popup": bool(popup_probe.get("counterparty_popup")),
        "after_open": after_open_state,
        "after_open_tree": after_open_tree,
        "commit": commit_result,
        "after_commit": after_commit_state,
        "after_commit_tree": after_commit_tree,
        "after_commit_embedded": after_commit_embedded,
        "visible_after_open": summarize_windows(
            visible_after_open, max_depth, max_children
        ),
        "visible_after_commit": summarize_windows(
            visible_after_commit, max_depth, max_children
        ),
        "cleanup": cleanup,
    }


def run_open_method(jab, vm_id, context, method, window_root=None):
    if method == "snapshot":
        return {
            "ok": True,
            "method": method,
            "reason": "read-only snapshot; no focus, no action, no key",
        }
    focus_ok = request_focus(jab, vm_id, context)
    try:
        if method == "esc-toggle":
            jab.press_key("esc", wait=0.08)
            ok = bool(jab.do_action(vm_id, context, action_name="togglePopup"))
            return {
                "ok": ok,
                "method": method,
                "request_focus": focus_ok,
            }
        if method == "double-toggle":
            first = bool(jab.do_action(vm_id, context, action_name="togglePopup"))
            time.sleep(0.12)
            second = bool(jab.do_action(vm_id, context, action_name="togglePopup"))
            return {
                "ok": bool(first and second),
                "method": method,
                "first": first,
                "second": second,
                "request_focus": focus_ok,
            }
        if method == "toggle":
            return {
                "ok": bool(jab.do_action(vm_id, context, action_name="togglePopup")),
                "method": method,
                "request_focus": focus_ok,
            }
        if method == "focus-alt-down":
            jab.press_hotkey("alt", "down", wait=0)
            return {"ok": True, "method": method, "request_focus": focus_ok}
        if method == "focus-f4":
            jab.press_key("f4", wait=0)
            return {"ok": True, "method": method, "request_focus": focus_ok}
        if method == "focus-space":
            jab.press_key("space", wait=0)
            return {"ok": True, "method": method, "request_focus": focus_ok}
        if method == "focus-enter":
            jab.press_key("enter", wait=0)
            return {"ok": True, "method": method, "request_focus": focus_ok}
        if method == "home-enter":
            return {**send_home_enter(jab), "request_focus": focus_ok}
        if method == "esc-home-enter":
            jab.press_key("esc", wait=0.08)
            return {**send_home_enter(jab), "request_focus": focus_ok}
        if method == "activate-enter":
            return {
                **activate_and_press(jab, window_root, ["enter"]),
                "method": method,
                "request_focus": focus_ok,
            }
        if method == "activate-home-enter":
            return {
                **activate_and_press(
                    jab, window_root, ["home", "enter"]
                ),
                "method": method,
                "request_focus": focus_ok,
            }
        if method == "activate-esc":
            return {
                **activate_and_press(jab, window_root, ["esc"]),
                "method": method,
                "request_focus": focus_ok,
            }
        if method == "embedded-select-enter":
            return {
                **select_embedded_customer_option(jab, vm_id, context, press_enter=False),
                "method": method,
                "request_focus": focus_ok,
            }
        if method == "embedded-action-enter":
            return {
                **action_embedded_customer_option(jab, vm_id, context, press_enter=False),
                "method": method,
                "request_focus": focus_ok,
            }
        if method == "embedded-select-activate-enter":
            selected = select_embedded_customer_option(jab, vm_id, context, press_enter=False)
            pressed = activate_and_press(jab, window_root, ["enter"])
            return {
                "ok": bool(selected.get("ok") and pressed.get("ok")),
                "method": method,
                "request_focus": focus_ok,
                "select": selected,
                "activate_press": pressed,
            }
        return {"ok": False, "method": method, "reason": "unknown method"}
    except Exception as exc:
        return {
            "ok": False,
            "method": method,
            "request_focus": focus_ok,
            "error": repr(exc),
        }


def request_focus(jab, vm_id, context):
    if not hasattr(jab.dll, "requestFocus"):
        return {"ok": None, "reason": "requestFocus unavailable"}
    try:
        return {"ok": bool(jab.dll.requestFocus(vm_id, context))}
    except Exception as exc:
        return {"ok": False, "error": repr(exc)}


def send_home_enter(jab):
    try:
        jab.press_key("home", wait=0.02)
        jab.press_key("enter", wait=0)
        return {"ok": True, "method": "home-enter"}
    except Exception as exc:
        return {"ok": False, "method": "home-enter", "error": repr(exc)}


def commit_method(jab, vm_id, context, method, window_root=None):
    if method == "activate-enter":
        return activate_and_press(jab, window_root, ["enter"])
    if method == "activate-home-enter":
        return activate_and_press(
            jab, window_root, ["home", "enter"]
        )
    if method == "activate-esc":
        return activate_and_press(jab, window_root, ["esc"])
    if method == "embedded-select-enter":
        return select_embedded_customer_option(jab, vm_id, context, press_enter=True)
    if method == "embedded-action-enter":
        return action_embedded_customer_option(jab, vm_id, context, press_enter=True)
    if method == "embedded-select-activate-enter":
        selected = select_embedded_customer_option(jab, vm_id, context, press_enter=False)
        pressed = activate_and_press(jab, window_root, ["enter"])
        return {
            "ok": bool(selected.get("ok") and pressed.get("ok")),
            "method": method,
            "select": selected,
            "activate_press": pressed,
        }
    return send_home_enter(jab)


def activate_and_press(jab, hwnd, keys):
    activation = activate_hwnd(hwnd)
    before = jab.get_foreground_window_info()
    sent = []
    errors = []
    for key in keys:
        try:
            jab.press_key(key, wait=0.06)
            sent.append(key)
        except Exception as exc:
            errors.append({"key": key, "error": repr(exc)})
            break
    after = jab.get_foreground_window_info()
    return {
        "ok": bool(activation.get("ok") and not errors),
        "activation": activation,
        "keys": sent,
        "errors": errors,
        "foreground_before_keys": before,
        "foreground_after_keys": after,
    }


def activate_hwnd(hwnd):
    if os.name != "nt" or not hwnd:
        return {"ok": False, "reason": "missing hwnd", "hwnd": hwnd}
    user32 = ctypes.windll.user32
    hwnd_obj = wintypes.HWND(int(hwnd))
    before = int(user32.GetForegroundWindow() or 0)
    if not user32.IsWindow(hwnd_obj):
        return {"ok": False, "reason": "not a window", "hwnd": int(hwnd)}
    user32.ShowWindow(hwnd_obj, 9)
    ok = bool(user32.SetForegroundWindow(hwnd_obj))
    time.sleep(0.12)
    after = int(user32.GetForegroundWindow() or 0)
    return {
        "ok": ok or after == int(hwnd),
        "hwnd": int(hwnd),
        "foreground_before": before,
        "foreground_after": after,
        "set_foreground_ok": ok,
    }


def root_hwnd(hwnd):
    if os.name != "nt" or not hwnd:
        return 0
    return int(ctypes.windll.user32.GetAncestor(wintypes.HWND(int(hwnd)), 2) or 0)


def wait_for_any_counterparty_popup(jab, before_windows, timeout=0.45, interval=0.08):
    deadline = time.monotonic() + max(float(timeout or 0), 0.0)
    last_visible = []
    last_new = []
    while time.monotonic() < deadline:
        last_visible = collect_visible_popup_windows(jab)
        last_new = collect_new_visible_popup_windows(jab, before_windows)
        popup = find_counterparty_popup(last_visible) or find_counterparty_popup(last_new)
        if popup:
            return {
                "ok": True,
                "counterparty_popup": popup,
                "visible_count": len(last_visible),
                "new_count": len(last_new),
            }
        time.sleep(interval)
    return {
        "ok": False,
        "counterparty_popup": None,
        "visible_count": len(last_visible),
        "new_count": len(last_new),
        "reason": "counterparty popup not detected",
    }


def find_counterparty_popup(windows):
    for window in windows or []:
        labels = [
            control
            for control in window.get("all_controls", []) or []
            if str(control.get("role") or "").strip().lower() == "label"
        ]
        names = [str(item.get("name") or "").strip() for item in labels]
        if COUNTERPARTY_EXPECTED not in names:
            continue
        selected = ""
        for item in labels:
            if "selected" in str(item.get("states") or "").lower():
                selected = str(item.get("name") or "").strip()
                break
        return {
            "hwnd": window.get("hwnd"),
            "root": window.get("root"),
            "options": names,
            "selected": selected,
            "first_option": next((name for name in names if name), ""),
        }
    return None


def snapshot_combo_tree(jab, vm_id, context, depth=6):
    return {
        "root": summarize_context_node(jab, vm_id, context, "target", depth=depth),
        "selection": selection_summary(jab, vm_id, context),
    }


def summarize_context_node(jab, vm_id, context, path, depth):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return {"path": path, "ok": False, "reason": "context info not readable"}
    item = {
        "path": path,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "role": info.role_en_US.strip() or info.role.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "index_in_parent": info.indexInParent,
        "children_count": info.childrenCount,
        "bounds": [info.x, info.y, info.width, info.height],
        "accessible_component": bool(info.accessibleComponent),
        "accessible_action": bool(info.accessibleAction),
        "accessible_selection": bool(info.accessibleSelection),
        "accessible_text": bool(info.accessibleText),
        "actions": jab.get_action_names(vm_id, context)
        if info.accessibleAction
        else [],
        "selection": selection_summary(jab, vm_id, context),
        "selected_child_indexes": selected_child_indexes(
            jab, vm_id, context, info.childrenCount
        ),
        "text": safe_text_value(jab, vm_id, context),
        "children": [],
    }
    if depth <= 0:
        return item
    children = []
    owned = []
    try:
        for index in range(min(info.childrenCount, jab.max_children)):
            child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
            if not child:
                continue
            owned.append(child)
            children.append(
                summarize_context_node(
                    jab, vm_id, child, f"{path}.{index}", depth=depth - 1
                )
            )
        item["children"] = children
    finally:
        jab.release_contexts(vm_id, owned)
    return item


def safe_text_value(jab, vm_id, context):
    try:
        value = jab.get_text_context_value(vm_id, context)
    except Exception as exc:
        return {"ok": False, "error": repr(exc)}
    return {"ok": True, "value": value}


def selection_summary(jab, vm_id, context):
    if not hasattr(jab.dll, "getAccessibleSelectionCountFromContext"):
        return {"available": False}
    try:
        count = int(jab.dll.getAccessibleSelectionCountFromContext(vm_id, context))
    except Exception as exc:
        return {"available": True, "ok": False, "error": repr(exc)}
    return {"available": True, "ok": True, "count": count}


def selected_child_indexes(jab, vm_id, context, child_count):
    if not hasattr(jab.dll, "isAccessibleChildSelectedFromContext"):
        return {"available": False}
    indexes = []
    errors = []
    for index in range(min(int(child_count or 0), 80)):
        try:
            if jab.dll.isAccessibleChildSelectedFromContext(vm_id, context, index):
                indexes.append(index)
        except Exception as exc:
            errors.append({"index": index, "error": repr(exc)})
            break
    return {
        "available": True,
        "indexes": indexes,
        "errors": errors,
    }


def probe_embedded_counterparty_tree(jab, vm_id, combo_context):
    target = find_embedded_counterparty_contexts(jab, vm_id, combo_context)
    result = {
        "ok": bool(target.get("ok")),
        "reason": target.get("reason"),
        "list": target.get("list_info"),
        "popup": target.get("popup_info"),
        "customer": target.get("customer_info"),
        "labels": target.get("labels") or [],
        "selected_labels": target.get("selected_labels") or [],
        "selection": target.get("selection"),
        "selected_child_indexes": target.get("selected_child_indexes"),
    }
    release_embedded_target(jab, vm_id, target)
    return result


def find_embedded_counterparty_contexts(jab, vm_id, combo_context):
    result = {
        "ok": False,
        "reason": "embedded counterparty list not found",
        "owned_contexts": [],
    }
    best = None

    def visit(context, path, depth, ancestors):
        nonlocal best
        info = jab.get_context_info(vm_id, context)
        if not info:
            return
        role = (info.role_en_US.strip() or info.role.strip()).lower()
        children = []
        labels = []
        owned = []
        if depth > 0:
            for index in range(min(info.childrenCount, jab.max_children)):
                child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
                if not child:
                    continue
                owned.append(child)
                child_info = jab.get_context_info(vm_id, child)
                child_role = (
                    child_info.role_en_US.strip() or child_info.role.strip()
                ).lower() if child_info else ""
                child_name = child_info.name.strip() if child_info else ""
                children.append((index, child, child_info, child_role, child_name))
                if child_role == "label":
                    labels.append((index, child, child_info, child_name))

        label_names = [name for _, _, _, name in labels if name]
        if role == "list" and COUNTERPARTY_EXPECTED in label_names:
            customer = next(
                item for item in labels if item[3] == COUNTERPARTY_EXPECTED
            )
            selected_labels = [
                name
                for _, _, child_info, name in labels
                if "selected" in (
                    child_info.states_en_US.strip() or child_info.states.strip()
                ).lower()
            ]
            keep = [context, customer[1]]
            keep.extend(item[1] for item in ancestors)
            best = {
                "ok": True,
                "list_context": context,
                "customer_context": customer[1],
                "customer_index": customer[0],
                "owned_contexts": unique_contexts(keep + owned),
                "list_info": info_to_small_dict(info, path),
                "customer_info": info_to_small_dict(customer[2], f"{path}.{customer[0]}"),
                "labels": label_names,
                "selected_labels": selected_labels,
                "selection": selection_summary(jab, vm_id, context),
                "selected_child_indexes": selected_child_indexes(
                    jab, vm_id, context, info.childrenCount
                ),
            }
            popup = next(
                (
                    ancestor_info
                    for _, ancestor_info in reversed(ancestors)
                    if (
                        ancestor_info.role_en_US.strip() or ancestor_info.role.strip()
                    ).lower()
                    == "popup menu"
                ),
                None,
            )
            if popup:
                best["popup_info"] = info_to_small_dict(popup, "ancestor")
            return

        next_ancestors = ancestors + [(context, info)]
        try:
            if depth > 0 and best is None:
                for index, child, _child_info, _child_role, _child_name in children:
                    visit(child, f"{path}.{index}", depth - 1, next_ancestors)
                    if best is not None:
                        break
        finally:
            if best is None:
                jab.release_contexts(vm_id, owned)

    visit(combo_context, "target", 8, [])
    if best:
        return best
    return result


def release_embedded_target(jab, vm_id, target):
    contexts = (target or {}).get("owned_contexts") or []
    if contexts:
        jab.release_contexts(vm_id, contexts)


def unique_contexts(contexts):
    result = []
    seen = set()
    for context in contexts or []:
        key = context_key(context)
        if key in seen:
            continue
        seen.add(key)
        result.append(context)
    return result


def context_key(context):
    try:
        value = getattr(context, "value", context)
        return ("int", int(value))
    except Exception:
        return ("repr", repr(context))


def select_embedded_customer_option(jab, vm_id, combo_context, press_enter):
    target = find_embedded_counterparty_contexts(jab, vm_id, combo_context)
    try:
        if not target.get("ok"):
            return {
                "ok": False,
                "method": "embedded-select-enter",
                "reason": target.get("reason"),
            }
        if not hasattr(jab.dll, "addAccessibleSelectionFromContext"):
            return {
                "ok": False,
                "method": "embedded-select-enter",
                "reason": "selection API unavailable",
                "target": embedded_target_summary(target),
            }
        list_context = target["list_context"]
        index = int(target.get("customer_index") or 0)
        before = selected_child_indexes(
            jab, vm_id, list_context, len(target.get("labels") or [])
        )
        if hasattr(jab.dll, "clearAccessibleSelectionFromContext"):
            jab.dll.clearAccessibleSelectionFromContext(vm_id, list_context)
        jab.dll.addAccessibleSelectionFromContext(vm_id, list_context, index)
        after_select = selected_child_indexes(
            jab, vm_id, list_context, len(target.get("labels") or [])
        )
        focus = request_focus(jab, vm_id, list_context)
        enter = None
        if press_enter:
            enter = send_home_enter(jab)
        return {
            "ok": index in (after_select.get("indexes") or []),
            "method": "embedded-select-enter",
            "target": embedded_target_summary(target),
            "before_selected": before,
            "after_selected": after_select,
            "request_focus_list": focus,
            "enter": enter,
        }
    except Exception as exc:
        return {"ok": False, "method": "embedded-select-enter", "error": repr(exc)}
    finally:
        release_embedded_target(jab, vm_id, target)


def action_embedded_customer_option(jab, vm_id, combo_context, press_enter):
    target = find_embedded_counterparty_contexts(jab, vm_id, combo_context)
    try:
        if not target.get("ok"):
            return {
                "ok": False,
                "method": "embedded-action-enter",
                "reason": target.get("reason"),
            }
        customer_context = target["customer_context"]
        focus = request_focus(jab, vm_id, customer_context)
        actions = jab.get_action_names(vm_id, customer_context)
        action_ok = False
        action_name = None
        if actions:
            action_name = "单击" if "单击" in actions else actions[0]
            action_ok = bool(jab.do_action(vm_id, customer_context, action_name=action_name))
        enter = None
        if press_enter:
            enter = send_home_enter(jab)
        return {
            "ok": bool(action_ok or focus.get("ok")),
            "method": "embedded-action-enter",
            "target": embedded_target_summary(target),
            "request_focus_customer": focus,
            "actions": actions,
            "action_name": action_name,
            "action_ok": action_ok,
            "enter": enter,
        }
    except Exception as exc:
        return {"ok": False, "method": "embedded-action-enter", "error": repr(exc)}
    finally:
        release_embedded_target(jab, vm_id, target)


def embedded_target_summary(target):
    return {
        "list": target.get("list_info"),
        "popup": target.get("popup_info"),
        "customer": target.get("customer_info"),
        "customer_index": target.get("customer_index"),
        "labels": target.get("labels"),
        "selected_labels": target.get("selected_labels"),
    }


def info_to_small_dict(info, path):
    if not info:
        return None
    return {
        "path": path,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "role": info.role_en_US.strip() or info.role.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "index_in_parent": info.indexInParent,
        "children_count": info.childrenCount,
        "bounds": [info.x, info.y, info.width, info.height],
        "accessible_action": bool(info.accessibleAction),
        "accessible_selection": bool(info.accessibleSelection),
        "accessible_text": bool(info.accessibleText),
    }


def find_counterparty_in_tree(tree):
    labels = []
    walk_tree((tree or {}).get("root"), labels)
    names = [item.get("name") for item in labels if item.get("name")]
    selected = [
        item.get("name")
        for item in labels
        if "selected" in str(item.get("states") or "").lower()
    ]
    return {
        "ok": COUNTERPARTY_EXPECTED in names,
        "labels": names,
        "selected": selected,
        "first_label": next((name for name in names if name), ""),
    }


def walk_tree(node, labels):
    if not isinstance(node, dict):
        return
    if str(node.get("role") or "").strip().lower() == "label":
        labels.append(node)
    for child in node.get("children") or []:
        walk_tree(child, labels)


def cleanup_all_visible_popups(jab):
    return cleanup_popups(collect_visible_popup_windows(jab))


def cleanup_popups(windows):
    hwnds = []
    for window in windows or []:
        hwnd = window.get("hwnd")
        if hwnd:
            hwnds.append(hwnd)
    result = []
    for hwnd in list(dict.fromkeys(hwnds)):
        result.append(close_popup_hwnd(hwnd))
    return result


def finish(report, args):
    output_dir = ROOT / "logs"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / (
        f"counterparty_methods_probe_{time.strftime('%Y%m%d_%H%M%S')}.json"
    )
    output_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    report["output_path"] = str(output_path)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json:
        print(text)
    else:
        print(f"探测结果: {output_path}")
        print(text)
    return 0 if report.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
