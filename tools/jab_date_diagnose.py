import argparse
import ctypes
import json
import os
import subprocess
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_probe import (  # noqa: E402
    AccessibleActions,
    AccessibleActionsToDo,
    JOBJECT,
    configure_jab,
    enum_windows,
    hwnd_int,
    load_access_bridge,
    run_windows_access_bridge,
)


DATE_PATH = "0.0.1.0.0.0.0.0.0.0.2.1"


def main():
    parser = argparse.ArgumentParser(description="诊断全局业务日期 JAB 断连问题")
    parser.add_argument("--dll", default=None)
    parser.add_argument("--path", default=DATE_PATH)
    parser.add_argument("--date-name", default=None)
    parser.add_argument("--name-depth", type=int, default=16)
    parser.add_argument("--max-children", type=int, default=200)
    parser.add_argument(
        "--action",
        choices=("none", "focus", "click-async", "click-blocking"),
        default="none",
    )
    parser.add_argument("--action-wait", type=float, default=0.5)
    parser.add_argument("--startup-wait", type=float, default=1.0)
    parser.add_argument("--include-children", action="store_true")
    parser.add_argument(
        "--restore-uclient",
        action="store_true",
        help="Restore and activate the Yonyou UClient window before probing.",
    )
    args = parser.parse_args()

    if os.name != "nt":
        raise SystemExit("This script must run with Windows Python.")

    report = {
        "active_window": active_window_title(),
        "python_processes": list_python_probe_processes(),
        "snapshots": [],
    }
    if args.restore_uclient:
        report["restore_uclient"] = restore_uclient_window()
        report["active_window_after_restore"] = active_window_title()

    dll, dll_path = load_access_bridge(args.dll)
    configure_jab(dll)
    stop_pump = None
    pump_thread = None
    if hasattr(dll, "initializeAccessBridge"):
        dll.initializeAccessBridge()
        report["init_mode"] = "initializeAccessBridge"
    else:
        stop_pump = threading.Event()
        pump_thread = threading.Thread(
            target=run_windows_access_bridge, args=(dll, stop_pump), daemon=True
        )
        pump_thread.start()
        report["init_mode"] = "Windows_run"
    report["dll"] = dll_path

    try:
        time.sleep(args.startup_wait)
        before = build_snapshot(
            dll,
            args.path,
            date_name=args.date_name,
            include_children=args.include_children,
            name_depth=args.name_depth,
            max_children=args.max_children,
        )
        report["snapshots"].append({"label": "before", **before})

        action_result = None
        if args.action != "none":
            action_result = run_action(dll, before.get("date_context"), args)
            time.sleep(args.action_wait)
            report["active_window_after_action"] = active_window_title()
            report["python_processes_after_action"] = list_python_probe_processes()
            after = build_snapshot(
                dll,
                args.path,
                date_name=args.date_name,
                include_children=args.include_children,
                name_depth=args.name_depth,
                max_children=args.max_children,
            )
            report["snapshots"].append({"label": "after", **after})
        report["action_result"] = action_result

        print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    finally:
        if stop_pump:
            stop_pump.set()
        if pump_thread:
            pump_thread.join(timeout=1)


def build_snapshot(
    dll,
    date_path,
    date_name=None,
    include_children=False,
    name_depth=16,
    max_children=200,
):
    windows = []
    java_windows = []
    date_context = None
    date_candidates = []
    for hwnd, title, class_name, pid, visible in enum_windows(
        include_children=include_children
    ):
        hwnd_value = hwnd_int(hwnd)
        is_java = safe_is_java_window(dll, hwnd)
        item = {
            "hwnd": hwnd_value,
            "title": title,
            "class": class_name,
            "pid": pid,
            "visible": bool(visible),
            "is_java": is_java,
        }
        if is_relevant_window(title, class_name, is_java):
            windows.append(item)
        if not is_java:
            continue
        java_windows.append(item)
        found = find_by_path(dll, hwnd, date_path)
        if found and found.get("context") and not date_context:
            date_context = found
        if date_name:
            date_candidates.extend(
                collect_name_candidates(
                    dll,
                    hwnd,
                    item,
                    date_name,
                    max_depth=name_depth,
                    max_children=max_children,
                )
            )

    return {
        "windows": windows,
        "java_windows": java_windows,
        "date_button": public_context(date_context),
        "date_candidates": sorted(
            date_candidates,
            key=lambda candidate: (-candidate["score"], candidate["path"]),
        ),
        "_date_context_found": bool(date_context),
        "date_context": date_context,
    }


def is_relevant_window(title, class_name, is_java):
    if is_java:
        return True
    if class_name.startswith("SunAwt"):
        return True
    if title in {"Yonyou UClient", "切换业务日期"}:
        return True
    return False


def safe_is_java_window(dll, hwnd):
    try:
        return bool(dll.isJavaWindow(hwnd))
    except Exception:
        return False


def find_by_path(dll, hwnd, path):
    vm_id = ctypes.c_long()
    root = JOBJECT()
    if not dll.getAccessibleContextFromHWND(
        hwnd, ctypes.byref(vm_id), ctypes.byref(root)
    ):
        return None

    context = root.value
    owned = []
    parts = [part for part in path.split(".") if part != ""]
    for offset, part in enumerate(parts):
        try:
            index = int(part)
        except ValueError:
            return None
        if offset == 0 and index == 0:
            continue
        child = dll.getAccessibleChildFromContext(vm_id.value, context, index)
        if not child:
            return None
        owned.append(child)
        context = child

    info = get_info(dll, vm_id.value, context)
    actions = get_actions(dll, vm_id.value, context)
    return {
        "vm_id": vm_id.value,
        "context": context,
        "owned": owned,
        "info": info,
        "actions": actions,
    }


def collect_name_candidates(
    dll,
    hwnd,
    window_item,
    target_name,
    max_depth=16,
    max_children=200,
):
    vm_id = ctypes.c_long()
    root = JOBJECT()
    if not dll.getAccessibleContextFromHWND(
        hwnd, ctypes.byref(vm_id), ctypes.byref(root)
    ):
        return []

    candidates = []
    walk_name_candidates(
        dll,
        vm_id.value,
        root.value,
        target_name,
        candidates,
        path="0",
        depth=0,
        max_depth=max_depth,
        max_children=max_children,
        window_item=window_item,
    )
    return candidates


def walk_name_candidates(
    dll,
    vm_id,
    context,
    target_name,
    candidates,
    path,
    depth,
    max_depth,
    max_children,
    window_item,
):
    info = get_info(dll, vm_id, context)
    if not info:
        return

    if info["name"] == target_name or info["description"] == target_name:
        actions = get_actions(dll, vm_id, context)
        candidates.append(
            {
                "path": path,
                "window": window_item,
                "info": info,
                "actions": actions,
                "score": candidate_score(window_item, info),
            }
        )

    if depth >= max_depth:
        return

    for index in range(min(info["children"], max_children)):
        child = dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        walk_name_candidates(
            dll,
            vm_id,
            child,
            target_name,
            candidates,
            path=f"{path}.{index}",
            depth=depth + 1,
            max_depth=max_depth,
            max_children=max_children,
            window_item=window_item,
        )


def candidate_score(window_item, info):
    states = info["states"].lower()
    x, y, width, height = info["bounds"]
    score = 0
    if window_item["visible"]:
        score += 10
    if "visible" in states:
        score += 10
    if "showing" in states:
        score += 30
    if width > 0 and height > 0 and x >= 0 and y >= 0:
        score += 30
    if info["role"] == "push button":
        score += 5
    if window_item["class"] == "SunAwtCanvas":
        score += 5
    return score


def get_info(dll, vm_id, context):
    from core.jab_probe import AccessibleContextInfo

    info = AccessibleContextInfo()
    if not dll.getAccessibleContextInfo(vm_id, context, ctypes.byref(info)):
        return None
    return {
        "name": info.name.strip(),
        "description": info.description.strip(),
        "role": info.role_en_US.strip() or info.role.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "bounds": [info.x, info.y, info.width, info.height],
        "children": info.childrenCount,
        "accessibleAction": bool(info.accessibleAction),
    }


def get_actions(dll, vm_id, context):
    if not hasattr(dll, "getAccessibleActions"):
        return []
    actions = AccessibleActions()
    if not dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
        return []
    return [
        actions.actionInfo[index].name.strip() for index in range(actions.actionsCount)
    ]


def public_context(context):
    if not context:
        return None
    return {
        "vm_id": context.get("vm_id"),
        "info": context.get("info"),
        "actions": context.get("actions"),
    }


def run_action(dll, context, args):
    if not context:
        return {"ok": False, "reason": "date context not found"}
    vm_id = context["vm_id"]
    jab_context = context["context"]

    if args.action == "focus":
        if not hasattr(dll, "requestFocus"):
            return {"ok": False, "reason": "requestFocus unavailable"}
        return {"ok": bool(dll.requestFocus(vm_id, jab_context)), "mode": "focus"}

    if not hasattr(dll, "doAccessibleActions"):
        return {"ok": False, "reason": "doAccessibleActions unavailable"}

    def call_action(result):
        todo = AccessibleActionsToDo()
        todo.actionsCount = 1
        todo.actions[0].name = "单击"
        failure = ctypes.c_int(-1)
        result["returned"] = bool(
            dll.doAccessibleActions(
                vm_id, jab_context, ctypes.byref(todo), ctypes.byref(failure)
            )
        )
        result["failure"] = failure.value

    result = {"mode": args.action, "returned": None, "failure": None}
    if args.action == "click-blocking":
        call_action(result)
        return result

    thread = threading.Thread(target=call_action, args=(result,), daemon=True)
    thread.start()
    thread.join(args.action_wait)
    result["thread_alive"] = thread.is_alive()
    return result


def active_window_title():
    try:
        import pygetwindow as gw

        return gw.getActiveWindowTitle()
    except Exception as exc:
        return f"<unavailable: {exc}>"


def restore_uclient_window():
    try:
        import pygetwindow as gw

        windows = gw.getWindowsWithTitle("Yonyou UClient")
        if not windows:
            return {"ok": False, "reason": "Yonyou UClient window not found"}
        window = windows[0]
        before = window_state(window)
        window.restore()
        time.sleep(0.3)
        window.activate()
        time.sleep(0.5)
        return {"ok": True, "before": before, "after": window_state(window)}
    except Exception as exc:
        return {"ok": False, "reason": str(exc)}


def window_state(window):
    return {
        "title": window.title,
        "left": window.left,
        "top": window.top,
        "width": window.width,
        "height": window.height,
        "visible": bool(window.visible),
        "isMinimized": bool(window.isMinimized),
    }


def list_python_probe_processes():
    try:
        completed = subprocess.run(
            [
                "wmic",
                "process",
                "where",
                "CommandLine like '%jab_date_diagnose.py%' or CommandLine like '%jab_probe.py%'",
                "get",
                "ProcessId,CommandLine",
                "/format:csv",
            ],
            check=False,
            capture_output=True,
            text=True,
            errors="replace",
            timeout=3,
        )
    except Exception as exc:
        return [{"error": str(exc)}]
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    return lines[1:]


if __name__ == "__main__":
    main()
