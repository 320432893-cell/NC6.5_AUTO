import argparse
import ctypes
import json
import sys
from ctypes import wintypes
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.utils import load_config  # noqa: E402
from tools.jab_probe import (  # noqa: E402
    AccessibleActions,
    AccessibleContextInfo,
    JOBJECT,
    configure_jab,
    load_access_bridge,
)


class Rect(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def main():
    parser = argparse.ArgumentParser(description="Read-only AWT/JAB window probe.")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--only-awt", action="store_true")
    parser.add_argument("--children", action="store_true")
    parser.add_argument("--startup-wait", type=float, default=0.5)
    args = parser.parse_args()

    cfg = load_config(args.config)
    dll, path = load_access_bridge((cfg.get("jab") or {}).get("dll_path"))
    configure_jab(dll)
    if hasattr(dll, "initializeAccessBridge"):
        dll.initializeAccessBridge()
    import time

    time.sleep(args.startup_wait)

    report = {
        "dll": path,
        "windows": collect_windows(
            dll, only_awt=args.only_awt, include_children=args.children
        ),
    }
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)


def collect_windows(dll, only_awt=False, include_children=False):
    user32 = ctypes.windll.user32
    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    windows = []

    seen = set()

    def add_window(hwnd):
        hwnd_int = int(hwnd)
        if hwnd_int in seen:
            return
        seen.add(hwnd_int)
        class_name = read_class_name(user32, hwnd)
        if only_awt and not class_name.startswith(("SunAwt", "Yonyou")):
            return
        title = read_title(user32, hwnd)
        rect = Rect()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        item = {
            "hwnd": int(hwnd),
            "title": title,
            "class": class_name,
            "visible": bool(user32.IsWindowVisible(hwnd)),
            "enabled": bool(user32.IsWindowEnabled(hwnd)),
            "rect": [
                rect.left,
                rect.top,
                rect.right - rect.left,
                rect.bottom - rect.top,
            ],
            "is_java": bool(dll.isJavaWindow(hwnd)),
            "root": None,
            "actions": [],
        }
        if item["is_java"]:
            vm_id = ctypes.c_long()
            context = JOBJECT()
            if dll.getAccessibleContextFromHWND(
                hwnd, ctypes.byref(vm_id), ctypes.byref(context)
            ):
                item["root"] = context_summary(dll, vm_id.value, context.value)
                item["actions"] = collect_action_summaries(
                    dll, vm_id.value, context.value, max_depth=5, max_children=80
                )[:20]
        windows.append(item)

    def child_callback(hwnd, _lparam):
        add_window(hwnd)
        return True

    def callback(hwnd, _lparam):
        add_window(hwnd)
        if include_children:
            user32.EnumChildWindows(hwnd, enum_proc(child_callback), 0)
        return True

    user32.EnumWindows(enum_proc(callback), 0)
    return windows


def read_title(user32, hwnd):
    length = user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def read_class_name(user32, hwnd):
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


def context_summary(dll, vm_id, context):
    info = AccessibleContextInfo()
    if not dll.getAccessibleContextInfo(vm_id, context, ctypes.byref(info)):
        return None
    return {
        "role": info.role_en_US.strip() or info.role.strip(),
        "name": info.name.strip(),
        "description": info.description.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "children": info.childrenCount,
        "bounds": [info.x, info.y, info.width, info.height],
        "accessibleAction": bool(info.accessibleAction),
        "accessibleSelection": bool(info.accessibleSelection),
        "accessibleText": bool(info.accessibleText),
    }


def collect_action_summaries(
    dll, vm_id, context, path="0", depth=0, max_depth=5, max_children=80
):
    info = AccessibleContextInfo()
    if not dll.getAccessibleContextInfo(vm_id, context, ctypes.byref(info)):
        return []
    role = info.role_en_US.strip() or info.role.strip()
    items = []
    if info.accessibleAction:
        actions = AccessibleActions()
        names = []
        if dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
            names = [
                actions.actionInfo[index].name.strip()
                for index in range(actions.actionsCount)
            ]
        if names:
            items.append(
                {
                    "path": path,
                    "role": role,
                    "name": info.name.strip(),
                    "description": info.description.strip(),
                    "states": info.states_en_US.strip() or info.states.strip(),
                    "bounds": [info.x, info.y, info.width, info.height],
                    "actions": names,
                }
            )
    if depth >= max_depth:
        return items
    for index in range(min(info.childrenCount, max_children)):
        child = dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        items.extend(
            collect_action_summaries(
                dll,
                vm_id,
                child,
                path=f"{path}.{index}",
                depth=depth + 1,
                max_depth=max_depth,
                max_children=max_children,
            )
        )
        if hasattr(dll, "releaseJavaObject"):
            dll.releaseJavaObject(vm_id, child)
    return items


def print_text(report):
    print("dll:", report["dll"])
    for item in report["windows"]:
        print(
            f"hwnd={item['hwnd']} class={item['class']!r} title={item['title']!r} "
            f"visible={item['visible']} enabled={item['enabled']} rect={item['rect']} "
            f"is_java={item['is_java']}"
        )
        if item["root"]:
            print(f"  root={item['root']}")
        for action in item["actions"]:
            print(f"  action={action}")


if __name__ == "__main__":
    raise SystemExit(main())
