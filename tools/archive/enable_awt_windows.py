import argparse
import ctypes
import os
from ctypes import wintypes


def main():
    parser = argparse.ArgumentParser(
        description="Enable visible small SunAwtWindow popups."
    )
    parser.add_argument("--hwnd", type=int, default=None)
    parser.add_argument("--all-visible-small", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    if os.name != "nt":
        raise SystemExit("This tool must run with Windows Python.")

    targets = collect_targets(args.hwnd, args.all_visible_small)
    if not args.dry_run:
        user32 = ctypes.windll.user32
        for item in targets:
            user32.EnableWindow(item["hwnd"], True)
            item["enabled_after"] = bool(user32.IsWindowEnabled(item["hwnd"]))
    print({"targets": normalize_targets(targets), "dry_run": args.dry_run})
    return 0


def collect_targets(hwnd, all_visible_small):
    user32 = ctypes.windll.user32
    result = []
    if hwnd:
        item = describe_window(user32, hwnd)
        if item:
            result.append(item)
        return result
    if not all_visible_small:
        return result

    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(current_hwnd, _lparam):
        item = describe_window(user32, int(current_hwnd))
        if not item:
            return True
        if (
            item["class"] == "SunAwtWindow"
            and item["title"] == ""
            and item["visible"]
            and 0 < item["rect"][2] <= 250
            and 0 < item["rect"][3] <= 250
        ):
            result.append(item)
        return True

    user32.EnumWindows(enum_proc(callback), 0)
    return result


def describe_window(user32, hwnd):
    if not user32.IsWindow(hwnd):
        return None
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return {
        "hwnd": int(hwnd),
        "title": read_title(user32, hwnd),
        "class": read_class_name(user32, hwnd),
        "visible": bool(user32.IsWindowVisible(hwnd)),
        "enabled": bool(user32.IsWindowEnabled(hwnd)),
        "rect": [rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top],
    }


def normalize_targets(targets):
    return [
        {
            **item,
            "hwnd": int(item["hwnd"]),
        }
        for item in targets
    ]


def read_title(user32, hwnd):
    length = user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def read_class_name(user32, hwnd):
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


if __name__ == "__main__":
    raise SystemExit(main())
