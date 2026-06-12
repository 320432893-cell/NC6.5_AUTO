import argparse
import ctypes
import json
from ctypes import wintypes


GWL_EXSTYLE = -20
GW_OWNER = 4


def main():
    parser = argparse.ArgumentParser(
        description="Read-only Win32 visible window overlay probe."
    )
    parser.add_argument(
        "--all", action="store_true", help="Include invisible/minimized windows."
    )
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    windows = collect_windows(include_all=args.all)
    ranked = sorted(
        windows, key=lambda item: (item["rect"][1], item["rect"][0], item["hwnd"])
    )
    if args.json:
        print(json.dumps(ranked, ensure_ascii=False, indent=2))
    else:
        for item in ranked:
            marker = "CANDIDATE" if is_overlay_candidate(item) else "         "
            print(
                f"{marker} hwnd={item['hwnd']} pid={item['pid']} visible={item['visible']} "
                f"class={item['class']!r} title={item['title']!r} rect={item['rect']} "
                f"owner={item['owner']} exstyle=0x{item['exstyle']:08x}"
            )
    return 0


def collect_windows(include_all=False):
    user32 = ctypes.windll.user32
    windows = []

    class Rect(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _lparam):
        visible = bool(user32.IsWindowVisible(hwnd))
        if not include_all and not visible:
            return True

        rect = Rect()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        width = rect.right - rect.left
        height = rect.bottom - rect.top
        if not include_all and (width <= 0 or height <= 0):
            return True

        title_len = user32.GetWindowTextLengthW(hwnd)
        title = ctypes.create_unicode_buffer(title_len + 1)
        user32.GetWindowTextW(hwnd, title, title_len + 1)
        class_name = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_name, 256)
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

        windows.append(
            {
                "hwnd": int(hwnd),
                "pid": int(pid.value),
                "visible": visible,
                "enabled": bool(user32.IsWindowEnabled(hwnd)),
                "class": class_name.value,
                "title": title.value,
                "rect": [rect.left, rect.top, rect.right, rect.bottom],
                "width": width,
                "height": height,
                "owner": int(user32.GetWindow(hwnd, GW_OWNER)),
                "exstyle": int(user32.GetWindowLongW(hwnd, GWL_EXSTYLE)),
            }
        )
        return True

    user32.EnumWindows(enum_proc(callback), 0)
    return windows


def is_overlay_candidate(item):
    if not item["visible"]:
        return False
    if item["width"] <= 0 or item["height"] <= 0:
        return False
    if item["class"] in {"Progman", "WorkerW", "Shell_TrayWnd"}:
        return False
    if item["title"]:
        return False
    x1, y1, x2, y2 = item["rect"]
    if x2 < 0 or y2 < 0 or x1 > 3000 or y1 > 2000:
        return False
    return item["width"] <= 900 and item["height"] <= 700


if __name__ == "__main__":
    raise SystemExit(main())
