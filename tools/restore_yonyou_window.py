import argparse
import ctypes
import os
from ctypes import wintypes


class Rect(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def main():
    parser = argparse.ArgumentParser(
        description="Restore off-screen Yonyou UClient windows."
    )
    parser.add_argument("--x", type=int, default=0)
    parser.add_argument("--y", type=int, default=0)
    parser.add_argument("--width", type=int, default=1280)
    parser.add_argument("--height", type=int, default=696)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    if os.name != "nt":
        raise SystemExit("This tool must run with Windows Python.")

    windows = collect_yonyou_windows()
    restored = []
    for item in windows:
        rect = item["rect"]
        offscreen = rect[0] < -1000 or rect[1] < -1000
        if not offscreen:
            continue
        if not args.dry_run:
            restore_window(item["hwnd"], args.x, args.y, args.width, args.height)
            item["after"] = read_window(item["hwnd"])
        restored.append(item)

    print({"windows": windows, "restored": restored, "dry_run": args.dry_run})
    return 0


def collect_yonyou_windows():
    user32 = ctypes.windll.user32
    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    result = []

    def callback(hwnd, _lparam):
        title = read_title(user32, hwnd)
        class_name = read_class_name(user32, hwnd)
        if (
            class_name in ("YonyouUWnd", "SunAwtFrame", "SunAwtCanvas")
            or title == "Yonyou UClient"
        ):
            rect = Rect()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            result.append(
                {
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
                }
            )
        return True

    user32.EnumWindows(enum_proc(callback), 0)
    return result


def restore_window(hwnd, x, y, width, height):
    user32 = ctypes.windll.user32
    sw_restore = 9
    swp_nozorder = 0x0004
    swp_showwindow = 0x0040
    user32.ShowWindow(hwnd, sw_restore)
    user32.SetWindowPos(hwnd, 0, x, y, width, height, swp_nozorder | swp_showwindow)
    user32.SetForegroundWindow(hwnd)


def read_window(hwnd):
    user32 = ctypes.windll.user32
    rect = Rect()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return {
        "title": read_title(user32, hwnd),
        "class": read_class_name(user32, hwnd),
        "visible": bool(user32.IsWindowVisible(hwnd)),
        "enabled": bool(user32.IsWindowEnabled(hwnd)),
        "rect": [rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top],
    }


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
