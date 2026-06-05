import argparse
import ctypes
from ctypes import wintypes


def main():
    parser = argparse.ArgumentParser(
        description="Close one Win32 window by hwnd with optional guards."
    )
    parser.add_argument("hwnd", type=int)
    parser.add_argument("--class-name", default=None)
    parser.add_argument("--title", default=None)
    parser.add_argument("--hide-first", action="store_true")
    args = parser.parse_args()

    user32 = ctypes.windll.user32
    hwnd = wintypes.HWND(args.hwnd)
    before = describe_window(user32, hwnd)
    if not before["exists"]:
        print({"ok": True, "reason": "already gone", "before": before})
        return 0

    if args.class_name is not None and before["class"] != args.class_name:
        print({"ok": False, "reason": "class mismatch", "before": before})
        return 1
    if args.title is not None and before["title"] != args.title:
        print({"ok": False, "reason": "title mismatch", "before": before})
        return 1

    if args.hide_first:
        user32.ShowWindow(hwnd, 0)
        user32.SetWindowPos(
            hwnd, 0, -32000, -32000, 0, 0, 0x0001 | 0x0010 | 0x0080 | 0x0200
        )

    user32.PostMessageW(hwnd, 0x0010, 0, 0)
    after = describe_window(user32, hwnd)
    print({"ok": True, "before": before, "after": after, "hide_first": args.hide_first})
    return 0


def describe_window(user32, hwnd):
    exists = bool(user32.IsWindow(hwnd))
    if not exists:
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
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return {
        "exists": True,
        "hwnd": int(hwnd.value),
        "pid": int(pid.value),
        "visible": bool(user32.IsWindowVisible(hwnd)),
        "class": class_name.value,
        "title": title.value,
        "rect": [rect.left, rect.top, rect.right, rect.bottom],
    }


if __name__ == "__main__":
    raise SystemExit(main())
