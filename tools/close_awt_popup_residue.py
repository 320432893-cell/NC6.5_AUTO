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
        description="Close small no-title SunAwtWindow residue."
    )
    parser.add_argument("--hwnd", type=int, action="append")
    parser.add_argument("--all-small", action="store_true")
    parser.add_argument(
        "--all-disabled-small",
        action="store_true",
        help="Deprecated alias for --all-small.",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    if os.name != "nt":
        raise SystemExit("This tool must run with Windows Python.")

    targets = collect_targets(
        args.hwnd or [], args.all_small or args.all_disabled_small
    )
    if not args.dry_run:
        close_targets(targets)
    print({"targets": targets, "dry_run": args.dry_run})
    return 0


def collect_targets(hwnd_values, all_small):
    user32 = ctypes.windll.user32
    targets = []
    for hwnd in hwnd_values:
        item = describe_window(user32, hwnd)
        if item:
            targets.append(item)
    if not all_small:
        return targets

    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _lparam):
        item = describe_window(user32, int(hwnd))
        if not item:
            return True
        if (
            item["class"] == "SunAwtWindow"
            and item["title"] == ""
            and 0 < item["rect"][2] <= 250
            and 0 < item["rect"][3] <= 250
        ):
            targets.append(item)
        return True

    user32.EnumWindows(enum_proc(callback), 0)
    return targets


def close_targets(targets):
    user32 = ctypes.windll.user32
    wm_close = 0x0010
    sw_hide = 0
    swp_nosize = 0x0001
    swp_noactivate = 0x0010
    swp_hidewindow = 0x0080
    swp_noownerzorder = 0x0200
    rdw_invalidate = 0x0001
    rdw_erase = 0x0004
    rdw_allchildren = 0x0080
    rdw_updatenow = 0x0100
    for item in targets:
        hwnd = item["hwnd"]
        user32.EnableWindow(hwnd, True)
        user32.ShowWindow(hwnd, sw_hide)
        user32.SetWindowPos(
            hwnd,
            0,
            -32000,
            -32000,
            0,
            0,
            swp_nosize | swp_noactivate | swp_hidewindow | swp_noownerzorder,
        )
        user32.PostMessageW(hwnd, wm_close, 0, 0)
        item["after"] = describe_window(user32, hwnd)
    user32.RedrawWindow(
        user32.GetDesktopWindow(),
        None,
        0,
        rdw_invalidate | rdw_erase | rdw_allchildren | rdw_updatenow,
    )


def describe_window(user32, hwnd):
    if not user32.IsWindow(hwnd):
        return None
    rect = Rect()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return {
        "hwnd": int(hwnd),
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
