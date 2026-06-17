# 职责：Win32 窗口句柄工具——根句柄/标题/类名/可见弹窗判定/关闭/边界
# 不做什么：不做 CLI 解析/不起 JABOperator(那是 receipt_new_probe 主入口)
# 允许依赖层：标准库、core JAB(经 jab 参数)、tools.jab_probe、tools.receipt_new_* 同层
# 谁不应该 import：core 层模块不应 import

import ctypes
import os
import sys
from ctypes import wintypes




class _ProbeNamespace:
    # 调用时从已加载的 receipt_new_probe 读顶层函数,使测试对
    # tools.receipt_new_probe.<name> 的 monkeypatch 与拆分前一致生效,且不在加载期 import 成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_new_probe"], name)


_probe = _ProbeNamespace()


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
    before = _probe.describe_hwnd(user32, hwnd_obj)
    if not before.get("exists"):
        return {"ok": True, "reason": "already gone", "before": before}
    if before.get("class_name") != "SunAwtWindow":
        return {"ok": False, "reason": "class mismatch", "before": before}
    user32.ShowWindow(hwnd_obj, 0)
    user32.SetWindowPos(
        hwnd_obj, 0, -32000, -32000, 0, 0, 0x0001 | 0x0010 | 0x0080 | 0x0200
    )
    user32.PostMessageW(hwnd_obj, 0x0010, 0, 0)
    return {"ok": True, "before": before, "after": _probe.describe_hwnd(user32, hwnd_obj)}



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
