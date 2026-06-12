# 职责：封装 JAB/NC 自动化会用到的 Win32 窗口查找、激活、关闭和残留 AWT 窗口清理
# 不做什么：不加载 Java Access Bridge，不读取 JAB context，不处理表格缓存之外的业务状态
# 允许依赖层：标准库 ctypes/os/time、日志、调用方传入的缓存清理回调
# 谁不应该 import：纯数据解析、Excel/Sheet 写入和匹配模块不应 import

import ctypes
import os
from ctypes import wintypes
import time

from core.logger import log


def hide_blank_awt_windows(enabled):
    """Force-hide only hidden no-title AWT residue left by JAB/Java."""
    if not enabled or os.name != "nt":
        return []
    if not hasattr(ctypes, "windll"):
        return []

    user32 = ctypes.windll.user32
    hidden = []
    redraw_hwnds = []
    sw_hide = 0
    wm_close = 0x0010
    gwl_exstyle = -20
    ws_ex_toolwindow = 0x00000080
    ws_ex_noactivate = 0x08000000
    swp_nosize = 0x0001
    swp_noactivate = 0x0010
    swp_hidewindow = 0x0080
    swp_noownerzorder = 0x0200
    rdw_invalidate = 0x0001
    rdw_erase = 0x0004
    rdw_allchildren = 0x0080
    rdw_updatenow = 0x0100

    class Rect(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _lparam):
        length = user32.GetWindowTextLengthW(hwnd)
        title = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, title, length + 1)

        class_name = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_name, 256)

        rect = Rect()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        width = rect.right - rect.left
        height = rect.bottom - rect.top

        if class_name.value in ("SunAwtFrame", "SunAwtCanvas"):
            redraw_hwnds.append(hwnd)
            return True

        if (
            class_name.value == "SunAwtWindow"
            and title.value == ""
            and 0 < width <= 250
            and 0 < height <= 250
        ):
            before_visible = bool(user32.IsWindowVisible(hwnd))
            if before_visible:
                return True
            style = user32.GetWindowLongW(hwnd, gwl_exstyle)
            user32.SetWindowLongW(
                hwnd, gwl_exstyle, style | ws_ex_toolwindow | ws_ex_noactivate
            )
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
            hidden.append(
                {
                    "hwnd": int(hwnd),
                    "before_visible": before_visible,
                    "left": rect.left,
                    "top": rect.top,
                    "width": width,
                    "height": height,
                }
            )

        return True

    user32.EnumWindows(enum_proc(callback), 0)
    redraw_hwnds.append(user32.GetDesktopWindow())
    for hwnd in redraw_hwnds:
        user32.RedrawWindow(
            hwnd,
            None,
            0,
            rdw_invalidate | rdw_erase | rdw_allchildren | rdw_updatenow,
        )
    if hidden:
        log.info(f"JAB 已隐藏空白 AWT 浮窗: {hidden}")
    return hidden


def close_window_by_title(
    title, class_name=None, wait=None, menu_wait=0.5, clear_table_cache=None
):
    if os.name != "nt":
        return False

    user32 = ctypes.windll.user32
    wm_close = 0x0010
    closed = []
    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True

        length = user32.GetWindowTextLengthW(hwnd)
        window_title = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, window_title, length + 1)

        window_class = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, window_class, 256)

        if window_title.value != title:
            return True
        if class_name and window_class.value != class_name:
            return True

        user32.PostMessageW(hwnd, wm_close, 0, 0)
        closed.append(
            {
                "hwnd": int(hwnd),
                "title": window_title.value,
                "class": window_class.value,
            }
        )
        return True

    user32.EnumWindows(enum_proc(callback), 0)
    if closed:
        log.info(f"已关闭窗口: {closed}")
        time.sleep(menu_wait if wait is None else wait)
        if clear_table_cache:
            clear_table_cache(title)
        return True

    log.info(f"未找到需关闭窗口: title={title} class={class_name}")
    return False


def activate_window_by_title(title, class_name=None, timeout=None, search_timeout=5.0):
    if os.name != "nt":
        return False

    deadline = time.time() + (timeout or search_timeout)
    while time.time() < deadline:
        hwnd = find_window_handle(title, class_name=class_name, visible_only=False)
        if hwnd:
            user32 = ctypes.windll.user32
            user32.ShowWindow(hwnd, 9)
            user32.SetForegroundWindow(hwnd)
            time.sleep(0.2)
            return True
        time.sleep(0.2)
    return False


def get_foreground_window_info():
    if os.name != "nt":
        return None

    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None

    length = user32.GetWindowTextLengthW(hwnd)
    window_title = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, window_title, length + 1)

    window_class = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, window_class, 256)
    return {
        "hwnd": int(hwnd),
        "title": window_title.value,
        "class_name": window_class.value,
    }


def foreground_window_matches(title, class_name=None):
    info = get_foreground_window_info()
    if not info:
        return False
    if info["title"] != title:
        return False
    if class_name and info["class_name"] != class_name:
        return False
    return True


def wait_window_by_title(
    title,
    class_name=None,
    timeout=None,
    include_children=False,
    visible_only=True,
    search_timeout=5.0,
):
    deadline = time.time() + (timeout or search_timeout)
    while time.time() < deadline:
        hwnd = find_window_handle(
            title,
            class_name=class_name,
            visible_only=visible_only,
            include_children=include_children,
        )
        if hwnd:
            return hwnd
        time.sleep(0.2)
    return None


def window_exists(title, class_name=None, include_children=False):
    return bool(
        find_window_handle(
            title,
            class_name=class_name,
            visible_only=True,
            include_children=include_children,
        )
    )


def find_window_handle(
    title, class_name=None, visible_only=True, include_children=False
):
    if os.name != "nt":
        return None

    user32 = ctypes.windll.user32
    found = []
    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _lparam):
        if visible_only and not user32.IsWindowVisible(hwnd):
            return True

        length = user32.GetWindowTextLengthW(hwnd)
        window_title = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, window_title, length + 1)

        window_class = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, window_class, 256)

        if window_title.value != title:
            return True
        if class_name and window_class.value != class_name:
            return True

        found.append(hwnd)
        return False

    def child_callback(hwnd, _lparam):
        return callback(hwnd, _lparam)

    user32.EnumWindows(enum_proc(callback), 0)
    if include_children and not found:
        user32.EnumWindows(
            enum_proc(
                lambda hwnd, _lparam: (
                    user32.EnumChildWindows(hwnd, enum_proc(child_callback), 0) or True
                )
            ),
            0,
        )
    return found[0] if found else None
