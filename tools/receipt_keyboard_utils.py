# 职责: 提供收款单正式明细输入需要的停止键、剪贴板、受保护按键和金额比较工具
# 不做什么: 不打开/操作参照窗口，不执行 NC 业务流程，不读取 Excel/配置
# 允许依赖层: 标准库 ctypes/decimal/sys/time，core.clipboard_utils
# 谁不应该 import: core 层模块不应 import；查询/Sheet 写入模块不应 import

import ctypes
from ctypes import wintypes
from decimal import Decimal, InvalidOperation
import sys
import time

from core.clipboard_utils import (
    configure_clipboard_api,
    get_clipboard_text,
    restore_clipboard_text,
    set_clipboard_text,
)

STOP_HOTKEY = "Space"
VK_SPACE = 0x20
VK_CONTROL = 0x11
VK_A = 0x41
VK_D = 0x44
VK_I = 0x49
VK_N = 0x4E
VK_Q = 0x51
VK_C = 0x43
VK_S = 0x53
VK_V = 0x56
VK_Y = 0x59
VK_MENU = 0x12


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("mi", MOUSEINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("union", INPUT_UNION)]

    @property
    def ki(self):
        return self.union.ki

    @ki.setter
    def ki(self, value):
        self.union.ki = value

    @property
    def mi(self):
        return self.union.mi

    @mi.setter
    def mi(self, value):
        self.union.mi = value

def is_stop_hotkey_pressed():
    if not hasattr(ctypes, "windll"):
        return False
    user32 = ctypes.windll.user32
    return bool(user32.GetAsyncKeyState(VK_SPACE) & 0x8000)

def normalize_amount(value):
    if value is None:
        return None
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def amount_matches(actual, expected):
    actual_amount = normalize_amount(actual)
    expected_amount = normalize_amount(expected)
    return (
        actual_amount is not None
        and expected_amount is not None
        and actual_amount == expected_amount
    )


def read_window_info(hwnd):
    if sys.platform != "win32" or not hwnd:
        return None
    user32 = ctypes.windll.user32
    hwnd = int(hwnd)
    length = user32.GetWindowTextLengthW(hwnd)
    title = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, title, length + 1)
    class_name = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_name, 256)
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    root = user32.GetAncestor(hwnd, 2)
    return {
        "hwnd": hwnd,
        "title": title.value,
        "class_name": class_name.value,
        "pid": int(pid.value),
        "visible": bool(user32.IsWindowVisible(hwnd)),
        "root_hwnd": int(root) if root else None,
    }


def send_input(inp):
    ctypes.windll.kernel32.SetLastError(0)
    sent = ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
    if sent != 1:
        error_code = ctypes.windll.kernel32.GetLastError()
        raise RuntimeError(f"SendInput failed, error={error_code}")


def send_virtual_key(vk, key_up=False):
    inp = INPUT()
    inp.type = 1
    inp.ki = KEYBDINPUT(vk, 0, 0x0002 if key_up else 0, 0, None)
    send_input(inp)


def send_unicode_char(char):
    code = ord(char)
    inp = INPUT()
    inp.type = 1
    inp.ki = KEYBDINPUT(0, code, 0x0004, 0, None)
    send_input(inp)
    inp_up = INPUT()
    inp_up.type = 1
    inp_up.ki = KEYBDINPUT(0, code, 0x0004 | 0x0002, 0, None)
    send_input(inp_up)


def send_text(text):
    for char in str(text):
        send_unicode_char(char)


def send_hotkey_ctrl_a():
    send_virtual_key(VK_CONTROL, key_up=False)
    send_virtual_key(VK_A, key_up=False)
    send_virtual_key(VK_A, key_up=True)
    send_virtual_key(VK_CONTROL, key_up=True)


def send_hotkey_ctrl_v():
    send_virtual_key(VK_CONTROL, key_up=False)
    send_virtual_key(VK_V, key_up=False)
    send_virtual_key(VK_V, key_up=True)
    send_virtual_key(VK_CONTROL, key_up=True)


def send_hotkey_ctrl_i():
    send_virtual_key(VK_CONTROL, key_up=False)
    send_virtual_key(VK_I, key_up=False)
    send_virtual_key(VK_I, key_up=True)
    send_virtual_key(VK_CONTROL, key_up=True)


def send_hotkey_ctrl_d():
    send_virtual_key(VK_CONTROL, key_up=False)
    send_virtual_key(VK_D, key_up=False)
    send_virtual_key(VK_D, key_up=True)
    send_virtual_key(VK_CONTROL, key_up=True)


def send_hotkey_ctrl_s():
    send_virtual_key(VK_CONTROL, key_up=False)
    send_virtual_key(VK_S, key_up=False)
    send_virtual_key(VK_S, key_up=True)
    send_virtual_key(VK_CONTROL, key_up=True)


def send_hotkey_ctrl_q():
    send_virtual_key(VK_CONTROL, key_up=False)
    send_virtual_key(VK_Q, key_up=False)
    send_virtual_key(VK_Q, key_up=True)
    send_virtual_key(VK_CONTROL, key_up=True)


def send_hotkey_alt_c():
    send_virtual_key(VK_MENU, key_up=False)
    send_virtual_key(VK_C, key_up=False)
    send_virtual_key(VK_C, key_up=True)
    send_virtual_key(VK_MENU, key_up=True)


def send_hotkey_alt_y():
    send_virtual_key(VK_MENU, key_up=False)
    send_virtual_key(VK_Y, key_up=False)
    send_virtual_key(VK_Y, key_up=True)
    send_virtual_key(VK_MENU, key_up=True)


def send_hotkey_alt_n():
    send_virtual_key(VK_MENU, key_up=False)
    send_virtual_key(VK_N, key_up=False)
    send_virtual_key(VK_N, key_up=True)
    send_virtual_key(VK_MENU, key_up=True)


def same_window_root(left, right):
    return bool(
        left
        and right
        and (
            left.get("hwnd") == right.get("root_hwnd")
            or left.get("root_hwnd") == right.get("root_hwnd")
            or left.get("hwnd") == right.get("hwnd")
        )
    )


def foreground_matches_window(target_window):
    if sys.platform != "win32":
        return {"ok": False, "reason": "必须在 Windows Python 下运行"}
    target_info = read_window_info((target_window or {}).get("hwnd"))
    foreground_info = read_window_info(ctypes.windll.user32.GetForegroundWindow())
    if not target_info or not foreground_info:
        return {
            "ok": False,
            "reason": "无法读取当前前台窗口或目标窗口",
            "target_window": target_info,
            "foreground": foreground_info,
        }
    same_root = same_window_root(foreground_info, target_info)
    return {
        "ok": same_root,
        "reason": None if same_root else "当前前台窗口不是目标 NC 窗口",
        "target_window": target_info,
        "foreground": foreground_info,
    }


def guarded_send_table_hotkey(table_window, key_name, sender, settle_seconds=0.8):
    guard = foreground_matches_window(table_window)
    if not guard.get("ok"):
        return {
            **guard,
            "ok": False,
            "key": key_name,
            "reason": f"{guard.get('reason')}，未发送 {key_name}",
        }
    try:
        sender()
    except Exception as exc:
        return {
            **guard,
            "ok": False,
            "mode": f"SendInput({key_name})",
            "key": key_name,
            "reason": f"{type(exc).__name__}: {exc}",
        }
    if settle_seconds and float(settle_seconds) > 0:
        time.sleep(float(settle_seconds))
    return {
        **guard,
        "ok": True,
        "mode": f"SendInput({key_name})",
        "key": key_name,
    }


def guarded_send_ctrl_i(table_window):
    return guarded_send_table_hotkey(
        table_window,
        "Ctrl+I",
        send_hotkey_ctrl_i,
        settle_seconds=0.0,
    )


def guarded_send_ctrl_d(table_window):
    return guarded_send_table_hotkey(
        table_window,
        "Ctrl+D",
        send_hotkey_ctrl_d,
        settle_seconds=0.0,
    )
