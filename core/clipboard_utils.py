import ctypes
import time
from ctypes import wintypes


GMEM_MOVEABLE = 0x0002
CF_UNICODETEXT = 13
_CLIPBOARD_API_CONFIGURED = False


def configure_clipboard_api():
    global _CLIPBOARD_API_CONFIGURED
    if _CLIPBOARD_API_CONFIGURED:
        return
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    user32.OpenClipboard.argtypes = [ctypes.c_void_p]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL
    user32.EmptyClipboard.argtypes = []
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE
    user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user32.SetClipboardData.restype = wintypes.HANDLE
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    _CLIPBOARD_API_CONFIGURED = True


def open_clipboard_with_retry(attempts=5, interval=0.05):
    configure_clipboard_api()
    user32 = ctypes.windll.user32
    max_attempts = max(1, int(attempts))
    for attempt in range(max_attempts):
        if user32.OpenClipboard(None):
            return True
        if attempt < max_attempts - 1:
            time.sleep(interval)
    return False


def get_clipboard_text(attempts=5, interval=0.05):
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    if not open_clipboard_with_retry(attempts=attempts, interval=interval):
        return None
    try:
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return None
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return None
        try:
            return ctypes.wstring_at(ptr)
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


def set_clipboard_text(text, attempts=5, interval=0.05):
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    encoded = str(text) + "\0"
    size = len(encoded) * ctypes.sizeof(ctypes.c_wchar)
    if not open_clipboard_with_retry(attempts=attempts, interval=interval):
        raise RuntimeError("OpenClipboard failed")
    try:
        if not user32.EmptyClipboard():
            raise RuntimeError("EmptyClipboard failed")
        handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, size)
        if not handle:
            raise RuntimeError("GlobalAlloc failed")
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            raise RuntimeError("GlobalLock failed")
        try:
            ctypes.memmove(ptr, ctypes.create_unicode_buffer(encoded), size)
        finally:
            kernel32.GlobalUnlock(handle)
        if not user32.SetClipboardData(CF_UNICODETEXT, handle):
            raise RuntimeError("SetClipboardData failed")
    finally:
        user32.CloseClipboard()


def restore_clipboard_text(text):
    if text is None:
        return False
    set_clipboard_text(text)
    return True
