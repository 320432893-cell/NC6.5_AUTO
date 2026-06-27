# 职责：收款单 body 表/表头 scope 的运行内定位缓存与 entry_state 解析
# 不做什么：不写入 NC、不做业务编排、不读 Excel
# 允许依赖层：core JAB、tools 下 body_table_locator;谁不应 import：core 层

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools.receipt_body_table_locator import locate_receipt_body_table_cached  # noqa: E402
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    receipt_header_dynamic_prefix,
)


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

_BODY_TABLE_SUFFIX_CACHE = {}

def cache_receipt_header_scope(jab, shared_cache, scope):
    if not isinstance(scope, dict) or not scope.get("ok"):
        return
    cached = dict(scope)
    try:
        setattr(jab, "_receipt_header_scope_cache", cached)
    except AttributeError:
        pass
    if isinstance(shared_cache, dict):
        shared_cache.clear()
        shared_cache.update(cached)

def build_body_table_cached_path(dynamic_index, scope_hwnd=None):
    if dynamic_index is None:
        return None
    cached = _BODY_TABLE_SUFFIX_CACHE.get(body_table_cache_key(scope_hwnd))
    if not cached:
        cached = _BODY_TABLE_SUFFIX_CACHE.get("last")
    suffix = (cached or {}).get("suffix")
    if not suffix:
        return None
    path = f"{receipt_header_dynamic_prefix(dynamic_index)}.{suffix}"
    return {
        "best": {
            "path": path,
            "window": {
                "hwnd": scope_hwnd,
                "class_name": "SunAwtCanvas",
            },
        }
    }

def resolve_body_table_by_dynamic_prefix(jab, dynamic_index, scope_hwnd=None):
    cached = build_body_table_cached_path(dynamic_index, scope_hwnd=scope_hwnd)
    if cached:
        located = locate_receipt_body_table_cached(
            jab,
            cached=cached,
            max_rows=5,
            scope_hwnd=scope_hwnd,
        )
        if located.get("cache_hit"):
            return {
                **located,
                "source": "learned-body-table-path",
                "cached_path": (cached or {}).get("best"),
            }

    located = locate_receipt_body_table_cached(
        jab,
        cached=None,
        max_rows=5,
        scope_hwnd=scope_hwnd,
    )
    learned = cache_body_table_suffix(
        dynamic_index,
        scope_hwnd,
        ((located or {}).get("best") or {}).get("path"),
    )
    return {
        **located,
        "source": "semantic-body-table-scan",
        "cached_path": (cached or {}).get("best"),
        "learned_suffix": learned,
    }

def cache_body_table_suffix(dynamic_index, scope_hwnd, path):
    prefix = receipt_header_dynamic_prefix(dynamic_index)
    if not prefix or not path or not str(path).startswith(f"{prefix}."):
        return None
    suffix = str(path)[len(prefix) + 1 :]
    cached = {
        "dynamic_index": dynamic_index,
        "scope_hwnd": scope_hwnd,
        "suffix": suffix,
        "path": path,
        "source": "semantic-body-table-scan",
    }
    _BODY_TABLE_SUFFIX_CACHE[body_table_cache_key(scope_hwnd)] = cached
    _BODY_TABLE_SUFFIX_CACHE["last"] = cached
    return cached

def body_table_cache_key(scope_hwnd):
    return int(scope_hwnd) if scope_hwnd is not None else None

def build_header_scope_for_followup(scope_hwnd, dynamic_index):
    if not scope_hwnd or dynamic_index is None:
        return None
    return {
        "ok": True,
        "scope_hwnd": scope_hwnd,
        "dynamic_index": dynamic_index,
        "mode": "provided-canvas-anchor",
    }

def extract_entry_scope_hwnd(report):
    state = (report or {}).get("entry_state") or {}
    hwnd = extract_entry_state_hwnd(state, prefer_canvas=True)
    if hwnd:
        return hwnd
    parsed = (report or {}).get("parsed") or {}
    for key in ("entry_state", "quick_entry_state"):
        hwnd = extract_entry_state_hwnd(parsed.get(key) or {}, prefer_canvas=True)
        if hwnd:
            return hwnd
    hwnd = extract_entry_state_hwnd(state, prefer_canvas=False)
    if hwnd:
        return hwnd
    for key in ("entry_state", "quick_entry_state"):
        hwnd = extract_entry_state_hwnd(parsed.get(key) or {}, prefer_canvas=False)
        if hwnd:
            return hwnd
    for key in (
        "windows_after_choose",
        "windows_after_open",
        "windows",
        "after_windows",
    ):
        for window in (report or {}).get(key) or parsed.get(key) or []:
            if (
                window.get("is_java")
                and window.get("visible")
                and window.get("hwnd")
                and window.get("class_name") == "SunAwtCanvas"
            ):
                return int(window["hwnd"])
    return None

def extract_entry_dynamic_index(report):
    state = (report or {}).get("entry_state") or {}
    dynamic_index = extract_dynamic_index_from_entry_state(state)
    if dynamic_index is not None:
        return dynamic_index
    parsed = (report or {}).get("parsed") or {}
    for key in ("entry_state", "quick_entry_state"):
        dynamic_index = extract_dynamic_index_from_entry_state(parsed.get(key) or {})
        if dynamic_index is not None:
            return dynamic_index
    return None

def extract_entry_anchor_path(report):
    state = (report or {}).get("entry_state") or {}
    path = extract_anchor_path_from_entry_state(state)
    if path:
        return path
    parsed = (report or {}).get("parsed") or {}
    for key in ("entry_state", "quick_entry_state"):
        path = extract_anchor_path_from_entry_state(parsed.get(key) or {})
        if path:
            return path
    return None

def extract_anchor_path_from_entry_state(state):
    for hit in (state or {}).get("hits") or []:
        control = hit.get("control") or {}
        if (
            control.get("name") == "财务组织(O)"
            or control.get("description") == "财务组织(O)"
        ):
            path = control.get("path")
            if path:
                return str(path)
    return None

def extract_dynamic_index_from_entry_state(state):
    for hit in (state or {}).get("hits") or []:
        control = hit.get("control") or {}
        direct = control.get("dynamic_index")
        if direct is not None:
            try:
                return int(direct)
            except (TypeError, ValueError):
                pass
        path = control.get("path") or ""
        dynamic_index = extract_receipt_module_dynamic_index(path)
        if dynamic_index is not None:
            return dynamic_index
    return None

def extract_receipt_module_dynamic_index(path):
    prefix = "0.0.1.0.0.0.0."
    text = str(path or "")
    if not text.startswith(prefix):
        return None
    part = text[len(prefix) :].split(".", 1)[0]
    try:
        return int(part)
    except ValueError:
        return None

def extract_entry_state_hwnd(state, prefer_canvas=False):
    if prefer_canvas:
        for hit in (state or {}).get("hits") or []:
            window = hit.get("window") or {}
            if window.get("class_name") == "SunAwtCanvas" and window.get("hwnd"):
                return int(window["hwnd"])
    for hit in (state or {}).get("hits") or []:
        window = hit.get("window") or {}
        hwnd = window.get("hwnd")
        if hwnd:
            return int(hwnd)
    return None
