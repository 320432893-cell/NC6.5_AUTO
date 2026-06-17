# 职责：从开单探测报告萃取入口上下文(scope_hwnd/动态前缀/锚点 path)并定位明细表
# 不做什么：不写入字段，不做报告渲染，不触发保存，不做行编排
# 允许依赖层：tools.receipt_body_table_locator、tools.receipt_self_made_fill_trial
# 谁不应该 import：core 层模块不应 import；本模块不应反向 import row_runner

import sys
import time

from tools.receipt_self_made_fill_trial import receipt_header_dynamic_prefix

BODY_TABLE_SUFFIX = "0.0.0.1.1.0.0.0.0.1.0.2.1.0.0.0.0.0"


class _FlowNamespace:
    # 按调用时从已加载的入口模块取属性：让测试对
    # tools.receipt_full_flow_entry.<name> 的 monkeypatch 与拆分前一致地生效，
    # 同时不在加载期 import 入口模块,彻底避免相互 import 成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_full_flow_entry"], name)


_flow = _FlowNamespace()


def run_with_jab_lock(jab_lock, func, *args, **kwargs):
    if jab_lock is None:
        return func(*args, **kwargs)
    with jab_lock:
        return func(*args, **kwargs)


def wait_receipt_header_anchor_in_current_canvas(
    jab,
    scope_hwnd,
    timeout=1.2,
    interval=0.2,
):
    started_at = time.perf_counter()
    deadline = started_at + max(float(timeout or 0), 0.0)
    interval = max(float(interval or 0.2), 0.01)
    attempts = []
    while True:
        remaining = max(deadline - time.perf_counter(), 0.0)
        attempt = _flow.resolve_receipt_header_anchor_in_canvas(
            jab,
            scope_hwnd,
            timeout=min(0.05, remaining) if remaining > 0 else 0.05,
        )
        attempts.append(attempt)
        if attempt.get("ok"):
            return {
                **attempt,
                "attempts": attempts,
                "poll_interval": interval,
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        if time.perf_counter() >= deadline:
            return {
                "ok": False,
                "reason": attempt.get("reason") or "当前 canvas 未找到财务组织(O) 锚点",
                "scope_hwnd": scope_hwnd,
                "attempts": attempts,
                "poll_interval": interval,
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        time.sleep(min(interval, max(deadline - time.perf_counter(), 0.0)))


def build_body_table_cached_path(dynamic_index, scope_hwnd=None):
    if dynamic_index is None:
        return None
    path = f"{receipt_header_dynamic_prefix(dynamic_index)}.{BODY_TABLE_SUFFIX}"
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
    located = _flow.locate_receipt_body_table_cached(
        jab,
        cached=cached,
        max_rows=5,
        scope_hwnd=scope_hwnd,
    )
    source = (
        "dynamic-prefix-body-path"
        if located.get("cache_hit")
        else "dynamic-prefix-body-path-fallback-scan"
    )
    return {**located, "source": source, "cached_path": (cached or {}).get("best")}


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
