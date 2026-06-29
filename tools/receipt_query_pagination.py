# 职责：收款查询结果分页运行时动作，包括等待结果、设置页大小、读取页码和点击下一页。
# 不做什么：不解析动态 path 树、不读取多页结果集合、不做 Excel 匹配、不生成报告。
# 允许依赖层：tools.receipt_query_pagination_paths、tools.receipt_query_result_tables 和 JAB operator-like 对象。
# 谁不应该 import：core 层模块不应 import；纯报表/匹配模块不应 import。

import re
import time

from tools.receipt_keyboard_utils import send_virtual_key
from tools.receipt_query_pagination_paths import (
    resolve_receipt_pagination_paths,
    resolve_receipt_pagination_paths_dynamic,
    with_runtime_pagination_paths,
)
from tools.receipt_query_result_tables import summarize_receipt_tables


def wait_after_query_confirm(jab, query_cfg):
    pagination = query_cfg.get("pagination") or {}
    wait_timeout = float(query_cfg.get("result_wait_timeout", 0.5))
    if not pagination or wait_timeout <= 0:
        return {
            "ok": None,
            "method": "disabled",
            "label": None,
            "seconds": 0.0,
        }

    started = time.perf_counter()
    interval = float(query_cfg.get("result_wait_interval", 0.1))
    window_class = pagination.get("window_class", "SunAwtCanvas")
    page_label_path = pagination.get("page_label_path")
    fallback_wait = float(query_cfg.get("result_wait_fallback", 0.0))
    while time.perf_counter() - started < wait_timeout:
        probe_cfg = {
            **query_cfg,
            "pagination": {
                **pagination,
                "module_index_table_ready_only": True,
            },
        }
        dynamic = resolve_receipt_pagination_paths_dynamic(jab, probe_cfg)
        if dynamic.get("ok"):
            setattr(jab, "_receipt_pagination_paths_cache", dynamic)
            return {
                "ok": True,
                "method": "result_table_path",
                "label": None,
                "result_table_path": dynamic.get("result_table_path"),
                "result_area_prefix": dynamic.get("result_area_prefix"),
                "seconds": round(time.perf_counter() - started, 3),
            }
        if page_label_path:
            label = read_page_label(jab, page_label_path, window_class)
            if label:
                return {
                    "ok": True,
                    "method": "page_label",
                    "label": label,
                    "seconds": round(time.perf_counter() - started, 3),
                }
        time.sleep(interval)

    if fallback_wait > 0:
        time.sleep(fallback_wait)
    return {
        "ok": False,
        "method": "timeout",
        "label": None,
        "seconds": round(time.perf_counter() - started + fallback_wait, 3),
    }


def set_receipt_page_size(jab, query_cfg):
    pagination = query_cfg.get("pagination") or {}
    if not pagination:
        return {"enabled": False, "page_size_ok": False}

    page_size = int(pagination.get("page_size", 500))
    resolved_paths = resolve_receipt_pagination_paths(jab, query_cfg)
    window_class = resolved_paths["window_class"]
    page_size_path = resolved_paths["page_size_text_path"]
    page_label_path = resolved_paths["page_label_path"]
    next_page_path = resolved_paths["next_page_button_path"]
    pager_resolution = resolved_paths["resolution"]
    result_page_resolution = resolved_paths["resolution"]
    dynamic_resolution = resolved_paths.get("dynamic_resolution")
    dynamic_diagnostics = resolved_paths.get("dynamic_diagnostics")
    runtime_query_cfg = with_runtime_pagination_paths(query_cfg, resolved_paths)
    pager_hwnd = resolved_paths.get("pager_hwnd")
    if not pager_hwnd or not bool(pagination.get("trust_cached_paths", True)):
        pager_window = jab.wait_context_by_path(
            page_label_path,
            class_name=window_class,
            role="label",
            timeout=float(pagination.get("pager_scope_timeout", 2.0)),
            require_showing=False,
            require_valid_bounds=False,
        )
        pager_hwnd = (
            int(pager_window.get("hwnd"))
            if pager_window and pager_window.get("hwnd") is not None
            else pager_hwnd
        )
    if not pager_hwnd:
        return {
            "enabled": True,
            "page_size": page_size,
            "page_size_ok": False,
            "page_size_changed": None,
            "before_page_size_text": None,
            "after_page_size_text": None,
            "pager_hwnd": None,
            "pager_scope_ok": False,
            "window_class": window_class,
            "page_label_path": page_label_path,
            "page_size_text_path": page_size_path,
            "next_page_button_path": next_page_path,
            "pager_resolution": pager_resolution,
            "result_page_resolution": result_page_resolution,
            "dynamic_resolution": dynamic_resolution,
            "dynamic_diagnostics": dynamic_diagnostics,
            "before_label": None,
            "after_label": None,
            "before_stability": {"ok": None, "reason": "pager_scope_not_found"},
            "after_stability": {"ok": None, "reason": "pager_scope_not_found"},
            "before_stability_seconds": 0.0,
            "wait_before_page_size_seconds": 0.0,
            "set_page_size_text_seconds": 0.0,
            "page_size_enter_seconds": 0.0,
            "after_stability_seconds": 0.0,
        }

    if bool(pagination.get("wait_before_page_size_stable", True)):
        before_stability_start = time.perf_counter()
        before_stability = wait_receipt_result_stable(
            jab, runtime_query_cfg, pager_hwnd=pager_hwnd
        )
        before_stability_seconds = time.perf_counter() - before_stability_start
    else:
        before_stability = {
            "ok": None,
            "skipped": True,
            "reason": "pre_stability_disabled",
            "label": None,
            "tables": [],
        }
        before_stability_seconds = 0.0
    before_page_size_text = read_page_size_text(
        jab, page_size_path, window_class, pager_hwnd
    )
    current_page_size = parse_int_text(before_page_size_text)
    page_size_changed = current_page_size != page_size
    if page_size_changed:
        before_label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
    else:
        before_label = getattr(jab, "_receipt_last_page_label", None)
        if before_label is None and bool(
            pagination.get("read_label_when_page_size_ok", True)
        ):
            before_label = read_page_label(
                jab, page_label_path, window_class, pager_hwnd
            )
    wait_before_page_size_seconds = 0.0
    set_page_size_text_seconds = 0.0
    enter_seconds = 0.0
    page_size_ok = True
    if page_size_changed:
        wait_before_page_size_start = time.perf_counter()
        time.sleep(float(pagination.get("wait_before_page_size", 0.0)))
        wait_before_page_size_seconds = (
            time.perf_counter() - wait_before_page_size_start
        )
        set_text_start = time.perf_counter()
        page_size_ok = jab.set_text_by_path(
            page_size_path,
            str(page_size),
            class_name=window_class,
            scope_hwnd=pager_hwnd,
            role="text",
            timeout=2,
            require_showing=False,
            require_valid_bounds=False,
        )
        set_page_size_text_seconds = time.perf_counter() - set_text_start
    if page_size_ok and page_size_changed:
        enter_start = time.perf_counter()
        press_enter_for_page_size(
            jab, wait=float(pagination.get("wait_after_page_size", 2.0))
        )
        enter_seconds = time.perf_counter() - enter_start
    if page_size_changed or bool(
        pagination.get("ready_check_when_page_size_ok", False)
    ):
        after_stability_start = time.perf_counter()
        after_stability = wait_receipt_result_ready(
            jab, runtime_query_cfg, pager_hwnd=pager_hwnd
        )
        after_stability_seconds = time.perf_counter() - after_stability_start
        after_label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
        after_page_size_text = read_page_size_text(
            jab, page_size_path, window_class, pager_hwnd
        )
    else:
        after_stability = {
            "ok": None,
            "skipped": True,
            "reason": "page_size_already_target",
            "label": before_label,
            "tables": [],
        }
        after_stability_seconds = 0.0
        after_label = before_label
        after_page_size_text = before_page_size_text
    if after_label:
        setattr(jab, "_receipt_last_page_label", after_label)
    return {
        "enabled": True,
        "page_size": page_size,
        "page_size_ok": bool(page_size_ok),
        "page_size_changed": bool(page_size_changed),
        "before_page_size_text": before_page_size_text,
        "after_page_size_text": after_page_size_text,
        "pager_hwnd": pager_hwnd,
        "pager_scope_ok": bool(pager_hwnd),
        "window_class": window_class,
        "result_table_path": resolved_paths.get("result_table_path"),
        "result_area_prefix": resolved_paths.get("result_area_prefix"),
        "page_label_path": page_label_path,
        "page_size_text_path": page_size_path,
        "next_page_button_path": next_page_path,
        "pager_resolution": pager_resolution,
        "result_page_resolution": result_page_resolution,
        "dynamic_resolution": dynamic_resolution,
        "dynamic_diagnostics": dynamic_diagnostics,
        "before_label": before_label,
        "after_label": after_label,
        "before_stability": before_stability,
        "after_stability": after_stability,
        "before_stability_seconds": round(before_stability_seconds, 3),
        "wait_before_page_size_seconds": round(wait_before_page_size_seconds, 3),
        "set_page_size_text_seconds": round(set_page_size_text_seconds, 3),
        "page_size_enter_seconds": round(enter_seconds, 3),
        "after_stability_seconds": round(after_stability_seconds, 3),
    }


def wait_receipt_result_stable(jab, query_cfg, pager_hwnd=None):
    pagination = query_cfg.get("pagination") or {}
    if not pagination:
        return {"enabled": False}

    window_class = pagination.get("window_class", "SunAwtCanvas")
    page_label_path = pagination["page_label_path"]
    timeout = float(pagination.get("stability_timeout", 12.0))
    interval = float(pagination.get("stability_interval", 1.0))
    required = int(pagination.get("stability_required", 2))
    deadline = time.time() + timeout
    previous = None
    stable_count = 0
    samples = []
    started = time.perf_counter()

    while time.time() < deadline:
        label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
        summary = summarize_receipt_tables(jab, query_cfg, scope_hwnd=pager_hwnd)
        sample = {"label": label, "tables": summary}
        samples.append(sample)
        if label and summary and sample == previous:
            stable_count += 1
        else:
            stable_count = 1
        previous = sample
        if stable_count >= required:
            return {
                "ok": True,
                "samples": len(samples),
                "label": label,
                "tables": summary,
                "seconds": round(time.perf_counter() - started, 3),
            }
        time.sleep(interval)

    last = samples[-1] if samples else {"label": None, "tables": []}
    return {
        "ok": False,
        "samples": len(samples),
        "label": last.get("label"),
        "tables": last.get("tables"),
        "seconds": round(time.perf_counter() - started, 3),
    }


def wait_receipt_result_ready(jab, query_cfg, pager_hwnd=None):
    pagination = query_cfg.get("pagination") or {}
    if not pagination:
        return {"enabled": False}
    if bool(pagination.get("result_ready_fast", True)):
        started = time.perf_counter()
        window_class = pagination.get("window_class", "SunAwtCanvas")
        label = read_page_label(
            jab, pagination["page_label_path"], window_class, pager_hwnd
        )
        tables = summarize_receipt_tables(jab, query_cfg, scope_hwnd=pager_hwnd)
        if label and tables:
            return {
                "ok": True,
                "fast": True,
                "samples": 1,
                "label": label,
                "tables": tables,
                "seconds": round(time.perf_counter() - started, 3),
            }
    return wait_receipt_result_stable(jab, query_cfg, pager_hwnd=pager_hwnd)


def click_next_page(jab, pagination, next_page_path, window_class, scope_hwnd=None):
    wait_after_next = float(pagination.get("wait_after_next", 2.0))
    action_timeout = float(pagination.get("next_action_timeout", 2.0))
    ok = jab.do_action_by_path(
        next_page_path,
        class_name=window_class,
        scope_hwnd=scope_hwnd,
        role="push button",
        action_name="单击",
        timeout=action_timeout,
        wait=wait_after_next,
        require_showing=True,
        require_valid_bounds=False,
    )
    if ok:
        return True, "action"
    return False, "failed"


def press_enter_for_page_size(jab, wait=2.0):
    try:
        jab.press_key("enter", wait=wait)
    except ModuleNotFoundError:
        send_virtual_key(0x0D)
        time.sleep(float(wait or 0))


def refresh_receipt_results(jab, query_cfg, runtime_query_cfg=None, pager_hwnd=None):
    effective_query_cfg = runtime_query_cfg or query_cfg
    pagination = effective_query_cfg.get("pagination") or {}
    if not pagination or not bool(pagination.get("refresh_after_page_size", True)):
        return {"enabled": False, "reason": "refresh_after_page_size_disabled"}
    wait_after_refresh = float(pagination.get("wait_after_refresh", 0.5))
    window_class = pagination.get("window_class", "SunAwtCanvas")
    page_label_path = pagination.get("page_label_path")
    page_size_path = pagination.get("page_size_text_path")
    started = time.perf_counter()
    try:
        jab.press_key("f5", wait=wait_after_refresh)
        method = "jab.press_key"
    except ModuleNotFoundError:
        send_virtual_key(0x74)
        time.sleep(wait_after_refresh)
        method = "virtual_key"
    stability = wait_receipt_result_stable(
        jab,
        runtime_query_cfg or query_cfg,
        pager_hwnd=pager_hwnd,
    )
    label = (
        read_page_label(jab, page_label_path, window_class, pager_hwnd)
        if page_label_path
        else None
    )
    page_size_text = (
        read_page_size_text(jab, page_size_path, window_class, pager_hwnd)
        if page_size_path
        else None
    )
    return {
        "enabled": True,
        "ok": True,
        "method": method,
        "key": "f5",
        "wait_after_refresh": wait_after_refresh,
        "label": label,
        "page_size_text": page_size_text,
        "stability": stability,
        "seconds": round(time.perf_counter() - started, 3),
    }


def read_page_label(jab, path, window_class, scope_hwnd=None):
    return jab.get_text_by_path(
        path,
        class_name=window_class,
        scope_hwnd=scope_hwnd,
        role="label",
        timeout=1,
        require_showing=False,
        require_valid_bounds=False,
    )


def read_page_size_text(jab, path, window_class, scope_hwnd=None):
    return jab.get_text_by_path(
        path,
        class_name=window_class,
        scope_hwnd=scope_hwnd,
        role="text",
        timeout=1,
        require_showing=False,
        require_valid_bounds=False,
    )


def parse_int_text(value):
    match = re.search(r"\d+", str(value or ""))
    return int(match.group(0)) if match else None


def parse_page_label(value):
    text = str(value or "")
    total_pages = None
    total_records = None
    page_match = re.search(r"共\s*(\d+)\s*页", text)
    record_match = re.search(r"(\d+)\s*条记录", text)
    if page_match:
        total_pages = int(page_match.group(1))
    if record_match:
        total_records = int(record_match.group(1))
    return {"total_pages": total_pages, "total_records": total_records}
