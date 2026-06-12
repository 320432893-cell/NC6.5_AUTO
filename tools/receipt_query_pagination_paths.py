# 职责：收款查询分页区和结果表的 JAB path 解析、验证、动态枚举；不点击翻页、不改分页大小、不做匹配。
# 允许依赖层：tools.jab_probe、tools.receipt_query_result_tables 和 JAB operator-like 对象；core/报表/Excel 模块不应 import。

from copy import deepcopy
from typing import Any

from tools.receipt_query_result_tables import enumerate_receipt_result_table_paths


RECEIPT_RESULT_TABLE_PATH_SUFFIX = "0.0.0"
RECEIPT_PAGE_LABEL_PATH_SUFFIX = "1.6"
RECEIPT_PAGE_SIZE_TEXT_PATH_SUFFIX = "1.7"
RECEIPT_NEXT_PAGE_BUTTON_PATH_SUFFIX = "1.2"


def with_runtime_pagination_paths(query_cfg, path_report):
    runtime = deepcopy(query_cfg)
    pagination = deepcopy(runtime.get("pagination") or {})
    for key in ("page_label_path", "page_size_text_path", "next_page_button_path"):
        if path_report.get(key):
            pagination[key] = path_report[key]
    if path_report.get("window_class"):
        pagination["window_class"] = path_report["window_class"]
    runtime["pagination"] = pagination
    return runtime


def resolve_receipt_pagination_paths(jab, query_cfg):
    pagination = query_cfg.get("pagination") or {}
    window_class = pagination.get("window_class", "SunAwtCanvas")
    fallback_page_size_path = pagination.get("page_size_text_path")
    fallback_next_page_path = pagination.get("next_page_button_path")
    if fallback_page_size_path is None and pagination.get("page_label_path") == "label":
        fallback_page_size_path = "size"
    if fallback_next_page_path is None and pagination.get("page_label_path") == "label":
        fallback_next_page_path = "next"
    dynamic = None
    cached = getattr(jab, "_receipt_pagination_paths_cache", None)
    if cached:
        cached_report = validate_receipt_pagination_path_report(
            jab,
            pagination,
            cached,
            resolution="cached",
        )
        if cached_report.get("ok"):
            return cached_report
    if bool(pagination.get("prefer_configured_paths", True)):
        configured_report = {
            "window_class": window_class,
            "page_label_path": pagination["page_label_path"],
            "page_size_text_path": pagination["page_size_text_path"],
            "next_page_button_path": pagination["next_page_button_path"],
        }
        configured_report = validate_receipt_pagination_path_report(
            jab,
            pagination,
            configured_report,
            resolution="configured_fast",
        )
        if configured_report.get("ok"):
            setattr(jab, "_receipt_pagination_paths_cache", configured_report)
            return configured_report
    if bool(pagination.get("dynamic_paths_enabled", True)):
        dynamic = resolve_receipt_pagination_paths_dynamic(jab, query_cfg)
        if dynamic.get("ok"):
            setattr(jab, "_receipt_pagination_paths_cache", dynamic)
            return dynamic
    report = {
        "ok": True,
        "resolution": "configured",
        "window_class": window_class,
        "pager_hwnd": None,
        "result_table_path": None,
        "result_area_prefix": None,
        "page_label_path": pagination["page_label_path"],
        "page_size_text_path": fallback_page_size_path,
        "next_page_button_path": fallback_next_page_path,
    }
    if dynamic:
        report["dynamic_resolution"] = dynamic.get("resolution")
        report["result_page_probe"] = dynamic.get("resolution")
        report["dynamic_diagnostics"] = dynamic.get("diagnostics")
    return report


def validate_receipt_pagination_path_report(jab, pagination, report, resolution):
    if resolution == "configured_fast" and not bool(
        pagination.get("prefer_configured_paths", True)
    ):
        return {"ok": False, "resolution": f"{resolution}_disabled", **report}
    window_class = report.get("window_class") or pagination.get(
        "window_class", "SunAwtCanvas"
    )
    timeout = float(pagination.get("configured_path_timeout", 0.1))
    label = validate_context_path(
        jab,
        report["page_label_path"],
        window_class,
        role="label",
        scope_hwnd=report.get("pager_hwnd"),
        timeout=timeout,
    )
    if not label.get("ok"):
        return {"ok": False, "resolution": f"{resolution}_invalid", **report}
    page_size = validate_context_path(
        jab,
        report["page_size_text_path"],
        window_class,
        role="text",
        scope_hwnd=label.get("hwnd"),
        timeout=timeout,
    )
    next_page = validate_context_path(
        jab,
        report["next_page_button_path"],
        window_class,
        role="push button",
        scope_hwnd=label.get("hwnd"),
        timeout=timeout,
    )
    prefix = report.get(
        "result_area_prefix"
    ) or infer_result_area_prefix_from_page_path(report["page_label_path"])
    result_table_path = report.get("result_table_path")
    result_table = {"ok": None}
    if prefix and not result_table_path:
        result_table_path = join_context_path(prefix, RECEIPT_RESULT_TABLE_PATH_SUFFIX)
    if result_table_path:
        result_table = validate_context_path(
            jab,
            result_table_path,
            window_class,
            role="table",
            scope_hwnd=label.get("hwnd"),
            timeout=timeout,
        )
    ok = bool(page_size.get("ok") and next_page.get("ok"))
    return {
        "ok": ok,
        "resolution": resolution if ok else f"{resolution}_invalid",
        "window_class": window_class,
        "pager_hwnd": label.get("hwnd"),
        "result_table_path": result_table_path if result_table.get("ok") else None,
        "result_area_prefix": prefix if result_table.get("ok") else None,
        "page_label_path": report["page_label_path"],
        "page_size_text_path": report["page_size_text_path"],
        "next_page_button_path": report["next_page_button_path"],
        "diagnostics": {
            "label": label,
            "page_size": page_size,
            "next_page": next_page,
            "result_table": result_table,
        },
    }


def resolve_receipt_pagination_paths_dynamic(jab, query_cfg):
    pagination = query_cfg.get("pagination") or {}
    window_class = pagination.get("window_class", "SunAwtCanvas")
    candidates = enumerate_receipt_result_table_paths(jab, query_cfg, window_class)
    diagnostics: dict[str, Any] = {"candidates": candidates[:10]}
    for candidate in candidates:
        prefix = infer_result_area_prefix_from_table_path(candidate["path"])
        if not prefix:
            continue
        paths = {
            "page_label_path": join_context_path(
                prefix, RECEIPT_PAGE_LABEL_PATH_SUFFIX
            ),
            "page_size_text_path": join_context_path(
                prefix, RECEIPT_PAGE_SIZE_TEXT_PATH_SUFFIX
            ),
            "next_page_button_path": join_context_path(
                prefix, RECEIPT_NEXT_PAGE_BUTTON_PATH_SUFFIX
            ),
        }
        label = validate_context_path(
            jab,
            paths["page_label_path"],
            window_class,
            role="label",
            scope_hwnd=candidate["hwnd"],
            timeout=float(pagination.get("dynamic_path_timeout", 0.2)),
        )
        page_size = validate_context_path(
            jab,
            paths["page_size_text_path"],
            window_class,
            role="text",
            scope_hwnd=candidate["hwnd"],
            timeout=float(pagination.get("dynamic_path_timeout", 0.2)),
        )
        next_page = validate_context_path(
            jab,
            paths["next_page_button_path"],
            window_class,
            role="push button",
            scope_hwnd=candidate["hwnd"],
            timeout=float(pagination.get("dynamic_path_timeout", 0.2)),
        )
        diagnostics["last_candidate"] = {
            "table": candidate,
            "prefix": prefix,
            "label": label,
            "page_size": page_size,
            "next_page": next_page,
        }
        if label.get("ok") and page_size.get("ok") and next_page.get("ok"):
            return {
                "ok": True,
                "resolution": "dynamic",
                "window_class": window_class,
                "pager_hwnd": candidate["hwnd"],
                "result_table_path": candidate["path"],
                "result_area_prefix": prefix,
                "diagnostics": diagnostics,
                **paths,
            }
    return {
        "ok": False,
        "resolution": "result_page_pager_not_found",
        "window_class": window_class,
        "pager_hwnd": None,
        "result_table_path": None,
        "result_area_prefix": None,
        "diagnostics": diagnostics,
        "page_label_path": pagination["page_label_path"],
        "page_size_text_path": pagination.get("page_size_text_path"),
        "next_page_button_path": pagination.get("next_page_button_path"),
    }


def infer_result_area_prefix_from_table_path(table_path):
    return strip_context_path_suffix(table_path, RECEIPT_RESULT_TABLE_PATH_SUFFIX)


def infer_result_area_prefix_from_page_path(page_path):
    for suffix in (
        RECEIPT_PAGE_LABEL_PATH_SUFFIX,
        RECEIPT_PAGE_SIZE_TEXT_PATH_SUFFIX,
        RECEIPT_NEXT_PAGE_BUTTON_PATH_SUFFIX,
    ):
        prefix = strip_context_path_suffix(page_path, suffix)
        if prefix:
            return prefix
    return None


def strip_context_path_suffix(path, suffix):
    path_parts = split_context_path(path)
    suffix_parts = split_context_path(suffix)
    if len(path_parts) < len(suffix_parts):
        return None
    if path_parts[-len(suffix_parts) :] != suffix_parts:
        return None
    prefix_parts = path_parts[: -len(suffix_parts)]
    if not prefix_parts:
        return None
    return ".".join(str(part) for part in prefix_parts)


def join_context_path(prefix, suffix):
    prefix_text = str(prefix).strip(".")
    suffix_text = str(suffix).strip(".")
    if not prefix_text:
        return suffix_text
    if not suffix_text:
        return prefix_text
    return f"{prefix_text}.{suffix_text}"


def split_context_path(path):
    try:
        return [int(part) for part in str(path).split(".") if part != ""]
    except ValueError:
        return []


def validate_context_path(jab, path, window_class, role, scope_hwnd=None, timeout=0.2):
    window = jab.wait_context_by_path(
        path,
        class_name=window_class,
        role=role,
        timeout=timeout,
        scope_hwnd=scope_hwnd,
        require_showing=True,
        require_valid_bounds=False,
    )
    if not window:
        return {"ok": False, "path": path, "role": role}
    return {
        "ok": True,
        "path": path,
        "role": role,
        "hwnd": window.get("hwnd"),
        "class": window.get("class"),
        "title": window.get("title"),
    }
