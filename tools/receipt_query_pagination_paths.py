# 职责：收款查询分页区和结果表的 JAB path 解析、验证、动态枚举；不点击翻页、不改分页大小、不做匹配。
# 允许依赖层：tools.jab_probe、tools.receipt_query_result_tables 和 JAB operator-like 对象；core/报表/Excel 模块不应 import。

from copy import deepcopy
from typing import Any

from tools.receipt_query_result_tables import enumerate_visible_table_paths


RECEIPT_RESULT_TABLE_PATH_SUFFIX = "0.0.0"
RECEIPT_PAGE_LABEL_PATH_SUFFIX = "1.6"
RECEIPT_PAGE_SIZE_TEXT_PATH_SUFFIX = "1.7"
RECEIPT_NEXT_PAGE_BUTTON_PATH_SUFFIX = "1.2"
RECEIPT_MODULE_PREFIX_BASE = "0.0.1.0.0.0.0"
RECEIPT_MODULE_DYNAMIC_MAX_INDEX = 8
RECEIPT_RESULT_AREA_PATH_SUFFIX = "0.0.0.1.1.0.0.0.1.1.1.0.0.0"


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
        if bool(pagination.get("trust_cached_paths", True)):
            return {
                "ok": True,
                "resolution": "cached_trusted",
                "window_class": cached.get("window_class", window_class),
                "pager_hwnd": cached.get("pager_hwnd"),
                "result_table_path": cached.get("result_table_path"),
                "result_area_prefix": cached.get("result_area_prefix"),
                "page_label_path": cached.get("page_label_path"),
                "page_size_text_path": cached.get("page_size_text_path"),
                "next_page_button_path": cached.get("next_page_button_path"),
                "dynamic_resolution": cached.get("resolution"),
                "dynamic_diagnostics": cached.get("diagnostics"),
            }
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


def validate_receipt_pagination_path_report(
    jab, pagination, report, resolution, timeout=None
):
    if resolution == "configured_fast" and not bool(
        pagination.get("prefer_configured_paths", True)
    ):
        return {"ok": False, "resolution": f"{resolution}_disabled", **report}
    window_class = report.get("window_class") or pagination.get(
        "window_class", "SunAwtCanvas"
    )
    timeout = (
        float(timeout)
        if timeout is not None
        else float(pagination.get("configured_path_timeout", 0.1))
    )
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
    if resolution == "dynamic_module_index" and not result_table.get("ok"):
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
    module_index_report = resolve_receipt_pagination_paths_by_module_index(
        jab, query_cfg
    )
    if module_index_report.get("ok"):
        return module_index_report
    candidates = enumerate_visible_table_paths(jab, window_class)
    diagnostics: dict[str, Any] = {
        "module_index_attempt": module_index_report,
        "candidates": candidates[:10],
    }
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


def resolve_receipt_pagination_paths_by_module_index(jab, query_cfg):
    pagination = query_cfg.get("pagination") or {}
    if not bool(pagination.get("module_index_paths_enabled", False)):
        return {
            "ok": False,
            "resolution": "module_index_disabled",
            "window_class": pagination.get("window_class", "SunAwtCanvas"),
            "attempts": [],
        }
    window_class = pagination.get("window_class", "SunAwtCanvas")
    attempts = []
    max_index = int(
        pagination.get("module_dynamic_max_index", RECEIPT_MODULE_DYNAMIC_MAX_INDEX)
    )
    preferred_index = getattr(jab, "_receipt_module_dynamic_index_cache", None)
    if preferred_index is None:
        configured_prefix = str(pagination.get("page_label_path") or "")
        configured_parts = split_context_path(configured_prefix)
        base_parts = split_context_path(RECEIPT_MODULE_PREFIX_BASE)
        if len(configured_parts) > len(base_parts):
            if configured_parts[: len(base_parts)] == base_parts:
                preferred_index = configured_parts[len(base_parts)]
    ordered_indexes = []
    if isinstance(preferred_index, int) and 0 <= preferred_index <= max_index:
        ordered_indexes.append(preferred_index)
    ordered_indexes.extend(
        index for index in range(max_index + 1) if index not in ordered_indexes
    )
    for dynamic_index in ordered_indexes:
        module_prefix = f"{RECEIPT_MODULE_PREFIX_BASE}.{dynamic_index}"
        prefix = join_context_path(module_prefix, RECEIPT_RESULT_AREA_PATH_SUFFIX)
        paths = {
            "result_table_path": join_context_path(
                prefix, RECEIPT_RESULT_TABLE_PATH_SUFFIX
            ),
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
        report = {
            "window_class": window_class,
            "pager_hwnd": None,
            "dynamic_index": dynamic_index,
            "module_prefix": module_prefix,
            "result_area_prefix": prefix,
            **paths,
        }
        table_ready = validate_context_path(
            jab,
            paths["result_table_path"],
            window_class,
            role="table",
            timeout=float(pagination.get("dynamic_path_timeout", 0.1)),
        )
        if table_ready.get("ok") and bool(
            pagination.get("module_index_table_ready_only", False)
        ):
            setattr(jab, "_receipt_module_dynamic_index_cache", dynamic_index)
            checked = {
                "ok": True,
                "resolution": "dynamic_module_index_table_ready",
                "window_class": window_class,
                "pager_hwnd": table_ready.get("hwnd"),
                "result_table_path": paths["result_table_path"],
                "result_area_prefix": prefix,
                "page_label_path": paths["page_label_path"],
                "page_size_text_path": paths["page_size_text_path"],
                "next_page_button_path": paths["next_page_button_path"],
                "diagnostics": {"result_table": table_ready},
            }
            checked["dynamic_index"] = dynamic_index
            checked["module_prefix"] = module_prefix
            return checked
        checked = validate_receipt_pagination_path_report(
            jab,
            pagination,
            report,
            resolution="dynamic_module_index",
            timeout=float(pagination.get("dynamic_path_timeout", 0.1)),
        )
        attempts.append(
            {
                "dynamic_index": dynamic_index,
                "module_prefix": module_prefix,
                "result_area_prefix": prefix,
                "ok": checked.get("ok"),
                "diagnostics": checked.get("diagnostics"),
            }
        )
        if checked.get("ok"):
            setattr(jab, "_receipt_module_dynamic_index_cache", dynamic_index)
            checked["dynamic_index"] = dynamic_index
            checked["module_prefix"] = module_prefix
            checked["diagnostics"] = {
                **(checked.get("diagnostics") or {}),
                "attempts": attempts,
            }
            return checked
    return {
        "ok": False,
        "resolution": "module_index_result_page_not_found",
        "window_class": window_class,
        "attempts": attempts,
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
        require_showing=False,
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
