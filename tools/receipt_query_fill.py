import argparse
from collections import Counter
from copy import deepcopy
import ctypes
from datetime import date
import json
from pathlib import Path
import re
import os
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_entry import (  # noqa: E402
    ReceiptEntryDryRunMatcher,
    ReceiptEntryWorkbook,
    ReceiptNCResultExtractor,
    format_receipt_amount_name_mismatch_reason,
    format_receipt_duplicate_reason,
    format_receipt_name_amount_mismatch_reason,
    format_receipt_not_found_reason,
    names_match,
    parse_date,
)
from core.utils import load_config  # noqa: E402
from tools.jab_probe import JOBJECT  # noqa: E402


class ReceiptPageGuardError(RuntimeError):
    pass


RECEIPT_RESULT_TABLE_PATH_SUFFIX = "0.0.0"
RECEIPT_PAGE_LABEL_PATH_SUFFIX = "1.6"
RECEIPT_PAGE_SIZE_TEXT_PATH_SUFFIX = "1.7"
RECEIPT_NEXT_PAGE_BUTTON_PATH_SUFFIX = "1.2"


class TimingRecorder:
    def __init__(self):
        self.items = []

    def measure(self, timing_name, func, *args, **kwargs):
        start = time.perf_counter()
        result = func(*args, **kwargs)
        self.add(timing_name, time.perf_counter() - start)
        return result

    def add(self, name, seconds):
        self.items.append({"name": name, "seconds": round(float(seconds), 3)})


def resolve_today(value):
    return date.today().isoformat() if value == "{today}" else value


def set_text(jab, jab_cfg, path, value):
    return jab.set_text_by_path(
        path,
        value,
        title=jab_cfg["dialog_title"],
        class_name=jab_cfg["dialog_class"],
        role="text",
        wait=float(jab_cfg.get("text_set_wait", 0.0)),
        timeout=2,
        require_showing=True,
    )


def set_finance_org_text(jab, jab_cfg, field_cfg, value):
    text_path = field_cfg.get("text_path")
    if text_path and jab.set_text_by_path(
        text_path,
        value,
        title=jab_cfg["dialog_title"],
        class_name=jab_cfg["dialog_class"],
        role="text",
        wait=float(field_cfg.get("path_wait", jab_cfg.get("text_set_wait", 0.0))),
        timeout=float(field_cfg.get("path_timeout", 0.5)),
        require_showing=True,
    ):
        return True

    return jab.set_text_near_label(
        field_cfg["label"],
        value,
        title=jab_cfg["dialog_title"],
        class_name=jab_cfg["dialog_class"],
        timeout=float(field_cfg.get("timeout", 2.0)),
        require_showing=True,
    )


def ensure_query_window(jab, config, query_cfg, jab_cfg, skip_open=False):
    title = jab_cfg["dialog_title"]
    class_name = jab_cfg["dialog_class"]
    timeout = float(query_cfg.get("open_timeout", query_cfg.get("timeout", 5)))
    existing_timeout = float(query_cfg.get("existing_dialog_timeout", 0.1))
    existing = jab.wait_window_by_title(
        title,
        class_name=class_name,
        timeout=existing_timeout,
        include_children=bool(query_cfg.get("dialog_include_children", True)),
        visible_only=bool(query_cfg.get("dialog_visible_only", True)),
    )
    if existing or skip_open:
        return bool(existing)

    batch_open_query = (config.get("jab_batch") or {}).get("open_query") or {}
    main_title = query_cfg.get("main_title", batch_open_query.get("main_title", ""))
    main_class = query_cfg.get("main_class", batch_open_query.get("main_class"))
    if main_title:
        jab.activate_window_by_title(
            main_title,
            class_name=main_class,
            timeout=float(query_cfg.get("activate_timeout", 5)),
        )
    jab.press_key(
        query_cfg.get("open_key", batch_open_query.get("key", "f3")),
        wait=float(query_cfg.get("open_wait", 0.0)),
    )
    return bool(
        jab.wait_window_by_title(
            title,
            class_name=class_name,
            timeout=timeout,
            include_children=bool(query_cfg.get("dialog_include_children", True)),
            visible_only=bool(query_cfg.get("dialog_visible_only", True)),
        )
    )


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
    interval = float(query_cfg.get("result_wait_interval", 0.05))
    window_class = pagination.get("window_class", "SunAwtCanvas")
    page_label_path = pagination.get("page_label_path")
    fallback_wait = float(query_cfg.get("result_wait_fallback", 0.0))
    while time.perf_counter() - started < wait_timeout:
        dynamic = resolve_receipt_pagination_paths_dynamic(jab, query_cfg)
        if dynamic.get("ok"):
            setattr(jab, "_receipt_pagination_paths_cache", dynamic)
            label = read_page_label(
                jab,
                dynamic["page_label_path"],
                window_class,
                dynamic.get("pager_hwnd"),
            )
            return {
                "ok": True,
                "method": "dynamic-result-table",
                "label": label,
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


def fill_receipt_query(
    config,
    org_code,
    date_from=None,
    date_to=None,
    confirm=False,
    read_results=False,
    dry_run_match=False,
    skip_open_query=False,
    max_rows=500,
    max_cols=80,
    set_page_size_only=False,
    write_back=False,
):
    query_cfg = config["receipt_entry"]["query"]
    jab_cfg = query_cfg["jab"]
    fields = jab_cfg["fields"]
    start = date_from or query_cfg["date_from"]
    end = date_to or query_cfg["date_to"]
    start = parse_date(resolve_today(start)).isoformat()
    end = parse_date(resolve_today(end)).isoformat()
    timings = TimingRecorder()
    total_start = time.perf_counter()

    jab = JABOperator(config)
    try:
        timings.measure("page_guard", guard_receipt_parent_page, jab, config, query_cfg)
        query_window_ok = timings.measure(
            "ensure_query_window",
            ensure_query_window,
            jab,
            config,
            query_cfg,
            jab_cfg,
            skip_open=skip_open_query,
        )
        if not query_window_ok:
            raise RuntimeError("未检测到收款单查询条件窗口")

        finance_org_ok = timings.measure(
            "set_finance_org",
            set_finance_org_text,
            jab,
            jab_cfg,
            fields["finance_org"],
            org_code,
        )
        if not finance_org_ok:
            raise RuntimeError(f"收款查询条件写入失败: finance_org={org_code}")

        steps = [
            ("document_date_from", fields["document_date"]["from_text_path"], start),
            ("document_date_to", fields["document_date"]["to_text_path"], end),
        ]
        for name, path, value in steps:
            if not timings.measure(name, set_text, jab, jab_cfg, path, value):
                raise RuntimeError(f"收款查询条件写入失败: {name}={value}")

        if confirm:
            ok = timings.measure(
                "confirm_query",
                jab.do_action_by_path,
                jab_cfg["confirm_button_path"],
                title=jab_cfg["dialog_title"],
                class_name=jab_cfg["dialog_class"],
                role="push button",
                action_name="单击",
                wait=float(jab_cfg.get("confirm_wait", 0.0)),
                timeout=float(jab_cfg.get("confirm_timeout", 1.0)),
                require_showing=True,
            )
            if not ok:
                raise RuntimeError("收款查询确定按钮点击失败")

        result = {"organization_code": org_code, "date_from": start, "date_to": end}
        if set_page_size_only:
            if confirm:
                wait_start = time.perf_counter()
                result["result_wait_before_page_size"] = wait_after_query_confirm(
                    jab, query_cfg
                )
                timings.add(
                    "result_wait_before_page_size", time.perf_counter() - wait_start
                )
            result["page_report"] = timings.measure(
                "set_receipt_page_size", set_receipt_page_size, jab, query_cfg
            )
        if read_results or dry_run_match:
            if confirm:
                wait_start = time.perf_counter()
                result["result_wait_before_read"] = wait_after_query_confirm(
                    jab, query_cfg
                )
                timings.add(
                    "result_wait_before_read", time.perf_counter() - wait_start
                )
            extractor = ReceiptNCResultExtractor(config)
            if dry_run_match:
                tables, page_report, dry_run_report = timings.measure(
                    "read_receipt_result_pages_until_match",
                    read_receipt_result_pages_until_match,
                    jab,
                    config,
                    extractor,
                    query_cfg,
                    org_code,
                    business_date=parse_date(end),
                    max_rows=max_rows,
                    max_cols=max_cols,
                    read_columns=receipt_result_read_columns(
                        query_cfg,
                        include_amount_candidates=True,
                    ),
                    write_back=write_back,
                )
            else:
                tables, page_report = timings.measure(
                    "read_receipt_result_pages",
                    read_receipt_result_pages,
                    jab,
                    query_cfg,
                    max_rows=max_rows,
                    max_cols=max_cols,
                    read_columns=receipt_result_read_columns(
                        query_cfg,
                        include_amount_candidates=False,
                    ),
                )
                dry_run_report = None
            result["table_summary"] = [
                {
                    "table_index": table.get("table_index"),
                    "row_count": table.get("row_count"),
                    "col_count": table.get("col_count"),
                }
                for table in tables
            ]
            result["page_report"] = page_report
            timings.measure(
                "guard_receipt_result_tables",
                guard_receipt_result_tables,
                tables,
                query_cfg,
            )
            if read_results:
                rows, issues = timings.measure(
                    "extract_nc_rows",
                    extractor.extract_by_indexes,
                    tables,
                    extractor.config.result_column_indexes["payer_name"],
                    extractor.config.result_column_indexes["original_amount"],
                )
                result["nc_rows"] = rows
                result["extract_issues"] = issues
            if dry_run_match:
                result["dry_run_match"] = dry_run_report
        timings.add("total", time.perf_counter() - total_start)
        result["timings"] = timings.items
        return result
    finally:
        jab.close()


def guard_receipt_parent_page(jab, config, query_cfg):
    guard_cfg = query_cfg.get("page_guard") or {}
    if not bool(guard_cfg.get("enabled", True)):
        return {"enabled": False, "ok": True}

    jab.ensure_started()
    state_label = (config.get("receipt_entry") or {}).get("state_label", "")
    if not state_label:
        raise ReceiptPageGuardError("receipt_entry.state_label is required")

    context, vm_id, owned_contexts = jab.find_context(
        state_label,
        roles=guard_cfg.get("state_label_roles", ()),
        timeout=float(guard_cfg.get("state_label_timeout", 1.5)),
        require_showing=bool(guard_cfg.get("state_label_require_showing", False)),
        window_title=guard_cfg.get("window_title"),
        window_class=guard_cfg.get("window_class"),
        visible_only=bool(guard_cfg.get("visible_only", True)),
    )
    if context:
        jab.release_contexts(vm_id, owned_contexts)
        return {"enabled": True, "ok": True, "state_label": state_label}

    raise ReceiptPageGuardError(
        f"当前 NC 页面未检测到目标页标识: {state_label!r}，拒绝执行收款查询/写回"
    )


def guard_receipt_result_tables(tables, query_cfg):
    guard_cfg = query_cfg.get("result_guard") or {}
    if not bool(guard_cfg.get("enabled", True)):
        return {"enabled": False, "ok": True}

    indexes = query_cfg.get("result_column_indexes") or {}
    document_type_col = int(guard_cfg.get("document_type_column", 2))
    document_type = str(guard_cfg.get("document_type", "收款单"))
    name_column = int(
        (query_cfg.get("result_column_indexes") or {}).get("payer_name", 2)
    )
    guard_name_column = bool(
        guard_cfg.get("name_column_must_not_equal_document_type", True)
    )
    blocked_keywords = tuple(guard_cfg.get("blocked_keywords", ("应收款", "应付款")))
    max_samples = int(guard_cfg.get("max_samples", 20))
    samples = []
    name_samples = []
    blocked = []
    expected = 0

    for table in tables:
        if table.get("col_count", 0) < max(indexes.values(), default=0) + 1:
            continue
        for row in table.get("rows") or []:
            cells = row.get("cells") or []
            value = first_non_empty_cell_at(cells, document_type_col)
            if not value:
                continue
            name_value = first_non_empty_cell_at(cells, name_column)
            if len(samples) < max_samples:
                samples.append(value)
            if len(name_samples) < max_samples and name_value:
                name_samples.append(name_value)
            if document_type in value:
                expected += 1
            row_text = "\t".join(str(cell or "") for cell in cells)
            row_blocked = [
                keyword for keyword in blocked_keywords if keyword in row_text
            ]
            if row_blocked:
                blocked.extend(row_blocked)

    if blocked:
        raise ReceiptPageGuardError(
            "收款查询结果表疑似来自错误页面，检测到禁用单据类型: "
            f"{sorted(set(blocked))[:10]}"
        )
    if samples and expected == 0:
        raise ReceiptPageGuardError(
            "收款查询结果表未检测到目标单据类型: "
            f"expected={document_type!r} samples={samples[:10]}"
        )
    if (
        guard_name_column
        and name_samples
        and all(document_type in value for value in name_samples)
    ):
        raise ReceiptPageGuardError(
            "收款查询结果表匹配名称列疑似配置到单据类型列: "
            f"name_column={name_column} samples={name_samples[:10]}"
        )
    return {
        "enabled": True,
        "ok": True,
        "document_type": document_type,
        "samples": samples,
        "name_samples": name_samples,
    }


def read_receipt_result_pages(
    jab,
    query_cfg,
    max_rows=500,
    max_cols=80,
    read_columns=None,
):
    pagination = query_cfg.get("pagination") or {}
    if not pagination:
        tables = read_receipt_tables(
            jab,
            query_cfg,
            max_rows=max_rows,
            max_cols=max_cols,
            read_columns=read_columns,
        )
        return tables, {"enabled": False, "pages": []}

    setup_start = time.perf_counter()
    setup_report = set_receipt_page_size(jab, query_cfg)
    setup_report["setup_seconds"] = round(time.perf_counter() - setup_start, 3)
    window_class = setup_report["window_class"]
    page_label_path = setup_report["page_label_path"]
    next_page_path = setup_report["next_page_button_path"]
    pager_hwnd = setup_report["pager_hwnd"]
    after_label = setup_report["after_label"]
    before_label = setup_report["before_label"]
    runtime_query_cfg = with_runtime_pagination_paths(query_cfg, setup_report)

    page_info = parse_page_label(after_label or before_label or "")
    total_pages = page_info.get("total_pages") or 1
    total_records = page_info.get("total_records")
    page_size = int(setup_report.get("page_size") or pagination.get("page_size", 500))
    if total_records is not None:
        planned_pages = 1 if total_records <= page_size else total_pages
        pagination_plan_reason = (
            "total_records_within_page_size"
            if total_records <= page_size
            else "total_records_exceed_page_size"
        )
    else:
        planned_pages = total_pages
        pagination_plan_reason = "total_records_unknown_use_total_pages"
    page_limit = min(planned_pages, int(pagination.get("max_pages", planned_pages)))

    collected = []
    page_reports = []
    seen_documents = set()
    current_stability = setup_report.get("after_stability")
    for page_number in range(1, page_limit + 1):
        label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
        wait_before_read_start = time.perf_counter()
        time.sleep(float(pagination.get("wait_before_read", 0.0)))
        wait_before_read_seconds = time.perf_counter() - wait_before_read_start
        read_start = time.perf_counter()
        tables = read_receipt_result_tables_runtime(
            jab,
            runtime_query_cfg,
            setup_report,
            max_rows=max_rows,
            max_cols=max_cols,
            read_columns=read_columns,
        )
        read_tables_seconds = time.perf_counter() - read_start
        page_tables = []
        for table in tables:
            rows = []
            for row in table.get("rows") or []:
                document_key = first_non_empty_cell(row.get("cells") or [])
                if document_key and document_key in seen_documents:
                    continue
                if document_key:
                    seen_documents.add(document_key)
                rows.append(row)
            if rows:
                page_tables.append({**table, "rows": rows, "row_count": len(rows)})
        collected.extend(page_tables)
        page_reports.append(
            {
                "page": page_number,
                "label": label,
                "tables": [
                    {
                        "table_index": table.get("table_index"),
                        "row_count": table.get("row_count"),
                        "col_count": table.get("col_count"),
                        "read_method": table.get("read_method"),
                        "path": table.get("path"),
                    }
                    for table in tables
                ],
                "stability": current_stability,
                "wait_before_read_seconds": round(wait_before_read_seconds, 3),
                "read_tables_seconds": round(read_tables_seconds, 3),
            }
        )
        wait_after_page_read_start = time.perf_counter()
        time.sleep(float(pagination.get("wait_after_page_read", 0.0)))
        page_reports[-1]["wait_after_page_read_seconds"] = round(
            time.perf_counter() - wait_after_page_read_start, 3
        )
        if page_number >= page_limit:
            break
        if not pager_hwnd or not label:
            page_reports[-1]["next_page_ok"] = False
            page_reports[-1]["next_page_method"] = "blocked_no_pager_scope"
            break
        next_start = time.perf_counter()
        ok, method = click_next_page(
            jab,
            pagination,
            next_page_path,
            window_class,
            pager_hwnd,
        )
        page_reports[-1]["next_page_seconds"] = round(
            time.perf_counter() - next_start, 3
        )
        if not ok:
            page_reports[-1]["next_page_ok"] = False
            page_reports[-1]["next_page_method"] = method
            break
        page_reports[-1]["next_page_ok"] = True
        page_reports[-1]["next_page_method"] = method
        stability_start = time.perf_counter()
        current_stability = wait_receipt_result_stable(
            jab, runtime_query_cfg, pager_hwnd=pager_hwnd
        )
        page_reports[-1]["after_next_stability_seconds"] = round(
            time.perf_counter() - stability_start, 3
        )
        page_reports[-1]["after_next_stability"] = current_stability

    return collected, {
        "enabled": True,
        **setup_report,
        "total_pages": total_pages,
        "total_records": total_records,
        "pagination_plan_reason": pagination_plan_reason,
        "planned_pages": planned_pages,
        "pages": page_reports,
    }


def read_receipt_result_pages_until_match(
    jab,
    config,
    extractor,
    query_cfg,
    org_code,
    business_date,
    max_rows=500,
    max_cols=80,
    read_columns=None,
    write_back=False,
):
    rows, candidates, excel_issues = ReceiptEntryWorkbook(config).preview_rows(
        today=business_date
    )
    org_candidates = [row for row in candidates if row.organization_code == org_code]
    target_rows = [
        row for row in org_candidates if row.receipt_date == business_date
    ] or org_candidates
    tables, page_report, match_snapshot = read_receipt_result_pages_incremental(
        jab,
        query_cfg,
        extractor,
        target_rows,
        max_rows=max_rows,
        max_cols=max_cols,
        read_columns=read_columns,
    )
    dry_run_report = build_dry_run_match_report_from_preview(
        config,
        extractor,
        tables,
        org_code,
        business_date,
        rows,
        candidates,
        excel_issues,
        write_back=write_back,
        target_rows=target_rows,
        configured_match_snapshot=match_snapshot,
    )
    dry_run_report["paging_match"] = page_report.get("match_progress")
    dry_run_report["target_rows"] = [row.row for row in target_rows]
    return tables, page_report, dry_run_report


def read_receipt_result_pages_incremental(
    jab,
    query_cfg,
    extractor,
    target_rows,
    max_rows=500,
    max_cols=80,
    read_columns=None,
):
    pagination = query_cfg.get("pagination") or {}
    if not pagination:
        tables = read_receipt_tables(
            jab,
            query_cfg,
            max_rows=max_rows,
            max_cols=max_cols,
            read_columns=read_columns,
        )
        return (
            tables,
            {
                "enabled": False,
                "pages": [],
                "match_progress": {
                    "enabled": True,
                    "stop_reason": "pagination_disabled",
                    "target_rows": [row.row for row in target_rows],
                },
            },
            None,
        )

    setup_start = time.perf_counter()
    setup_report = set_receipt_page_size(jab, query_cfg)
    setup_report["setup_seconds"] = round(time.perf_counter() - setup_start, 3)
    window_class = setup_report["window_class"]
    page_label_path = setup_report["page_label_path"]
    next_page_path = setup_report["next_page_button_path"]
    pager_hwnd = setup_report["pager_hwnd"]
    runtime_query_cfg = with_runtime_pagination_paths(query_cfg, setup_report)

    page_info = parse_page_label(
        setup_report.get("after_label") or setup_report.get("before_label") or ""
    )
    total_pages = page_info.get("total_pages") or 1
    total_records = page_info.get("total_records")
    page_size = int(setup_report.get("page_size") or pagination.get("page_size", 500))
    page_limit = min(total_pages, int(pagination.get("max_pages", total_pages)))

    collected = []
    page_reports = []
    seen_documents = set()
    current_stability = setup_report.get("after_stability")
    matcher = ReceiptEntryDryRunMatcher()
    matched = {}
    match_issues = []
    nc_rows = []
    extract_issues = []
    match_stop_reason = None
    target_row_numbers = [row.row for row in target_rows]

    for page_number in range(1, page_limit + 1):
        label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
        wait_before_read_start = time.perf_counter()
        time.sleep(float(pagination.get("wait_before_read", 0.0)))
        wait_before_read_seconds = time.perf_counter() - wait_before_read_start
        read_start = time.perf_counter()
        tables = read_receipt_result_tables_runtime(
            jab,
            runtime_query_cfg,
            setup_report,
            max_rows=max_rows,
            max_cols=max_cols,
            read_columns=read_columns,
        )
        read_tables_seconds = time.perf_counter() - read_start
        page_tables = dedupe_page_tables(tables, seen_documents)
        collected.extend(page_tables)
        nc_rows, extract_issues = extractor.extract_by_indexes(
            collected,
            extractor.config.result_column_indexes["payer_name"],
            amount_column=extractor.config.result_column_indexes["original_amount"],
        )
        matched, match_issues = matcher.match(target_rows, nc_rows)
        stop_state = evaluate_paging_match_stop(
            target_rows,
            matched,
            match_issues,
            total_records,
            page_size,
            page_number,
            page_limit,
        )
        page_reports.append(
            {
                "page": page_number,
                "label": label,
                "tables": [
                    {
                        "table_index": table.get("table_index"),
                        "row_count": table.get("row_count"),
                        "col_count": table.get("col_count"),
                        "read_method": table.get("read_method"),
                        "path": table.get("path"),
                    }
                    for table in tables
                ],
                "stability": current_stability,
                "wait_before_read_seconds": round(wait_before_read_seconds, 3),
                "read_tables_seconds": round(read_tables_seconds, 3),
                "match": {
                    "target_rows": target_row_numbers,
                    "matched_rows": sorted(matched),
                    "unresolved_rows": stop_state["unresolved_rows"],
                    "issues": [
                        {
                            "excel_row": issue.excel_row,
                            "reason": issue.reason,
                            "nc_rows": issue.nc_rows,
                        }
                        for issue in match_issues[:20]
                    ],
                    "extract_issues": [
                        {
                            "table_index": issue.table_index,
                            "row_index": issue.row_index,
                            "reason": issue.reason,
                        }
                        for issue in extract_issues[:20]
                    ],
                    "stop_reason": stop_state["reason"],
                },
            }
        )
        wait_after_page_read_start = time.perf_counter()
        time.sleep(float(pagination.get("wait_after_page_read", 0.0)))
        page_reports[-1]["wait_after_page_read_seconds"] = round(
            time.perf_counter() - wait_after_page_read_start, 3
        )
        if stop_state["stop"]:
            match_stop_reason = stop_state["reason"]
            break
        if page_number >= page_limit:
            match_stop_reason = "page_limit_reached"
            break
        if not pager_hwnd or not label:
            page_reports[-1]["next_page_ok"] = False
            page_reports[-1]["next_page_method"] = "blocked_no_pager_scope"
            match_stop_reason = "blocked_no_pager_scope"
            break
        next_start = time.perf_counter()
        ok, method = click_next_page(
            jab,
            pagination,
            next_page_path,
            window_class,
            pager_hwnd,
        )
        page_reports[-1]["next_page_seconds"] = round(
            time.perf_counter() - next_start, 3
        )
        if not ok:
            page_reports[-1]["next_page_ok"] = False
            page_reports[-1]["next_page_method"] = method
            match_stop_reason = method
            break
        page_reports[-1]["next_page_ok"] = True
        page_reports[-1]["next_page_method"] = method
        stability_start = time.perf_counter()
        current_stability = wait_receipt_result_stable(
            jab, runtime_query_cfg, pager_hwnd=pager_hwnd
        )
        page_reports[-1]["after_next_stability_seconds"] = round(
            time.perf_counter() - stability_start, 3
        )
        page_reports[-1]["after_next_stability"] = current_stability

    if match_stop_reason is None:
        match_stop_reason = "no_pages_read"

    match_snapshot = {
        "nc_rows": nc_rows,
        "extract_issues": extract_issues,
        "matched": matched,
        "match_issues": match_issues,
    }
    return (
        collected,
        {
            "enabled": True,
            **setup_report,
            "total_pages": total_pages,
            "total_records": total_records,
            "pagination_plan_reason": (
                "total_records_within_page_size"
                if total_records is not None and total_records <= page_size
                else "match_driven"
            ),
            "planned_pages": page_limit,
            "pages": page_reports,
            "match_progress": {
                "enabled": True,
                "target_rows": target_row_numbers,
                "matched_rows": sorted(matched),
                "unresolved_rows": unresolved_excel_rows(
                    target_rows, matched, match_issues
                ),
                "stop_reason": match_stop_reason,
            },
        },
        match_snapshot,
    )


def dedupe_page_tables(tables, seen_documents):
    page_tables = []
    for table in tables:
        rows = []
        for row in table.get("rows") or []:
            document_key = first_non_empty_cell(row.get("cells") or [])
            if document_key and document_key in seen_documents:
                continue
            if document_key:
                seen_documents.add(document_key)
            rows.append(row)
        if rows:
            page_tables.append({**table, "rows": rows, "row_count": len(rows)})
    return page_tables


def evaluate_paging_match_stop(
    excel_rows,
    matched,
    match_issues,
    total_records,
    page_size,
    page_number,
    page_limit,
):
    unresolved = unresolved_excel_rows(excel_rows, matched, match_issues)
    if not unresolved:
        return {"stop": True, "reason": "all_targets_resolved", "unresolved_rows": []}
    if total_records is not None and total_records <= page_size:
        return {
            "stop": True,
            "reason": "total_records_within_page_size",
            "unresolved_rows": unresolved,
        }
    if page_number >= page_limit:
        return {
            "stop": True,
            "reason": "page_limit_reached",
            "unresolved_rows": unresolved,
        }
    return {"stop": False, "reason": "unresolved_targets", "unresolved_rows": unresolved}


def unresolved_excel_rows(excel_rows, matched, match_issues):
    resolved = set(matched)
    for issue in match_issues:
        if issue.reason.startswith("重复"):
            resolved.add(issue.excel_row)
    return [row.row for row in excel_rows if row.row not in resolved]


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
    pager_window = jab.wait_context_by_path(
        page_label_path,
        class_name=window_class,
        role="label",
        timeout=float(pagination.get("pager_scope_timeout", 2.0)),
        require_showing=True,
        require_valid_bounds=False,
    )
    pager_hwnd = (
        int(pager_window.get("hwnd"))
        if pager_window and pager_window.get("hwnd") is not None
        else resolved_paths.get("pager_hwnd")
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
    before_label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
    before_page_size_text = read_page_size_text(
        jab, page_size_path, window_class, pager_hwnd
    )
    current_page_size = parse_int_text(before_page_size_text)
    page_size_changed = current_page_size != page_size
    wait_before_page_size_seconds = 0.0
    set_page_size_text_seconds = 0.0
    enter_seconds = 0.0
    page_size_ok = True
    if page_size_changed:
        wait_before_page_size_start = time.perf_counter()
        time.sleep(float(pagination.get("wait_before_page_size", 0.0)))
        wait_before_page_size_seconds = time.perf_counter() - wait_before_page_size_start
        set_text_start = time.perf_counter()
        page_size_ok = jab.set_text_by_path(
            page_size_path,
            str(page_size),
            class_name=window_class,
            scope_hwnd=pager_hwnd,
            role="text",
            timeout=2,
            require_showing=True,
            require_valid_bounds=False,
        )
        set_page_size_text_seconds = time.perf_counter() - set_text_start
    if page_size_ok and page_size_changed:
        enter_start = time.perf_counter()
        jab.press_key("enter", wait=float(pagination.get("wait_after_page_size", 2.0)))
        enter_seconds = time.perf_counter() - enter_start
    if page_size_changed or bool(pagination.get("ready_check_when_page_size_ok", False)):
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
        "page_size_text_path": pagination["page_size_text_path"],
        "next_page_button_path": pagination["next_page_button_path"],
    }
    if dynamic:
        report["dynamic_resolution"] = dynamic.get("resolution")
        report["result_page_probe"] = dynamic.get("resolution")
        report["dynamic_diagnostics"] = dynamic.get("diagnostics")
    return report


def validate_receipt_pagination_path_report(jab, pagination, report, resolution):
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
    prefix = report.get("result_area_prefix") or infer_result_area_prefix_from_page_path(
        report["page_label_path"]
    )
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
    diagnostics = {"candidates": candidates[:10]}
    for candidate in candidates:
        prefix = infer_result_area_prefix_from_table_path(candidate["path"])
        if not prefix:
            continue
        paths = {
            "page_label_path": join_context_path(prefix, RECEIPT_PAGE_LABEL_PATH_SUFFIX),
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
        "page_size_text_path": pagination["page_size_text_path"],
        "next_page_button_path": pagination["next_page_button_path"],
    }


def enumerate_receipt_result_table_paths(jab, query_cfg, window_class):
    if not hasattr(jab, "dll"):
        return []
    jab.ensure_started()
    if jab.dll is None:
        return []
    result = []
    table_index = 0
    for hwnd, title, class_name, pid, visible in jab.get_scoped_windows(
        include_children=True
    ):
        if class_name != window_class:
            continue
        if not visible:
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vm_id), ctypes.byref(root_context)
        ):
            continue

        tables = find_table_paths_in_context(
            jab,
            vm_id.value,
            root_context.value,
            depth=0,
            index_path=[],
            owned_contexts=[],
        )
        contexts_to_release = []
        for table in tables:
            table_info = table["table_info"]
            candidate = {
                "table_index": table_index,
                "path": "0" + "".join(f".{index}" for index in table["index_path"]),
                "hwnd": int(hwnd),
                "window_title": title,
                "window_class": class_name,
                "window_visible": visible,
                "pid": pid,
                "row_count": int(table_info.rowCount),
                "col_count": int(table_info.columnCount),
            }
            table_index += 1
            if is_receipt_result_table_candidate(candidate, query_cfg):
                result.append(candidate)
            contexts_to_release.extend(table["owned_contexts"])
        if contexts_to_release:
            unique_contexts = list(dict.fromkeys(contexts_to_release))
            jab.release_contexts(vm_id.value, unique_contexts)
    return result


def find_table_paths_in_context(
    jab,
    vm_id,
    context,
    depth,
    index_path,
    owned_contexts,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return []

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if role == "table":
        table_info = jab.get_table_info(vm_id, context)
        if table_info and table_info.rowCount > 0 and table_info.columnCount > 0:
            return [
                {
                    "index_path": list(index_path),
                    "owned_contexts": list(owned_contexts),
                    "table_info": table_info,
                }
            ]
        return []

    if depth >= jab.max_depth:
        return []

    tables = []
    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_tables = find_table_paths_in_context(
            jab,
            vm_id,
            child,
            depth + 1,
            index_path + [index],
            owned_contexts + [child],
        )
        if child_tables:
            tables.extend(child_tables)
        else:
            jab.release_contexts(vm_id, [child])
    return tables


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
    return [int(part) for part in str(path).split(".") if part != ""]


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


def summarize_receipt_tables(jab, query_cfg, scope_hwnd=None):
    indexes = receipt_result_read_columns(query_cfg)
    min_cols = max(indexes) + 1 if indexes else None
    exact_cols = query_cfg.get("result_table_cols")
    if hasattr(jab, "read_table_summaries"):
        return [
            table
            for table in jab.read_table_summaries(
                min_rows=1,
                min_cols=min_cols,
                scope_hwnd=scope_hwnd,
                exact_cols=exact_cols,
            )
            if is_receipt_result_table_candidate(table, query_cfg)
        ]
    tables = read_receipt_tables(
        jab,
        query_cfg,
        max_rows=0,
        max_cols=0,
        read_columns=[],
        scope_hwnd=scope_hwnd,
    )
    return [
        {
            "table_index": table.get("table_index"),
            "row_count": table.get("row_count"),
            "col_count": table.get("col_count"),
        }
        for table in tables
        if is_receipt_result_table_candidate(table, query_cfg)
    ]


def receipt_result_read_columns(query_cfg, include_amount_candidates=False):
    indexes = query_cfg.get("result_column_indexes") or {}
    columns = {
        int(column)
        for column in indexes.values()
        if isinstance(column, int) and column >= 0
    }
    if include_amount_candidates:
        columns.update({6, 7, 8})
    return sorted(columns)


def read_receipt_tables(
    jab,
    query_cfg,
    max_rows=500,
    max_cols=80,
    read_columns=None,
    scope_hwnd=None,
):
    exact_cols = query_cfg.get("result_table_cols")
    if read_columns and hasattr(jab, "read_all_table_selected_columns"):
        tables = jab.read_all_table_selected_columns(
            read_columns,
            max_rows=max_rows,
            min_rows=1,
            min_cols=max(read_columns) + 1,
            scope_hwnd=scope_hwnd,
            exact_cols=exact_cols,
        )
    else:
        tables = jab.read_all_table_cells(
            max_rows=max_rows,
            max_cols=max_cols,
            scope_hwnd=scope_hwnd,
            exact_cols=exact_cols,
        )
    return [
        table
        for table in tables
        if is_receipt_result_table_candidate(table, query_cfg)
    ]


def read_receipt_result_table_by_path(
    jab,
    table_path,
    table_hwnd,
    query_cfg,
    max_rows=500,
    max_cols=80,
    read_columns=None,
    table_index=0,
):
    if not table_path or not table_hwnd:
        return []
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        table_path,
        class_name=(query_cfg.get("pagination") or {}).get(
            "window_class", "SunAwtCanvas"
        ),
        scope_hwnd=table_hwnd,
        role="table",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return []
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return []
        exact_cols = query_cfg.get("result_table_cols")
        if exact_cols is not None and table_info.columnCount != int(exact_cols):
            return []
        if read_columns:
            table = jab.read_table_selected_columns_from_context(
                table_index,
                context,
                vm_id,
                table_info,
                window_info,
                sorted({int(column) for column in read_columns}),
                max_rows=max_rows,
            )
        else:
            table = jab.read_table_cells_from_context(
                table_index,
                context,
                vm_id,
                table_info,
                window_info,
                max_rows=max_rows,
                max_cols=max_cols,
            )
        table["read_method"] = "result-table-path"
        table["path"] = table_path
        return [table] if is_receipt_result_table_candidate(table, query_cfg) else []
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def read_receipt_result_tables_runtime(
    jab,
    query_cfg,
    setup_report,
    max_rows=500,
    max_cols=80,
    read_columns=None,
):
    tables = read_receipt_result_table_by_path(
        jab,
        setup_report.get("result_table_path"),
        setup_report.get("pager_hwnd"),
        query_cfg,
        max_rows=max_rows,
        max_cols=max_cols,
        read_columns=read_columns,
    )
    if tables:
        return tables
    return read_receipt_tables(
        jab,
        query_cfg,
        max_rows=max_rows,
        max_cols=max_cols,
        read_columns=read_columns,
        scope_hwnd=setup_report.get("pager_hwnd"),
    )


def is_receipt_result_table_candidate(table, query_cfg):
    if table.get("row_count", 0) <= 0:
        return False
    result_table_cols = int(query_cfg.get("result_table_cols", 41))
    return int(table.get("col_count", 0)) == result_table_cols


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


def read_page_label(jab, path, window_class, scope_hwnd=None):
    return jab.get_text_by_path(
        path,
        class_name=window_class,
        scope_hwnd=scope_hwnd,
        role="label",
        timeout=1,
        require_showing=True,
        require_valid_bounds=False,
    )


def read_page_size_text(jab, path, window_class, scope_hwnd=None):
    return jab.get_text_by_path(
        path,
        class_name=window_class,
        scope_hwnd=scope_hwnd,
        role="text",
        timeout=1,
        require_showing=True,
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


def first_non_empty_cell(cells):
    for cell in cells:
        text = str(cell or "").strip()
        if text:
            return text
    return ""


def first_non_empty_cell_at(cells, column):
    if column >= len(cells):
        return ""
    return str(cells[column] or "").strip()


def build_dry_run_match_report(
    config,
    extractor,
    tables,
    org_code,
    business_date,
    write_back=False,
):
    rows, candidates, excel_issues = ReceiptEntryWorkbook(config).preview_rows(
        today=business_date
    )
    return build_dry_run_match_report_from_preview(
        config,
        extractor,
        tables,
        org_code,
        business_date,
        rows,
        candidates,
        excel_issues,
        write_back=write_back,
    )


def build_dry_run_match_report_from_preview(
    config,
    extractor,
    tables,
    org_code,
    business_date,
    rows,
    candidates,
    excel_issues,
    write_back=False,
    target_rows=None,
    configured_match_snapshot=None,
):
    org_candidates = [row for row in candidates if row.organization_code == org_code]
    match_candidates = target_rows or org_candidates
    report = {
        "business_date": business_date.isoformat(),
        "excel_rows": len(rows),
        "excel_candidates": len(candidates),
        "org_candidates": len(org_candidates),
        "match_candidates": len(match_candidates),
        "excel_issues": len(excel_issues),
        "candidate_banks": dict(
            sorted(Counter(row.bank for row in match_candidates).items())
        ),
        "write_back": {"enabled": bool(write_back), "updated": 0, "rows": []},
        "variants": [],
    }
    matcher = ReceiptEntryDryRunMatcher()
    configured_amount_column = extractor.config.result_column_indexes["original_amount"]
    configured_name_column = extractor.config.result_column_indexes["payer_name"]
    dry_run_all_variants = bool(
        (config.get("receipt_entry") or {})
        .get("query", {})
        .get("dry_run_all_variants", False)
    )
    if dry_run_all_variants:
        amount_columns = unique_ordered([configured_amount_column, 8, 6, 7])
        name_columns = unique_ordered([configured_name_column, 2, 4, 19])
    else:
        amount_columns = [configured_amount_column]
        name_columns = [configured_name_column]
    for column in name_columns:
        for amount_column in amount_columns:
            variant_name = f"name_col{column}_amount_col{amount_column}"
            is_configured_variant = (
                column == configured_name_column
                and amount_column == configured_amount_column
            )
            if is_configured_variant and configured_match_snapshot:
                nc_rows = configured_match_snapshot["nc_rows"]
                extract_issues = configured_match_snapshot["extract_issues"]
                matched = configured_match_snapshot["matched"]
                match_issues = configured_match_snapshot["match_issues"]
                source = "incremental"
            else:
                nc_rows, extract_issues = extractor.extract_by_indexes(
                    tables,
                    column,
                    amount_column=amount_column,
                )
                matched, match_issues = matcher.match(match_candidates, nc_rows)
                source = "computed"
            report["variants"].append(
                {
                    "name": variant_name,
                    "name_column": column,
                    "amount_column": amount_column,
                    "source": source,
                    "nc_rows": len(nc_rows),
                    "nc_summary": summarize_nc_rows(nc_rows),
                    "match_diagnostics": diagnose_match_inputs(match_candidates, nc_rows),
                    "extract_issues": len(extract_issues),
                    "matches": len(matched),
                    "match_issues": len(match_issues),
                    "matched_excel_rows": sorted(matched.keys())[:20],
                    "issue_samples": [
                        {
                            "excel_row": issue.excel_row,
                            "reason": issue.reason,
                            "nc_rows": issue.nc_rows,
                        }
                        for issue in match_issues[:20]
                    ],
                    "extract_issue_samples": [
                        {
                            "table_index": issue.table_index,
                            "row_index": issue.row_index,
                            "reason": issue.reason,
                        }
                        for issue in extract_issues[:20]
                    ],
                }
            )
            if is_configured_variant:
                report["write_back"] = build_receipt_write_back_report(
                    config,
                    match_candidates,
                    matched,
                    match_issues,
                    enabled=write_back,
                )
    return report


def unique_ordered(values):
    result = []
    seen = set()
    for value in values:
        if value in seen:
            continue
        seen.add(value)
        result.append(value)
    return result


def build_receipt_write_back_report(
    config,
    excel_rows,
    matched,
    match_issues,
    enabled=False,
):
    issue_by_row = {issue.excel_row: issue for issue in match_issues}
    statuses = {}
    duplicate_rows = []
    exception_rows = []
    for excel_row in excel_rows:
        if excel_row.row in matched:
            statuses[excel_row.row] = "已做过"
            continue
        issue = issue_by_row.get(excel_row.row)
        if issue and issue.reason == format_receipt_not_found_reason():
            statuses[excel_row.row] = "未做过"
        elif issue:
            statuses[excel_row.row] = issue.reason
            exception_rows.append(excel_row.row)
            if issue.reason.startswith("重复"):
                duplicate_rows.append(excel_row.row)

    report = {
        "enabled": bool(enabled),
        "planned": len(statuses),
        "matched_rows": sorted(matched),
        "not_found_rows": sorted(
            row for row, status in statuses.items() if status == "未做过"
        ),
        "duplicate_rows": sorted(duplicate_rows),
        "exception_rows": sorted(exception_rows),
        "skipped_duplicate_rows": sorted(duplicate_rows),
        "updated": 0,
        "rows": [],
    }
    if enabled:
        write_result = ReceiptEntryWorkbook(config).write_nc_done_statuses(statuses)
        report["updated"] = write_result["updated"]
        report["rows"] = write_result["rows"]
    return report


def diagnose_match_inputs(excel_rows, nc_rows):
    nc_by_amount = {}
    for nc_row in nc_rows:
        nc_by_amount.setdefault(nc_row.original_amount, []).append(nc_row)

    amount_only_hits = 0
    name_amount_hits = 0
    duplicate_hits = 0
    name_only_hits = 0
    name_amount_samples = []
    name_mismatch_samples = []
    amount_mismatch_samples = []
    no_amount_samples = []
    for excel_row in excel_rows:
        amount_candidates = nc_by_amount.get(excel_row.raw_amount, [])
        if not amount_candidates:
            name_candidates = [
                nc_row
                for nc_row in nc_rows
                if names_match(excel_row.payer_name, nc_row.name)
            ]
            if name_candidates:
                name_only_hits += 1
                if len(amount_mismatch_samples) < 10:
                    amount_mismatch_samples.append(
                        {
                            "excel_row": excel_row.row,
                            "excel_amount": str(excel_row.raw_amount),
                            "excel_name": excel_row.payer_name,
                            "reason": format_receipt_name_amount_mismatch_reason(
                                excel_amount=excel_row.raw_amount,
                                excel_name=excel_row.payer_name,
                                nc_amounts=[
                                    row.original_amount for row in name_candidates
                                ],
                            ),
                            "nc_amounts": [
                                str(row.original_amount) for row in name_candidates[:5]
                            ],
                            "nc_rows": [row.row_index for row in name_candidates[:5]],
                        }
                    )
                continue
            if len(no_amount_samples) < 10:
                no_amount_samples.append(
                    {
                        "excel_row": excel_row.row,
                        "excel_amount": str(excel_row.raw_amount),
                        "excel_name": excel_row.payer_name,
                    }
                )
            continue

        amount_only_hits += 1
        matched_names = [
            nc_row
            for nc_row in amount_candidates
            if names_match(excel_row.payer_name, nc_row.name)
        ]
        if len(matched_names) == 1:
            name_amount_hits += 1
            if len(name_amount_samples) < 10:
                name_amount_samples.append(
                    {
                        "excel_row": excel_row.row,
                        "excel_amount": str(excel_row.raw_amount),
                        "excel_name": excel_row.payer_name,
                        "nc_name": matched_names[0].name,
                        "nc_row": matched_names[0].row_index,
                    }
                )
        elif len(matched_names) > 1:
            duplicate_hits += 1
            if len(name_mismatch_samples) < 10:
                name_mismatch_samples.append(
                    {
                        "excel_row": excel_row.row,
                        "excel_amount": str(excel_row.raw_amount),
                        "excel_name": excel_row.payer_name,
                        "reason": format_receipt_duplicate_reason(len(matched_names)),
                        "nc_names": [row.name for row in matched_names[:5]],
                        "nc_rows": [row.row_index for row in matched_names[:5]],
                    }
                )
        elif len(name_mismatch_samples) < 10:
            name_mismatch_samples.append(
                {
                    "excel_row": excel_row.row,
                    "excel_amount": str(excel_row.raw_amount),
                    "excel_name": excel_row.payer_name,
                    "reason": format_receipt_amount_name_mismatch_reason(
                        excel_amount=excel_row.raw_amount,
                        excel_name=excel_row.payer_name,
                        nc_names=[row.name for row in amount_candidates],
                    ),
                    "nc_names": [row.name for row in amount_candidates[:5]],
                    "nc_rows": [row.row_index for row in amount_candidates[:5]],
                }
            )
    return {
        "amount_only_hits": amount_only_hits,
        "name_amount_hits": name_amount_hits,
        "duplicate_hits": duplicate_hits,
        "name_only_hits": name_only_hits,
        "name_amount_samples": name_amount_samples,
        "name_mismatch_samples": name_mismatch_samples,
        "amount_mismatch_samples": amount_mismatch_samples,
        "no_amount_samples": no_amount_samples,
    }


def summarize_nc_rows(nc_rows):
    if not nc_rows:
        return {
            "amount_min": None,
            "amount_max": None,
            "name_samples": [],
        }
    amounts = [row.original_amount for row in nc_rows]
    names = []
    seen = set()
    for row in nc_rows:
        name = row.name.strip()
        if not name or name in seen:
            continue
        seen.add(name)
        names.append(name)
        if len(names) >= 10:
            break
    return {
        "amount_min": str(min(amounts)),
        "amount_max": str(max(amounts)),
        "name_samples": names,
    }


def main():
    parser = argparse.ArgumentParser(description="Fill NC receipt query conditions")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--org-code", required=True)
    parser.add_argument("--date-from", default=None)
    parser.add_argument("--date-to", default=None)
    parser.add_argument(
        "--confirm",
        action="store_true",
        help="click query confirm after filling conditions",
    )
    parser.add_argument(
        "--read-results",
        action="store_true",
        help="read visible NC table rows after filling/querying",
    )
    parser.add_argument(
        "--dry-run-match",
        action="store_true",
        help="compare Excel candidates against NC result columns without writing Excel",
    )
    parser.add_argument(
        "--write-back",
        action="store_true",
        help="write matched receipt statuses to Excel; requires --dry-run-match",
    )
    parser.add_argument(
        "--include-filled-status",
        action="store_true",
        help="include Excel rows that already have NC status; useful for overwrite reruns",
    )
    parser.add_argument(
        "--no-open-query",
        action="store_true",
        help="do not press F3; require the receipt query dialog to already be open",
    )
    parser.add_argument("--max-rows", type=int, default=500)
    parser.add_argument("--max-cols", type=int, default=80)
    parser.add_argument(
        "--probe-stage",
        choices=("query", "page-size", "sample-read"),
        default=None,
        help="run a narrow NC stability probe instead of the full dry-run flow",
    )
    args = parser.parse_args()

    config = load_config(args.config)
    if args.write_back and not args.dry_run_match:
        raise SystemExit("--write-back requires --dry-run-match")
    if args.include_filled_status:
        config = deepcopy(config)
        config["receipt_entry"].setdefault("candidate_check", {})[
            "only_blank_status"
        ] = False
    confirm = args.confirm
    read_results = args.read_results
    dry_run_match = args.dry_run_match
    set_page_size_only = False
    max_rows = args.max_rows
    max_cols = args.max_cols
    if args.probe_stage:
        confirm = True
        read_results = False
        dry_run_match = False
        if args.probe_stage == "page-size":
            set_page_size_only = True
        elif args.probe_stage == "sample-read":
            read_results = True
            max_rows = min(max_rows, 20)
            config = deepcopy(config)
            pagination = config["receipt_entry"]["query"].setdefault("pagination", {})
            pagination["max_pages"] = 1

    try:
        result = fill_receipt_query(
            config,
            org_code=args.org_code,
            date_from=args.date_from,
            date_to=args.date_to,
            confirm=confirm,
            read_results=read_results,
            dry_run_match=dry_run_match,
            skip_open_query=args.no_open_query,
            max_rows=max_rows,
            max_cols=max_cols,
            set_page_size_only=set_page_size_only,
            write_back=args.write_back,
        )
    except ReceiptPageGuardError as exc:
        print(f"receipt page guard failed: {exc}", file=sys.stderr)
        if os.environ.get("RECEIPT_GUARD_TRACEBACK"):
            raise
        return 2
    print(
        "filled receipt query: "
        f"org={result['organization_code']} "
        f"date_from={result['date_from']} date_to={result['date_to']} "
        f"confirm={confirm}"
    )
    timings = result.get("timings") or []
    if timings:
        print("receipt query timings:")
        for item in timings:
            print(f"  {item['name']}: {item['seconds']}s")
    if args.probe_stage:
        print(
            json.dumps(
                {
                    "probe_stage": args.probe_stage,
                    "page_report": result.get("page_report"),
                    "table_summary": result.get("table_summary"),
                },
                ensure_ascii=True,
                indent=2,
            )
        )
    if read_results:
        rows = result["nc_rows"]
        issues = result["extract_issues"]
        print(f"receipt query results: rows={len(rows)} issues={len(issues)}")
        for row in rows[:20]:
            customer = getattr(row, "customer", None)
            if customer is None:
                customer = getattr(row, "name", "")
            print(
                "  nc_row="
                f"{row.row_index} date={row.document_date.isoformat()} "
                f"amount={row.original_amount} customer={customer}"
            )
        for issue in issues[:20]:
            print(
                "  issue="
                f"table={issue.table_index} row={issue.row_index} "
                f"reason={issue.reason}"
            )
    if dry_run_match:
        print(
            json.dumps(
                {
                    "page_report": result.get("page_report"),
                    "dry_run_match": result["dry_run_match"],
                },
                ensure_ascii=True,
                indent=2,
            )
        )


if __name__ == "__main__":
    raise SystemExit(main())
