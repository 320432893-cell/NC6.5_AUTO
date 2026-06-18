# 职责：分页读取收款单结果并按 Excel 目标行做增量匹配停止。
# 不做什么：不填写查询条件，不解析 CLI，不推断分页路径，不写 Excel。
# 允许依赖层：core.receipt_entry/core.receipt_matching、收款单分页/结果表读取模块、匹配报告模块。
# 谁不应该 import：JAB 底层生命周期模块、与收款单查询无关的业务流程、临时探针脚本。

import time
from typing import Any

from core.receipt_entry import ReceiptEntryWorkbook
from core.receipt_matching import ReceiptEntryDryRunMatcher
from tools.receipt_query_pagination import (
    click_next_page,
    parse_page_label,
    read_page_label,
    set_receipt_page_size,
    wait_receipt_result_stable,
)
from tools.receipt_query_pagination_paths import with_runtime_pagination_paths
from tools.receipt_query_report import build_dry_run_match_report_from_preview
from tools.receipt_query_result_tables import (
    first_non_empty_cell,
    read_receipt_result_tables_runtime,
    read_receipt_tables,
)


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
    setup_report: dict[str, Any] = set_receipt_page_size(jab, query_cfg)
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
    return {
        "stop": False,
        "reason": "unresolved_targets",
        "unresolved_rows": unresolved,
    }


def unresolved_excel_rows(excel_rows, matched, match_issues):
    resolved = set(matched)
    for issue in match_issues:
        if issue.reason.startswith("重复"):
            resolved.add(issue.excel_row)
    return [row.row for row in excel_rows if row.row not in resolved]
