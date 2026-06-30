# 职责：按分页读取收款单查询结果表并生成每页读取报告。
# 不做什么：不做 Excel 匹配停止判断，不写 Excel，不解析 CLI 参数，不负责分页路径推断细节。
# 允许依赖层：收款单分页控件模块、收款单结果表读取模块、JABOperator 表格/控件接口。
# 谁不应该 import：与收款单查询无关的业务流程、临时探针脚本。

import time
from typing import Any

from core.receipt_query_pagination import (
    click_next_page,
    parse_page_label,
    read_page_label,
    set_receipt_page_size,
    wait_receipt_result_stable,
)
from core.receipt_query_pagination_paths import with_runtime_pagination_paths
from core.receipt_query_result_tables import (
    first_non_empty_cell,
    read_receipt_result_tables_runtime,
    read_receipt_tables,
)


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
    setup_report: dict[str, Any] = set_receipt_page_size(jab, query_cfg)
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
