from datetime import date
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_nc_extract import ReceiptNCResultExtractor  # noqa: E402
from core.receipt_parsing import parse_date  # noqa: E402

# 兼容导出：旧测试和少量 probe 仍从 receipt_query_fill import 查询辅助函数。
# 清理条件：外部调用迁到 receipt_query_guard/pagination/page_reader/match_reader/result_tables/report 后删除本段 re-export。
from tools.receipt_query_guard import (  # noqa: E402
    ReceiptPageGuardError as ReceiptPageGuardError,
    guard_receipt_parent_page as guard_receipt_parent_page,
    guard_receipt_result_tables as guard_receipt_result_tables,
)
from tools.receipt_query_pagination import (  # noqa: E402
    click_next_page as click_next_page,
    parse_int_text as parse_int_text,
    parse_page_label as parse_page_label,
    read_page_label as read_page_label,
    read_page_size_text as read_page_size_text,
    set_receipt_page_size as set_receipt_page_size,
    wait_after_query_confirm as wait_after_query_confirm,
    wait_receipt_result_ready as wait_receipt_result_ready,
    wait_receipt_result_stable as wait_receipt_result_stable,
)
from tools.receipt_query_pagination_paths import (  # noqa: E402
    infer_result_area_prefix_from_page_path as infer_result_area_prefix_from_page_path,
    infer_result_area_prefix_from_table_path as infer_result_area_prefix_from_table_path,
    join_context_path as join_context_path,
    resolve_receipt_pagination_paths as resolve_receipt_pagination_paths,
    resolve_receipt_pagination_paths_dynamic as resolve_receipt_pagination_paths_dynamic,
    split_context_path as split_context_path,
    strip_context_path_suffix as strip_context_path_suffix,
    validate_context_path as validate_context_path,
    validate_receipt_pagination_path_report as validate_receipt_pagination_path_report,
    with_runtime_pagination_paths as with_runtime_pagination_paths,
)
from tools.receipt_query_reader import (  # noqa: E402
    dedupe_page_tables as dedupe_page_tables,
    evaluate_paging_match_stop as evaluate_paging_match_stop,
    read_receipt_result_pages as read_receipt_result_pages,
    read_receipt_result_pages_incremental as read_receipt_result_pages_incremental,
    read_receipt_result_pages_until_match as read_receipt_result_pages_until_match,
    unresolved_excel_rows as unresolved_excel_rows,
)
from tools.receipt_query_result_tables import (  # noqa: E402
    enumerate_receipt_result_table_paths as enumerate_receipt_result_table_paths,
    find_table_paths_in_context as find_table_paths_in_context,
    first_non_empty_cell as first_non_empty_cell,
    first_non_empty_cell_at as first_non_empty_cell_at,
    is_receipt_result_table_candidate as is_receipt_result_table_candidate,
    read_receipt_result_table_by_path as read_receipt_result_table_by_path,
    read_receipt_result_tables_runtime as read_receipt_result_tables_runtime,
    read_receipt_tables as read_receipt_tables,
    receipt_result_read_columns as receipt_result_read_columns,
    summarize_receipt_tables as summarize_receipt_tables,
)
from tools.receipt_query_report import (  # noqa: E402
    build_dry_run_match_report as build_dry_run_match_report,
    build_dry_run_match_report_from_preview as build_dry_run_match_report_from_preview,
    build_receipt_write_back_report as build_receipt_write_back_report,
    diagnose_match_inputs as diagnose_match_inputs,
    summarize_nc_rows as summarize_nc_rows,
    unique_ordered as unique_ordered,
)


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
                timings.add("result_wait_before_read", time.perf_counter() - wait_start)
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


def main():
    from tools.receipt_query_cli import main as cli_main

    return cli_main()


if __name__ == "__main__":
    raise SystemExit(main())
