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

# guard 三项与 wait_after_query_confirm 经本模块 re-export 供 cli / 旧测试消费;
# 其余为本模块自身使用,按需直取源模块(不再做全量兼容转发)。
from tools.receipt_query_guard import (  # noqa: E402
    ReceiptPageGuardError as ReceiptPageGuardError,
    guard_receipt_parent_page as guard_receipt_parent_page,
    guard_receipt_result_tables as guard_receipt_result_tables,
)
from tools.receipt_query_pagination import (  # noqa: E402
    set_receipt_page_size,
    wait_after_query_confirm as wait_after_query_confirm,
)
from tools.receipt_query_dynamic_fields import (  # noqa: E402
    find_query_condition_scope,
    set_query_dynamic_text,
)
from tools.receipt_query_page_reader import read_receipt_result_pages  # noqa: E402
from tools.receipt_query_match_reader import (  # noqa: E402
    read_receipt_result_pages_until_match,
)
from tools.receipt_query_result_tables import receipt_result_read_columns  # noqa: E402


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


def set_finance_org_text(jab, jab_cfg, value, dynamic_scope=None):
    scope = dynamic_scope or find_query_condition_scope(jab, jab_cfg)
    result = set_query_dynamic_text(jab, jab_cfg, scope, "finance_org", value)
    return bool(result.get("ok"))


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
        interval=float(query_cfg.get("window_poll_interval", 0.05)),
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
    open_key = query_cfg.get("open_key", batch_open_query.get("key", "f3"))
    deadline = time.perf_counter() + timeout
    interval = float(query_cfg.get("window_poll_interval", 0.05))
    press_interval = float(query_cfg.get("open_key_retry_interval", 0.2))
    next_press_at = 0.0
    while time.perf_counter() < deadline:
        now = time.perf_counter()
        if now >= next_press_at:
            jab.press_key(open_key, wait=0.0)
            next_press_at = now + press_interval
        opened = jab.wait_window_by_title(
            title,
            class_name=class_name,
            timeout=interval,
            include_children=bool(query_cfg.get("dialog_include_children", True)),
            visible_only=bool(query_cfg.get("dialog_visible_only", True)),
            interval=interval,
        )
        if opened:
            return True
    return False


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

        query_scope = timings.measure(
            "query.dynamic-scope",
            find_query_condition_scope,
            jab,
            jab_cfg,
        )
        if not query_scope.get("ok"):
            raise RuntimeError(
                f"查询条件动态 path 定位失败:{query_scope.get('reason')}"
            )
        finance_org_ok = timings.measure(
            "set_finance_org",
            set_finance_org_text,
            jab,
            jab_cfg,
            org_code,
            query_scope,
        )
        if not finance_org_ok:
            raise RuntimeError(f"收款查询条件写入失败: finance_org={org_code}")

        steps = [
            ("document_date_from", start),
            ("document_date_to", end),
        ]
        for name, value in steps:
            written = timings.measure(
                name,
                set_query_dynamic_text,
                jab,
                jab_cfg,
                query_scope,
                name,
                value,
            )
            if not written.get("ok"):
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
                        include_amount_candidates=False,
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
