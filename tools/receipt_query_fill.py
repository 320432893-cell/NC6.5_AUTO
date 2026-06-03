import argparse
from collections import Counter
from copy import deepcopy
from datetime import date
import json
from pathlib import Path
import re
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
    format_receipt_duplicate_reason,
    names_match,
    parse_date,
)
from core.utils import load_config  # noqa: E402


def resolve_today(value):
    return date.today().isoformat() if value == "{today}" else value


def set_text(jab, jab_cfg, path, value):
    return jab.set_text_by_path(
        path,
        value,
        title=jab_cfg["dialog_title"],
        class_name=jab_cfg["dialog_class"],
        role="text",
        timeout=2,
        require_showing=True,
    )


def ensure_query_window(jab, config, query_cfg, jab_cfg, skip_open=False):
    title = jab_cfg["dialog_title"]
    class_name = jab_cfg["dialog_class"]
    timeout = float(query_cfg.get("open_timeout", query_cfg.get("timeout", 5)))
    existing = jab.wait_window_by_title(
        title,
        class_name=class_name,
        timeout=0.5,
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
        wait=float(query_cfg.get("open_wait", 0.8)),
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
):
    query_cfg = config["receipt_entry"]["query"]
    jab_cfg = query_cfg["jab"]
    fields = jab_cfg["fields"]
    start = date_from or query_cfg["date_from"]
    end = date_to or query_cfg["date_to"]
    start = parse_date(resolve_today(start)).isoformat()
    end = parse_date(resolve_today(end)).isoformat()

    jab = JABOperator(config)
    try:
        if not ensure_query_window(
            jab,
            config,
            query_cfg,
            jab_cfg,
            skip_open=skip_open_query,
        ):
            raise RuntimeError("未检测到收款单查询条件窗口")

        steps = [
            (
                "finance_org",
                fields["finance_org"]["text_path"],
                org_code,
            ),
            (
                "document_date_from",
                fields["document_date"]["from_text_path"],
                start,
            ),
            (
                "document_date_to",
                fields["document_date"]["to_text_path"],
                end,
            ),
        ]
        for name, path, value in steps:
            if not set_text(jab, jab_cfg, path, value):
                raise RuntimeError(f"收款查询条件写入失败: {name}={value}")

        if confirm:
            ok = jab.do_action_by_path(
                jab_cfg["confirm_button_path"],
                title=jab_cfg["dialog_title"],
                class_name=jab_cfg["dialog_class"],
                name="确定(Y)",
                role="push button",
                action_name="单击",
                wait=1,
                timeout=2,
                require_showing=True,
            )
            if not ok:
                raise RuntimeError("收款查询确定按钮点击失败")

        result = {"organization_code": org_code, "date_from": start, "date_to": end}
        if set_page_size_only:
            if confirm:
                time.sleep(float(query_cfg.get("result_wait", 1.0)))
            result["page_report"] = set_receipt_page_size(jab, query_cfg)
        if read_results or dry_run_match:
            if confirm:
                time.sleep(float(query_cfg.get("result_wait", 1.0)))
            tables, page_report = read_receipt_result_pages(
                jab,
                query_cfg,
                max_rows=max_rows,
                max_cols=max_cols,
                read_columns=receipt_result_read_columns(
                    query_cfg,
                    include_amount_candidates=dry_run_match,
                ),
            )
            result["table_summary"] = [
                {
                    "table_index": table.get("table_index"),
                    "row_count": table.get("row_count"),
                    "col_count": table.get("col_count"),
                }
                for table in tables
            ]
            result["page_report"] = page_report
            extractor = ReceiptNCResultExtractor(config)
            if read_results:
                rows, issues = extractor.extract(tables)
                result["nc_rows"] = rows
                result["extract_issues"] = issues
            if dry_run_match:
                result["dry_run_match"] = build_dry_run_match_report(
                    config,
                    extractor,
                    tables,
                    org_code,
                    business_date=parse_date(end),
                )
        return result
    finally:
        jab.close()


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
            max_rows=max_rows,
            max_cols=max_cols,
            read_columns=read_columns,
        )
        return tables, {"enabled": False, "pages": []}

    setup_report = set_receipt_page_size(jab, query_cfg)
    window_class = setup_report["window_class"]
    page_label_path = setup_report["page_label_path"]
    next_page_path = pagination["next_page_button_path"]
    pager_hwnd = setup_report["pager_hwnd"]
    after_label = setup_report["after_label"]
    before_label = setup_report["before_label"]

    page_info = parse_page_label(after_label or before_label or "")
    total_pages = page_info.get("total_pages") or 1
    total_records = page_info.get("total_records")
    page_limit = min(total_pages, int(pagination.get("max_pages", total_pages)))

    collected = []
    page_reports = []
    seen_documents = set()
    current_stability = setup_report.get("after_stability")
    for page_number in range(1, page_limit + 1):
        label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
        time.sleep(float(pagination.get("wait_before_read", 0.0)))
        tables = read_receipt_tables(
            jab,
            max_rows=max_rows,
            max_cols=max_cols,
            read_columns=read_columns,
        )
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
                    }
                    for table in tables
                ],
                "stability": current_stability,
            }
        )
        time.sleep(float(pagination.get("wait_after_page_read", 0.0)))
        if page_number >= page_limit:
            break
        if not pager_hwnd or not label:
            page_reports[-1]["next_page_ok"] = False
            page_reports[-1]["next_page_method"] = "blocked_no_pager_scope"
            break
        ok, method = click_next_page(
            jab,
            pagination,
            next_page_path,
            window_class,
            pager_hwnd,
        )
        if not ok:
            page_reports[-1]["next_page_ok"] = False
            page_reports[-1]["next_page_method"] = method
            break
        page_reports[-1]["next_page_ok"] = True
        page_reports[-1]["next_page_method"] = method
        current_stability = wait_receipt_result_stable(
            jab, query_cfg, pager_hwnd=pager_hwnd
        )
        page_reports[-1]["after_next_stability"] = current_stability

    return collected, {
        "enabled": True,
        **setup_report,
        "total_pages": total_pages,
        "total_records": total_records,
        "pages": page_reports,
    }


def set_receipt_page_size(jab, query_cfg):
    pagination = query_cfg.get("pagination") or {}
    if not pagination:
        return {"enabled": False, "page_size_ok": False}

    page_size = int(pagination.get("page_size", 500))
    window_class = pagination.get("window_class", "SunAwtCanvas")
    page_size_path = pagination["page_size_text_path"]
    page_label_path = pagination["page_label_path"]
    pager_window = jab.wait_context_by_path(
        page_label_path,
        class_name=window_class,
        role="label",
        timeout=float(pagination.get("pager_scope_timeout", 2.0)),
        require_showing=True,
        require_valid_bounds=False,
    )
    pager_hwnd = pager_window.get("hwnd") if pager_window else None

    if bool(pagination.get("wait_before_page_size_stable", True)):
        before_stability = wait_receipt_result_stable(
            jab, query_cfg, pager_hwnd=pager_hwnd
        )
    else:
        before_stability = {
            "ok": None,
            "label": read_page_label(jab, page_label_path, window_class, pager_hwnd),
            "tables": summarize_receipt_tables(jab, query_cfg),
        }
    before_label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
    time.sleep(float(pagination.get("wait_before_page_size", 0.0)))
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
    if page_size_ok:
        jab.press_key("enter", wait=float(pagination.get("wait_after_page_size", 2.0)))
    after_stability = wait_receipt_result_stable(jab, query_cfg, pager_hwnd=pager_hwnd)
    after_label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
    return {
        "enabled": True,
        "page_size": page_size,
        "page_size_ok": bool(page_size_ok),
        "pager_hwnd": pager_hwnd,
        "pager_scope_ok": bool(pager_hwnd),
        "window_class": window_class,
        "page_label_path": page_label_path,
        "before_label": before_label,
        "after_label": after_label,
        "before_stability": before_stability,
        "after_stability": after_stability,
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

    while time.time() < deadline:
        label = read_page_label(jab, page_label_path, window_class, pager_hwnd)
        summary = summarize_receipt_tables(jab, query_cfg)
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
            }
        time.sleep(interval)

    last = samples[-1] if samples else {"label": None, "tables": []}
    return {
        "ok": False,
        "samples": len(samples),
        "label": last.get("label"),
        "tables": last.get("tables"),
    }


def summarize_receipt_tables(jab, query_cfg):
    indexes = receipt_result_read_columns(query_cfg)
    min_cols = max(indexes) + 1 if indexes else None
    if hasattr(jab, "read_table_summaries"):
        return jab.read_table_summaries(min_rows=2, min_cols=min_cols)
    tables = read_receipt_tables(
        jab,
        max_rows=0,
        max_cols=0,
        read_columns=[],
    )
    return [
        {
            "table_index": table.get("table_index"),
            "row_count": table.get("row_count"),
            "col_count": table.get("col_count"),
        }
        for table in tables
        if table.get("row_count", 0) >= 2
        and (min_cols is None or table.get("col_count", 0) >= min_cols)
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


def read_receipt_tables(jab, max_rows=500, max_cols=80, read_columns=None):
    if read_columns and hasattr(jab, "read_all_table_selected_columns"):
        return jab.read_all_table_selected_columns(
            read_columns,
            max_rows=max_rows,
            min_rows=2,
            min_cols=max(read_columns) + 1,
        )
    return jab.read_all_table_cells(max_rows=max_rows, max_cols=max_cols)


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

    ok = jab.do_action_by_path(
        next_page_path,
        class_name=window_class,
        scope_hwnd=scope_hwnd,
        click_mode="bounds",
        timeout=float(pagination.get("next_bounds_timeout", 2.0)),
        wait=wait_after_next,
        require_showing=True,
        require_valid_bounds=False,
    )
    if ok:
        return True, "bounds"
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


def build_dry_run_match_report(config, extractor, tables, org_code, business_date):
    rows, candidates, excel_issues = ReceiptEntryWorkbook(config).preview_rows(
        today=business_date
    )
    org_candidates = [row for row in candidates if row.organization_code == org_code]
    report = {
        "business_date": business_date.isoformat(),
        "excel_rows": len(rows),
        "excel_candidates": len(candidates),
        "org_candidates": len(org_candidates),
        "excel_issues": len(excel_issues),
        "candidate_banks": dict(
            sorted(Counter(row.bank for row in org_candidates).items())
        ),
        "variants": [],
    }
    matcher = ReceiptEntryDryRunMatcher()
    amount_columns = (8, 6, 7)
    for amount_column in amount_columns:
        name = f"confirmed_col2_amount_col{amount_column}"
        column = 2
        nc_rows, extract_issues = extractor.extract_by_indexes(
            tables,
            column,
            amount_column=amount_column,
        )
        matched, match_issues = matcher.match(org_candidates, nc_rows)
        report["variants"].append(
            {
                "name": name,
                "name_column": column,
                "amount_column": amount_column,
                "nc_rows": len(nc_rows),
                "nc_summary": summarize_nc_rows(nc_rows),
                "match_diagnostics": diagnose_match_inputs(org_candidates, nc_rows),
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
    return report


def diagnose_match_inputs(excel_rows, nc_rows):
    nc_by_amount = {}
    for nc_row in nc_rows:
        nc_by_amount.setdefault(nc_row.original_amount, []).append(nc_row)

    amount_only_hits = 0
    name_amount_hits = 0
    duplicate_hits = 0
    name_amount_samples = []
    name_mismatch_samples = []
    no_amount_samples = []
    for excel_row in excel_rows:
        amount_candidates = nc_by_amount.get(excel_row.raw_amount, [])
        if not amount_candidates:
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
                    "reason": "金额命中但名称不符",
                    "nc_names": [row.name for row in amount_candidates[:5]],
                    "nc_rows": [row.row_index for row in amount_candidates[:5]],
                }
            )
    return {
        "amount_only_hits": amount_only_hits,
        "name_amount_hits": name_amount_hits,
        "duplicate_hits": duplicate_hits,
        "name_amount_samples": name_amount_samples,
        "name_mismatch_samples": name_mismatch_samples,
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
    )
    print(
        "filled receipt query: "
        f"org={result['organization_code']} "
        f"date_from={result['date_from']} date_to={result['date_to']} "
        f"confirm={confirm}"
    )
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
            print(
                "  nc_row="
                f"{row.row_index} date={row.document_date.isoformat()} "
                f"amount={row.original_amount} customer={row.customer}"
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
