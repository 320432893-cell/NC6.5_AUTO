# 职责：收款单查询填充工具的 CLI 参数解析和终端输出。
# 不做什么：不实现查询条件填写、不读结果表、不做分页匹配、不直接封装 JAB 底层能力。
# 允许依赖层：tools.receipt_query_fill 的流程函数、core.utils 配置加载。
# 谁不应该 import：core 层模块不应 import；测试应优先测具体流程/读表模块而非 CLI 输出。

import argparse
from copy import deepcopy
import json
import os
import sys

from core.utils import load_config
from tools.receipt_query_fill import (
    ReceiptPageGuardError,
    fill_receipt_query as run_fill_receipt_query,
)


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
        result = run_fill_receipt_query(
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
