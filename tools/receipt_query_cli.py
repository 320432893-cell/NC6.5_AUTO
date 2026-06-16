# 职责：收款单查询填充工具的 CLI 参数解析和终端输出。
# 不做什么：不实现查询条件填写、不读结果表、不做分页匹配、不直接封装 JAB 底层能力。
# 允许依赖层：tools.receipt_query_fill 的流程函数、core.utils 配置加载。
# 谁不应该 import：core 层模块不应 import；测试应优先测具体流程/读表模块而非 CLI 输出。

import argparse
from copy import deepcopy
import json
import os
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.run_state import RunStateRecorder  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_query_fill import (  # noqa: E402
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
    parser.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="在 stdout 最后一行输出统一结果信封（GUI 解析用）。",
    )
    args = parser.parse_args()
    _started_at = time.perf_counter()

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

    recorder = RunStateRecorder(command="receipt-query", config=config)
    _exit_code: list[int] = []  # 用列表做可变容器,在 finally 中读取退出码
    try:
        recorder.set_stage(
            "打开查询",
            step_index=1,
            total_steps=1
            + int(confirm)
            + int(read_results or dry_run_match)
            + int(dry_run_match and args.write_back),
        )
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
            if args.json_output:
                _print_query_envelope(
                    ok=False,
                    exit_code=2,
                    result=None,
                    error_category="environment",
                    error_message=str(exc),
                    started_at=_started_at,
                )
            _exit_code.append(2)
            recorder.event("page-guard-failed", error=str(exc))
            return 2
        recorder.set_stage("填查询条件", step_index=2)
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
        page_report = result.get("page_report") or {}
        if page_report:
            print("receipt query page report:")
            for key in (
                "setup_seconds",
                "page_size_changed",
                "before_page_size_text",
                "after_page_size_text",
                "pager_resolution",
                "result_page_resolution",
                "dynamic_resolution",
                "total_records",
                "planned_pages",
            ):
                if key in page_report:
                    print(f"  {key}: {page_report.get(key)}")
            for page in (page_report.get("pages") or [])[:3]:
                print(
                    "  page="
                    f"{page.get('page')} rows="
                    f"{sum((table.get('row_count') or 0) for table in page.get('tables') or [])} "
                    f"read_tables_seconds={page.get('read_tables_seconds')} "
                    f"wait_before_read_seconds={page.get('wait_before_read_seconds')} "
                    f"wait_after_page_read_seconds={page.get('wait_after_page_read_seconds')}"
                )
        if confirm:
            recorder.set_stage("确认查询")
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
            recorder.set_stage("读结果")
            rows = result["nc_rows"]
            issues = result["extract_issues"]
            recorder.update_counts(nc_rows=len(rows), extract_issues=len(issues))
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
            recorder.set_stage("匹配写回")
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
        if args.json_output:
            _print_query_envelope(
                ok=True,
                exit_code=0,
                result=result,
                error_category="none",
                error_message="",
                started_at=_started_at,
            )
        _exit_code.append(0)
    except Exception:
        recorder.finish("failed", error="未捕获异常")
        raise
    finally:
        if _exit_code:
            recorder.finish("success" if _exit_code[0] == 0 else "failed")


def _print_query_envelope(
    *,
    ok: bool,
    exit_code: int,
    result: "dict | None",
    error_category: str,
    error_message: str,
    started_at: float,
) -> None:
    """在 stdout 最后一行打印统一结果信封（§1.2 契约）。"""
    items: list[dict] = []
    if result is not None:
        nc_rows = result.get("nc_rows") or []
        for row in nc_rows:
            items.append(
                {
                    "ref": str(getattr(row, "row_index", "")),
                    "outcome": "success",
                    "reason": "",
                }
            )
        for issue in result.get("extract_issues") or []:
            items.append(
                {
                    "ref": f"table={getattr(issue, 'table_index', '')} row={getattr(issue, 'row_index', '')}",
                    "outcome": "failed",
                    "reason": getattr(issue, "reason", ""),
                }
            )
    succeeded = sum(1 for it in items if it["outcome"] == "success")
    failed = sum(1 for it in items if it["outcome"] == "failed")
    elapsed = round(time.perf_counter() - started_at, 3)
    envelope = {
        "ok": ok,
        "command": "receipt-query",
        "exit_code": exit_code,
        "summary": {
            "total": len(items),
            "succeeded": succeeded,
            "failed": failed,
            "skipped": 0,
        },
        "items": items,
        "error": {"category": error_category, "message": error_message},
        "elapsed_s": elapsed,
        "resumable": {"can_resume": False, "resume_command": None},
    }
    print(json.dumps(envelope, ensure_ascii=True))


if __name__ == "__main__":
    raise SystemExit(main())
