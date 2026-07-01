# 职责：基于 Sheet2 当前状态补做收款单后验查询，并把匹配结果合并回 Sheet2。
# 不做什么：不录入、不保存、不打开新增收款单。
# 允许依赖层：core 收款计划、Sheet2 读写、后验查询函数。
# 谁不应该 import：GUI 层不应直接 import，本模块由 tools 入口调用。

from dataclasses import asdict
import time

from core.receipt_models import ReceiptBatchResultRow
from core.receipt_post_save_query import run_post_save_batch_query
from core.receipt_report import (
    post_query_failure_reasons,
    serializable,
    write_last_report,
)


POST_QUERY_ISSUE_PREFIXES = (
    "后验未匹配",
    "后验待确认",
    "查询失败",
    "后验查询失败",
    "NC 查询结果为空",
    "NC查询结果为空",
)


def run_post_query_only(config, workbook, recorder=None):
    started = time.perf_counter()
    report = {
        "launcher": "receipt_full_flow_entry.py",
        "mode": "post-query-only",
        "query_after_save": True,
        "post_query_requested": True,
        "post_query_only": True,
        "started_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "rows": [],
    }
    try:
        if recorder:
            recorder.set_stage("预检计划")
        plan_rows, issues, summary = workbook.build_local_plan(write_sheet=False)
        existing_rows = workbook.read_result_sheet_rows()
        targets = select_post_query_targets(existing_rows, plan_rows)
        report["local_plan"] = {
            "summary": summary,
            "issue_count": len(issues),
            "selected_rows": [row.row for row in targets],
            "write_plan_sheet": False,
            "write_selected_plan_sheet": False,
        }
        report["sheet2_scan"] = summarize_sheet2_scan(existing_rows, targets)
        if not targets:
            report.update(
                {
                    "ok": False,
                    "reason": "Sheet2 没有需要补做后验查询的行",
                    "post_query_executed": False,
                    "post_query_skipped": {
                        "reason": "Sheet2 没有需要补做后验查询的行",
                        "rows": 0,
                        "exit_code": 2,
                    },
                    "total_seconds": round(time.perf_counter() - started, 3),
                }
            )
            if recorder:
                recorder.update_counts(total=0, succeeded=0, failed=0, skipped=0)
                recorder.finish("failed", error=report["reason"])
            write_last_report(report)
            return report, 2

        if recorder:
            recorder.update_counts(total=len(targets), succeeded=0, failed=0, skipped=0)
            recorder.set_stage("后验查询", step_index=0, total_steps=len(targets))
            recorder.event("post-query-only-start", rows=[row.row for row in targets])
        row_reports = build_post_query_row_reports(existing_rows, targets)
        report["rows"] = row_reports
        batch_results, post_query = run_post_save_batch_query(
            config,
            targets,
            row_reports,
        )
        report["post_query"] = post_query
        report["post_query_executed"] = True
        merged_results = merge_sheet2_results(existing_rows, plan_rows, batch_results)
        workbook.write_batch_result_sheet(merged_results)
        failures = post_query_failure_reasons(post_query)
        if failures:
            report["post_query_failed_rows"] = failures
        update_row_reports_from_results(report["rows"], batch_results)
        succeeded = sum(1 for result in batch_results if result.nc_document_no)
        failed = len(batch_results) - succeeded
        if recorder:
            recorder.update_counts(total=len(targets), succeeded=succeeded, failed=failed, skipped=0)
            recorder.event(
                "post-query-only-done",
                ok=not failures,
                matched=succeeded,
                issues=failed,
            )
            recorder.finish("success" if not failures else "failed")
        report["ok"] = not failures
        report["total_seconds"] = round(time.perf_counter() - started, 3)
        write_last_report(report)
        return report, 0 if report["ok"] else 1
    except Exception as exc:
        report.update(
            {
                "ok": False,
                "reason": str(exc),
                "total_seconds": round(time.perf_counter() - started, 3),
            }
        )
        if recorder:
            recorder.finish("failed", error=str(exc))
        write_last_report(report)
        raise


def select_post_query_targets(sheet2_rows, plan_rows):
    plan_by_row = {int(row.row): row for row in plan_rows}
    targets = []
    for item in sheet2_rows or []:
        values = item.get("values") or {}
        source_row = values.get("原Sheet1行号")
        if not isinstance(source_row, int):
            continue
        if not should_post_query(values):
            continue
        plan_row = plan_by_row.get(source_row)
        if plan_row is None:
            continue
        targets.append(plan_row)
    targets.sort(key=lambda row: (row.organization_code, row.receipt_date, row.row))
    return targets


def should_post_query(values):
    local_status = str(values.get("本地预检状态") or "").strip()
    post_status = str(values.get("后验核对状态") or "").strip()
    reason = str(values.get("异常原因") or "").strip()
    if local_status != "通过":
        return False
    if post_status == "后验通过":
        return False
    if not reason:
        return True
    return is_post_query_issue_reason(reason)


def is_post_query_issue_reason(reason):
    text = str(reason or "").strip()
    return any(text.startswith(prefix) for prefix in POST_QUERY_ISSUE_PREFIXES)


def build_post_query_row_reports(sheet2_rows, targets):
    values_by_row = {
        int((item.get("values") or {}).get("原Sheet1行号")): (item.get("values") or {})
        for item in sheet2_rows or []
        if isinstance((item.get("values") or {}).get("原Sheet1行号"), int)
    }
    reports = []
    for row in targets:
        values = values_by_row.get(row.row) or {}
        nc_customer_name = str(values.get("NC客户名称") or "").strip()
        if not nc_customer_name:
            nc_customer_name = str(values.get("🟪银行来款名") or "").strip()
        reports.append(
            {
                "ok": True,
                "excel_row": row.row,
                "plan_row": serializable(asdict(row)),
                "nc_customer_name": nc_customer_name,
                "post_query_only": True,
            }
        )
    return reports


def update_row_reports_from_results(row_reports, batch_results):
    results_by_row = {int(result.plan_row.row): result for result in batch_results}
    for report in row_reports:
        result = results_by_row.get(int(report.get("excel_row") or 0))
        if result is None:
            continue
        report["nc_document_no"] = result.nc_document_no
        if result.exception_reason:
            report["ok"] = False
            report["failed_step"] = "post-query"
            report["reason"] = result.exception_reason


def summarize_sheet2_scan(sheet2_rows, targets):
    data_rows = [
        item
        for item in sheet2_rows or []
        if isinstance((item.get("values") or {}).get("原Sheet1行号"), int)
    ]
    return {
        "data_rows": len(data_rows),
        "selected_rows": [row.row for row in targets],
        "selected_count": len(targets),
    }


def merge_sheet2_results(existing_rows, plan_rows, updates):
    plan_by_row = {int(row.row): row for row in plan_rows}
    updates_by_row = {int(result.plan_row.row): result for result in updates or []}
    results = []
    for item in existing_rows or []:
        values = item.get("values") or {}
        source_row = values.get("原Sheet1行号")
        if not isinstance(source_row, int):
            continue
        plan_row = plan_by_row.get(source_row)
        if plan_row is None:
            continue
        if source_row in updates_by_row:
            results.append(updates_by_row[source_row])
            continue
        results.append(result_from_sheet2_values(plan_row, values))
    return results


def result_from_sheet2_values(plan_row, values):
    return ReceiptBatchResultRow(
        plan_row=plan_row,
        local_status=str(values.get("本地预检状态") or "").strip(),
        exception_reason=str(values.get("异常原因") or "").strip(),
        nc_customer_name=str(values.get("NC客户名称") or "").strip(),
        nc_document_no=str(values.get("NC单据号") or values.get("后验查询结果") or "").strip(),
    )
