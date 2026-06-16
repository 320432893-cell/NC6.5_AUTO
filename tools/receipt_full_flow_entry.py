# 职责：编排收款单完整流程测试入口，消费 ReceiptPlanRow 跑开单/表头/明细/手续费/可选保存
# 不做什么：不做录入前 NC 查重，不复用历史 T0 保存脚本，不把保存设为默认行为
# 允许依赖层：core 收款计划/配置/JAB、tools 下已正式化的收款开单/明细/查询组件
# 谁不应该 import：core 层模块不应 import 本入口；凭证批量模块不应 import

import argparse
from dataclasses import asdict
from decimal import Decimal
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_entry import ReceiptEntryWorkbook  # noqa: E402
from core.receipt_models import ReceiptBatchResultRow  # noqa: E402
from core.run_state import RunStateRecorder  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table_cached  # noqa: E402
from tools.receipt_detail_async_verifier import DetailPipelineVerifier  # noqa: E402
from tools.receipt_detail_row_cleanup import delete_extra_row_if_present  # noqa: E402
from tools.receipt_detail_rows import StepTimer, run_fee_only  # noqa: E402
from tools.receipt_detail_writer import write_detail_line_by_screen  # noqa: E402
from tools.receipt_modal_guard import recover_cancelable_modal_before_save  # noqa: E402
from tools.receipt_new_probe import (  # noqa: E402
    collect_receipt_new_windows,
    detect_self_made_entry_state,
)
from tools.receipt_post_save_query import run_post_save_batch_query  # noqa: E402
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    detect_existing_self_made_entry,
    fill_header,
    locate_receipt_header_scope,
    read_body_table,
    run_receipt_new_probe,
    wait_header_account_description,
)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "收款单完整流程测试入口：消费 ReceiptPlanRow，默认开单/写表头/写明细/手续费后"
            "停在保存前；显式 --save 才会保存。"
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--excel-path", default=None)
    parser.add_argument(
        "--excel-row", type=int, default=None, help="只测试指定 Sheet1 行"
    )
    parser.add_argument(
        "--excel-rows",
        default=None,
        help="只测试指定 Sheet1 行号，逗号分隔；例如 1801,1802,1803",
    )
    parser.add_argument("--limit", type=int, default=1, help="最多测试几行，默认 1")
    parser.add_argument(
        "--write-plan-sheet",
        action="store_true",
        help="运行前先把本地预检结果写入 Sheet2",
    )
    parser.add_argument(
        "--write-selected-plan-sheet",
        action="store_true",
        help="运行前只把本次选中的计划行写入 Sheet2",
    )
    parser.add_argument(
        "--validation-mode",
        choices=("strict", "skip_invalid_rows"),
        default=None,
        help="覆盖 receipt_entry.validation_policy.mode",
    )
    parser.add_argument(
        "--save",
        action="store_true",
        help="高风险：发送 Ctrl+S 保存收款单。默认不保存。",
    )
    parser.add_argument(
        "--yes-i-understand",
        action="store_true",
        help="配合 --save 跳过交互确认；仍会输出高风险提示。",
    )
    parser.add_argument(
        "--query-after-save",
        action="store_true",
        help="保存后按主体/日期区间查询 NC。当前仅在 --save 后允许。",
    )
    parser.add_argument(
        "--start-delay",
        type=float,
        default=2.0,
        help="启动真实 NC 动作前等待秒数，默认 2。",
    )
    parser.add_argument("--json", action="store_true", help="只输出 JSON 报告")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    config = load_config(args.config)
    if args.validation_mode:
        policy = config.setdefault("receipt_entry", {}).setdefault(
            "validation_policy", {}
        )
        policy["mode"] = args.validation_mode
        policy["skip_invalid_rows"] = args.validation_mode == "skip_invalid_rows"
    if args.save:
        confirm_save(args)
    if args.query_after_save and not args.save:
        raise SystemExit("--query-after-save 当前只允许配合 --save 使用")
    # CLI 参数校验在创建记录器之前完成,避免早退把 run_state 留在 running 态
    if args.write_plan_sheet and args.write_selected_plan_sheet:
        raise SystemExit("--write-plan-sheet 和 --write-selected-plan-sheet 只能选一个")

    recorder = RunStateRecorder(command="receipt-full-flow", config=config)
    recorder.set_stage("预检计划")

    report = {
        "launcher": "receipt_full_flow_entry.py",
        "mode": "save" if args.save else "no-save",
        "query_after_save": bool(args.query_after_save),
        "started_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "rows": [],
    }
    started = time.perf_counter()
    workbook = ReceiptEntryWorkbook(config, excel_path=args.excel_path)
    plan_rows, issues, summary = workbook.build_local_plan(write_sheet=False)
    selected_rows = select_plan_rows(plan_rows, issues, args)
    report["local_plan"] = {
        "summary": summary,
        "issue_count": len(issues),
        "selected_rows": [row.row for row in selected_rows],
        "write_plan_sheet": bool(args.write_plan_sheet),
        "write_selected_plan_sheet": bool(args.write_selected_plan_sheet),
    }
    if not selected_rows:
        report.update(
            {
                "ok": False,
                "reason": "没有可测试的通过预检计划行",
                "total_seconds": round(time.perf_counter() - started, 3),
            }
        )
        recorder.event("no-rows", reason="没有可测试的通过预检计划行")
        recorder.finish("failed", error="没有可测试的通过预检计划行")
        print_report(report, args)
        return 2
    recorder.update_counts(total=len(selected_rows), succeeded=0, failed=0, skipped=0)
    if args.write_plan_sheet:
        workbook.write_plan_sheet(plan_rows, issues)
    elif args.write_selected_plan_sheet:
        workbook.write_plan_sheet(
            selected_rows,
            filter_issues_for_rows(issues, selected_rows),
        )

    if not args.json:
        print("收款单完整流程测试入口")
        print(f"模式：{'真实保存' if args.save else '不保存，停在保存前'}")
        print(f"计划行：{[row.row for row in selected_rows]}")
        print(f"启动后等待 {args.start_delay:g} 秒，请切到 NC 收款单录入页面。")
    time.sleep(max(args.start_delay, 0))

    exit_code = 0
    _succeeded = 0
    _failed = 0
    try:
        for _step_idx, row in enumerate(selected_rows):
            recorder.set_stage(
                "处理行",
                step_index=_step_idx + 1,
                total_steps=len(selected_rows),
                excel_row=row.row,
            )
            row_report = run_one_row(
                config, row, save_enabled=args.save, recorder=recorder
            )
            report["rows"].append(row_report)
            if row_report.get("ok"):
                _succeeded += 1
                recorder.update_counts(
                    total=len(selected_rows),
                    succeeded=_succeeded,
                    failed=_failed,
                    skipped=0,
                )
                recorder.event("row-done", excel_row=row.row, outcome="success")
            else:
                _failed += 1
                recorder.update_counts(
                    total=len(selected_rows),
                    succeeded=_succeeded,
                    failed=_failed,
                    skipped=0,
                )
                recorder.event(
                    "row-done",
                    excel_row=row.row,
                    outcome="failed",
                    error=row_report.get("reason") or row_report.get("failed_step"),
                )
                exit_code = 1
                break
        batch_results = build_batch_results(selected_rows, report["rows"])
        if args.query_after_save and report["rows"] and exit_code == 0:
            recorder.set_stage("后验查询")
            batch_results, post_query = run_post_save_batch_query(
                config,
                selected_rows,
                report["rows"],
            )
            report["post_query"] = post_query
        if args.write_selected_plan_sheet:
            workbook.write_batch_result_sheet(batch_results)

        report["ok"] = exit_code == 0
        report["total_seconds"] = round(time.perf_counter() - started, 3)
        recorder.finish("success" if exit_code == 0 else "failed")
        print_report(report, args)
        return exit_code
    except Exception as _exc:
        recorder.finish("failed", error=str(_exc))
        raise


def confirm_save(args):
    print("高风险：--save 会发送 Ctrl+S，真实保存 NC 收款单。")
    print("保存前请确认：当前账号权限、Excel 行、NC 页面、测试单据清理方案均已确认。")
    if args.yes_i_understand:
        return
    answer = input("确认保存请输入 SAVE: ").strip()
    if answer != "SAVE":
        raise SystemExit("用户取消保存")


def select_plan_rows(plan_rows, issues, args):
    issue_rows = {issue.excel_row for issue in issues if issue.excel_row is not None}
    runnable = [row for row in plan_rows if row.row not in issue_rows]
    target_rows = parse_excel_rows_arg(getattr(args, "excel_rows", None))
    if args.excel_row is not None and target_rows:
        raise SystemExit("--excel-row 和 --excel-rows 只能选一个")
    if target_rows:
        target_set = set(target_rows)
        runnable = [row for row in runnable if row.row in target_set]
        by_number = {row.row: row for row in runnable}
        runnable = [
            by_number[row_number]
            for row_number in target_rows
            if row_number in by_number
        ]
    elif args.excel_row is not None:
        runnable = [row for row in runnable if row.row == args.excel_row]
    limit = max(int(args.limit or 0), 0)
    if limit:
        runnable = runnable[:limit]
    return runnable


def parse_excel_rows_arg(value):
    if value in (None, ""):
        return []
    rows = []
    for part in str(value).split(","):
        text = part.strip()
        if not text:
            continue
        try:
            row_number = int(text)
        except ValueError as exc:
            raise SystemExit(f"--excel-rows 行号无效: {text!r}") from exc
        if row_number <= 0:
            raise SystemExit(f"--excel-rows 行号必须大于 0: {row_number}")
        rows.append(row_number)
    return list(dict.fromkeys(rows))


def filter_issues_for_rows(issues, selected_rows):
    selected = {row.row for row in selected_rows}
    return [
        issue
        for issue in issues
        if issue.excel_row is None or issue.excel_row in selected
    ]


def run_one_row(config, row, save_enabled=False, recorder=None):
    def _stage(stage, **fields):
        if recorder is not None:
            recorder.set_stage(stage, **fields)

    def _event(name, **fields):
        if recorder is not None:
            recorder.event(name, **fields)

    timings = StepTimer()
    flow_started_at = time.perf_counter()
    row_report = {
        "excel_row": row.row,
        "plan_row": serializable(asdict(row)),
        "steps": [],
        "save_enabled": bool(save_enabled),
    }
    business = business_from_plan_row(row)
    _stage("开单", excel_row=row.row)
    open_step = timings.measure("open.self-made", open_self_made_entry, config)
    row_report["steps"].append({"name": "open-self-made", **open_step})
    if not open_step.get("ok"):
        _event("open-failed", excel_row=row.row, error=open_step.get("reason"))
        return fail(row_report, "open-self-made", timings, open_step.get("reason"))

    jab = JABOperator(config)
    pipeline_verifier = None
    try:
        timings.measure("jab.ensure-started", jab.ensure_started)
        health = timings.measure("jab.health-check", check_jab_ready, jab)
        row_report["jab_health"] = health
        if not health.get("ok"):
            _event("jab-health-failed", excel_row=row.row, error=health.get("reason"))
            return fail(row_report, "jab-health-check", timings, health.get("reason"))
        _stage("表头", excel_row=row.row)
        header_steps = timings.measure(
            "header.fill",
            fill_header,
            jab,
            business,
        )
        row_report["header_steps"] = header_steps
        row_report["nc_customer_name"] = extract_header_accepted_text(
            header_steps,
            "客户",
        )
        if any(not step.get("ok") for step in header_steps):
            _event("header-fill-failed", excel_row=row.row, error="表头字段写入失败")
            return fail(row_report, "header-fill", timings, "表头字段写入失败")
        if not row_report["nc_customer_name"]:
            _event("header-fill-failed", excel_row=row.row, error="客户名称未读回")
            return fail(row_report, "header-fill", timings, "客户名称未读回")
        _stage("明细主行", excel_row=row.row)
        located = timings.measure(
            "body.locate",
            locate_receipt_body_table_cached,
            jab,
            max_rows=5,
        )
        row_report["table_candidates"] = located.get("candidates", [])[:5]
        if not located.get("best"):
            _event("locate-table-failed", excel_row=row.row, error="未定位到明细表")
            return fail(row_report, "locate-body-table", timings, "未定位到明细表")
        row_report["before_table"] = timings.measure(
            "body.read-before", read_body_table, jab, "before_detail_fill"
        )
        pipeline_verifier = DetailPipelineVerifier(
            config,
            located,
            flow_started_at=flow_started_at,
        )
        pipeline_verifier.start()
        pipeline_field_task_ids = []
        pipeline_snapshot_task_ids = []
        pipeline_row_count_task_id = None

        def submit_detail_verify(row_index, field, business_values, _step):
            task_id = pipeline_verifier.submit_field(
                row_index,
                field,
                business_values,
            )
            pipeline_field_task_ids.append(task_id)
            return task_id

        detail_steps = timings.measure(
            "detail.main-line",
            write_detail_line_by_screen,
            jab,
            business,
            located,
            after_field=submit_detail_verify,
        )
        row_report["detail_steps"] = detail_steps
        pipeline_snapshot_task_ids.append(
            pipeline_verifier.submit_snapshot(
                "after-main-line",
                max_rows=3,
                min_matches=len(detail_steps),
            )
        )
        if not all(step.get("ok") for step in detail_steps):
            _event("detail-main-failed", excel_row=row.row, error="明细主行写入失败")
            return fail(row_report, "detail-main-line", timings, "明细主行写入失败")
        if row.fee > 0:
            _stage("手续费", excel_row=row.row)
            row_report["extra_row_delete"] = {
                "ok": True,
                "skipped": True,
                "reason": "手续费非 0，保留主行后自动带出的第 2 行给手续费覆盖",
            }
            add_row, fee_steps, clear_account, delete_extra = timings.measure(
                "detail.fee-line",
                run_fee_only,
                jab,
                located,
                str(row.fee),
                after_field=submit_detail_verify,
            )
            row_report["fee_row_add"] = add_row
            row_report["fee_steps"] = fee_steps
            pipeline_snapshot_task_ids.append(
                pipeline_verifier.submit_snapshot(
                    "after-fee-line",
                    max_rows=4,
                )
            )
            row_report["fee_account_clear"] = clear_account
            row_report["fee_extra_row_delete"] = delete_extra
            if delete_extra.get("ok"):
                pipeline_row_count_task_id = pipeline_verifier.submit_row_count(2)
            if (
                not add_row.get("ok")
                or not all(step.get("ok") for step in fee_steps)
                or not clear_account.get("ok")
                or not delete_extra.get("ok")
            ):
                _event("fee-line-failed", excel_row=row.row, error="手续费行处理失败")
                return fail(row_report, "detail-fee-line", timings, "手续费行处理失败")
        else:
            row_report["extra_row_delete"] = timings.measure(
                "detail.delete-extra-after-main",
                delete_extra_row_if_present,
                jab,
                located,
                1,
            )
            if not row_report["extra_row_delete"].get("ok"):
                _event(
                    "delete-extra-failed",
                    excel_row=row.row,
                    error="主行后多余行删除失败",
                )
                return fail(
                    row_report,
                    "detail-delete-extra-after-main",
                    timings,
                    "主行后多余行删除失败",
                )
            row_report["fee_skipped"] = {
                "ok": True,
                "reason": "手续费为 0，跳过手续费行",
            }
            pipeline_row_count_task_id = pipeline_verifier.submit_row_count(1)
        pipeline_wait_ids = []
        if pipeline_field_task_ids:
            pipeline_wait_ids.append(pipeline_field_task_ids[-1])
        if pipeline_row_count_task_id:
            pipeline_wait_ids.append(pipeline_row_count_task_id)
        pipeline_wait_started = time.perf_counter()
        row_report["detail_pipeline_verify"] = pipeline_verifier.wait(
            pipeline_wait_ids,
            timeout=2.0,
        )
        timings.add(
            "detail.pipeline-final-wait",
            time.perf_counter() - pipeline_wait_started,
        )
        row_report["detail_pipeline_snapshots"] = pipeline_snapshot_task_ids
        if not row_report["detail_pipeline_verify"].get("ok"):
            row_report["after_table"] = timings.measure(
                "body.read-after-fallback", read_body_table, jab, "after_detail_fill"
            )
            _event(
                "pipeline-verify-failed",
                excel_row=row.row,
                error="后台明细验证未通过，已执行整表读 fallback",
            )
            return fail(
                row_report,
                "detail-pipeline-verify",
                timings,
                "后台明细验证未通过，已执行整表读 fallback",
            )
        row_report["after_table"] = {
            "ok": True,
            "skipped": True,
            "reason": "后台 pipeline verifier 已覆盖最后字段与最终行数，跳过同步整表读",
        }
        account_check = timings.measure(
            "header.account-readback-after-detail",
            wait_header_account_description,
            jab,
            1.0,
        )
        row_report["header_account"] = account_check
        if not account_check.get("accepted"):
            _event(
                "account-readback-failed",
                excel_row=row.row,
                error="收款银行账户为空",
            )
            return fail(
                row_report,
                "header-account-readback-after-detail",
                timings,
                "收款银行账户为空",
            )
        _stage("保存前守卫", excel_row=row.row)
        row_report["pre_save_modal_recovery"] = timings.measure(
            "guard.pre-save-modal-recovery",
            recover_cancelable_modal_before_save,
            jab,
            probe_receipt_entry_page,
        )
        if not row_report["pre_save_modal_recovery"].get("ok"):
            _event(
                "pre-save-guard-failed",
                excel_row=row.row,
                error=row_report["pre_save_modal_recovery"].get("reason"),
            )
            return fail(
                row_report,
                "pre-save-modal-recovery",
                timings,
                row_report["pre_save_modal_recovery"].get("reason"),
            )
        if save_enabled:
            _stage("保存", excel_row=row.row)
            save_result = timings.measure("save.ctrl-s", save_receipt_by_ctrl_s, jab)
            row_report["save"] = save_result
            if not save_result.get("ok"):
                _event(
                    "save-failed", excel_row=row.row, error=save_result.get("reason")
                )
                return fail(row_report, "save", timings, save_result.get("reason"))
        else:
            row_report["save"] = {
                "ok": True,
                "skipped": True,
                "reason": "no-save 模式：已停在保存前，未发送 Ctrl+S",
            }
        row_report["ok"] = True
        row_report["timings"] = timings.items
        return row_report
    finally:
        if pipeline_verifier is not None:
            pipeline_verifier.close(timeout=0.2)
        jab.close()


def open_self_made_entry(config):
    existing = detect_existing_self_made_entry(config)
    if existing.get("ok"):
        return existing
    opened = run_receipt_new_probe()
    if opened.get("ok"):
        return opened
    opened["reason"] = opened.get("reason") or "未能进入收款单自制录入态"
    return opened


def probe_receipt_entry_page(jab):
    scope = locate_receipt_header_scope(jab)
    if scope.get("ok"):
        return {"ok": True, "method": "header-scope", "scope": scope}
    windows = collect_receipt_new_windows(jab)
    state = detect_self_made_entry_state(windows)
    return {
        "ok": bool(state.get("ok")),
        "method": "entry-state",
        "scope": scope,
        "entry_state": state,
    }


def business_from_plan_row(row):
    return {
        "finance_org_code": row.organization_code,
        "finance_org_name": row.organization_name,
        "document_date": row.receipt_date.isoformat(),
        "customer_code": row.customer_code,
        "currency": row.currency,
        "header_currency_code": row.header_currency_code or row.currency,
        "bank_label": row.bank,
        "bank_account": row.account_no,
        "amount": str(row.raw_amount),
        "fee": str(row.fee),
        "has_fee": row.fee > 0,
        "settlement": "网银",
        "main_subject": "1002",
        "main_business_type": "货款",
        "fee_subject": "660305",
        "fee_business_type": "手续费",
    }


def extract_header_accepted_text(header_steps, label):
    for step in header_steps or []:
        if step.get("label") != label:
            continue
        text = str(step.get("accepted_text") or "").strip()
        if text:
            return text
        backend = step.get("backend_state") or {}
        for key in ("description", "text", "name"):
            value = str(backend.get(key) or "").strip()
            if value and value != str(step.get("value") or "").strip():
                return value
    return ""


def build_batch_results(selected_rows, row_reports):
    reports_by_row = {int(report.get("excel_row")): report for report in row_reports}
    results = []
    for row in selected_rows:
        report = reports_by_row.get(row.row) or {}
        ok = bool(report.get("ok"))
        reason = "" if ok else format_row_failure_reason(report)
        results.append(
            ReceiptBatchResultRow(
                plan_row=row,
                local_status="通过" if ok else "异常",
                exception_reason=reason,
                nc_customer_name=str(report.get("nc_customer_name") or "").strip(),
                nc_document_no=str(report.get("nc_document_no") or "").strip(),
            )
        )
    return results


def format_row_failure_reason(report):
    failed_step = str(report.get("failed_step") or "").strip()
    reason = str(report.get("reason") or "").strip()
    if failed_step.startswith("save"):
        return f"保存失败-{reason or failed_step}"
    if failed_step:
        return (
            f"录入失败-{failed_step}:{reason}" if reason else f"录入失败-{failed_step}"
        )
    return reason or "录入失败"


def save_receipt_by_ctrl_s(jab, timeout=12.0):
    clicked = jab.click_save(timeout=3.0)
    prompt_seen = jab.wait_save_success(timeout=2.0)
    started = time.perf_counter()
    last_state = None
    while time.perf_counter() - started < timeout:
        windows = collect_receipt_new_windows(jab)
        state = detect_self_made_entry_state(windows)
        last_state = state
        if not state.get("ok"):
            return {
                "ok": True,
                "clicked": bool(clicked),
                "prompt_seen": bool(prompt_seen),
                "seconds": round(time.perf_counter() - started, 3),
                "oracle": "自制录入态按钮消失",
                "entry_state": state,
            }
        time.sleep(0.3)
    return {
        "ok": False,
        "clicked": bool(clicked),
        "prompt_seen": bool(prompt_seen),
        "seconds": round(time.perf_counter() - started, 3),
        "reason": "保存后仍检测到自制录入态按钮，未确认回到新增态",
        "entry_state": last_state,
    }


def fail(row_report, failed_step, timings, reason):
    row_report.update(
        {
            "ok": False,
            "failed_step": failed_step,
            "reason": reason,
            "timings": timings.items,
        }
    )
    return row_report


def serializable(value):
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, dict):
        return {key: serializable(item) for key, item in value.items()}
    if isinstance(value, list | tuple):
        return [serializable(item) for item in value]
    return value


def print_report(report, args):
    text = json.dumps(report, ensure_ascii=False, indent=2, default=str)
    print(text)
    if args.json:
        return
    print()
    if report.get("ok"):
        print("成功：完整流程测试通过。")
    else:
        print(f"失败：{report.get('reason') or '至少一行完整流程测试失败'}")
    print(f"总用时：{report.get('total_seconds')}s")


if __name__ == "__main__":
    raise SystemExit(main())
