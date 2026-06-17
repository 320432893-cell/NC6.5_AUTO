# 职责：编排收款单完整流程测试入口，消费 ReceiptPlanRow 跑开单/表头/明细/手续费/可选保存
# 不做什么：不做录入前 NC 查重，不复用历史 T0 保存脚本，不把保存设为默认行为
# 允许依赖层：core 收款计划/配置/JAB、tools 下已正式化的收款开单/明细/查询组件
# 谁不应该 import：core 层模块不应 import 本入口；凭证批量模块不应 import

import argparse
from dataclasses import asdict
from decimal import Decimal
import json
import sys
import threading
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
from tools.receipt_body_table_locator import locate_receipt_body_table_cached  # noqa: E402
from tools.receipt_detail_async_verifier import DetailPipelineVerifier  # noqa: E402
from tools.receipt_detail_row_cleanup import delete_extra_row_if_present  # noqa: E402
from tools.receipt_detail_rows import StepTimer, run_fee_only  # noqa: E402
from tools.receipt_detail_writer import (  # noqa: E402
    write_detail_line_by_screen,
    write_field_once,
)
from tools.receipt_keyboard_utils import foreground_matches_window  # noqa: E402
from tools.receipt_modal_guard import recover_cancelable_modal_now  # noqa: E402
from tools.receipt_new_probe import (  # noqa: E402
    annotate_foreground_root_for_targets,
    collect_receipt_new_windows,
    detect_self_made_entry_state,
    filter_usable_new_buttons,
    find_named_controls_in_windows,
    foreground_info,
)
from tools.receipt_post_save_query import run_post_save_batch_query  # noqa: E402
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    fill_header,
    find_receipt_header_field_by_dynamic_path,
    is_valid_customer_name_candidate,
    read_body_table,
    receipt_header_dynamic_prefix,
    resolve_receipt_header_anchor_in_canvas,
    run_receipt_new_probe,
    run_receipt_new_probe_with_jab,
    wait_header_account_description,
)

BODY_TABLE_SUFFIX = "0.0.0.1.1.0.0.0.0.1.0.2.1.0.0.0.0.0"


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
        help="高风险：键盘触发 Ctrl+S 保存收款单。默认不保存。",
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
        default=0.0,
        help="启动真实 NC 动作前等待秒数，默认 0。",
    )
    parser.add_argument(
        "--pause-after-header-field",
        default=None,
        help="诊断用：指定表头字段写完后暂停，回车后继续；例如 客户。",
    )
    parser.add_argument(
        "--diagnose-header-after-pause",
        action="store_true",
        help="诊断用：暂停恢复后读回已写表头字段，只报告不补救。",
    )
    parser.add_argument(
        "--diagnose-detail-repair",
        action="store_true",
        help="诊断用：明细 verifier 首次通过后强制模拟 1 个字段 pending，演练正式修复分支。",
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
        if args.start_delay > 0:
            print(f"启动后等待 {args.start_delay:g} 秒，请切到 NC 收款单录入页面。")
        else:
            print("启动后立即执行；请提前切到 NC 收款单录入页面。")
    if args.start_delay > 0:
        time.sleep(args.start_delay)

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
                config,
                row,
                save_enabled=args.save,
                recorder=recorder,
                pause_after_header_field=args.pause_after_header_field,
                diagnose_header_after_pause=args.diagnose_header_after_pause,
                diagnose_detail_repair=args.diagnose_detail_repair,
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
            post_query_issues = post_query_failure_reasons(post_query)
            if post_query_issues:
                report["post_query_failed_rows"] = post_query_issues
                exit_code = 1
        if args.write_selected_plan_sheet:
            workbook.write_batch_result_sheet(batch_results)

        report["ok"] = exit_code == 0
        report["total_seconds"] = round(time.perf_counter() - started, 3)
        write_last_report(report)
        recorder.finish("success" if exit_code == 0 else "failed")
        print_report(report, args)
        return exit_code
    except Exception as _exc:
        recorder.finish("failed", error=str(_exc))
        raise


def confirm_save(args):
    print("高风险：--save 会用键盘热键 Ctrl+S 真实保存 NC 收款单。")
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


def run_one_row(
    config,
    row,
    save_enabled=False,
    recorder=None,
    pause_after_header_field=None,
    diagnose_header_after_pause=False,
    diagnose_detail_repair=False,
):
    current_stage = {"name": ""}

    def _stage(stage, **fields):
        current_stage["name"] = stage
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
    jab = JABOperator(config)
    jab_lock = threading.RLock()
    pipeline_verifier = None
    modal_events = []
    try:
        timings.measure("jab.ensure-started", jab.ensure_started)
        _stage("开单", excel_row=row.row)
        open_step = timings.measure(
            "open.self-made",
            run_with_jab_lock,
            jab_lock,
            open_self_made_entry,
            config,
            jab,
        )
        row_report["steps"].append({"name": "open-self-made", **open_step})
        if not open_step.get("ok"):
            _event("open-failed", excel_row=row.row, error=open_step.get("reason"))
            return fail(row_report, "open-self-made", timings, open_step.get("reason"))
        row_report["modal_recovery"] = {"events": modal_events}

        def recover_modal_after_failure():
            result = recover_cancelable_modal_now(
                jab,
                stage=current_stage.get("name") or "",
            )
            if result.get("attempted"):
                modal_events.append(result)
            return result

        entry_scope_hwnd = extract_entry_scope_hwnd(open_step)
        entry_dynamic_index = extract_entry_dynamic_index(open_step)
        entry_anchor_path = extract_entry_anchor_path(open_step)
        if entry_scope_hwnd and entry_dynamic_index is None:
            anchor_retry = timings.measure(
                "header.anchor-retry-current-canvas",
                run_with_jab_lock,
                jab_lock,
                wait_receipt_header_anchor_in_current_canvas,
                jab,
                entry_scope_hwnd,
                timeout=1.2,
                interval=0.2,
            )
            row_report["entry_header_anchor_retry"] = anchor_retry
            if anchor_retry.get("ok"):
                entry_dynamic_index = anchor_retry.get("dynamic_index")
                entry_anchor_path = anchor_retry.get("label_path") or entry_anchor_path
        row_report["entry_scope_hwnd"] = entry_scope_hwnd
        row_report["entry_dynamic_index"] = entry_dynamic_index
        row_report["entry_anchor_path"] = entry_anchor_path
        row_report["locator_policy"] = {
            "header": (
                "财务组织用控件名粘贴+Enter并确认中文；其它表头字段优先用动态前缀"
                "+稳定后缀 path，path 失败才单字段语义兜底"
            ),
            "body": "明细表用校正后的动态前缀直接拼稳定 path，校验失败才现场扫描",
        }
        if not entry_scope_hwnd or entry_dynamic_index is None:
            reason = "当前 canvas 未解析到财务组织(O) 前缀，停止；不走语义兜底"
            _event("header-anchor-failed", excel_row=row.row, error=reason)
            return fail(row_report, "header-anchor", timings, reason)
        _stage("表头", excel_row=row.row)
        header_pause_reports = []
        header_steps_so_far_labels = []

        def after_header_field(label, _value, step):
            if label and label not in header_steps_so_far_labels:
                header_steps_so_far_labels.append(label)
            if pause_after_header_field != label:
                return None
            print(
                f"诊断暂停：表头字段 [{label}] 已写入。"
                "可以现在人工打开干扰窗口或清理已输入字段；回车后继续检查。"
            )
            input("完成干扰后按回车继续: ")
            report = {
                "ok": True,
                "paused_after": label,
                "field_path": step.get("path"),
            }
            if diagnose_header_after_pause:
                report["header_readback"] = diagnose_written_header_fields(
                    jab,
                    list(header_steps_so_far_labels),
                    step.get("dynamic_index"),
                    step.get("dynamic_prefix"),
                    entry_scope_hwnd,
                )
                report["ok"] = all(
                    item.get("present") for item in report["header_readback"]
                )
                if not report["ok"]:
                    report["reason"] = "暂停恢复后检测到已写表头字段为空或不可读"
            header_pause_reports.append(report)
            return report

        if pause_after_header_field:
            header_steps = timings.measure(
                "header.fill",
                run_with_jab_lock,
                jab_lock,
                fill_header,
                jab,
                business,
                scope_hwnd=entry_scope_hwnd,
                dynamic_index=entry_dynamic_index,
                anchor_path=entry_anchor_path,
                recover_after_failure=recover_modal_after_failure,
                after_field=after_header_field,
            )
        else:
            header_steps = timings.measure(
                "header.fill",
                run_with_jab_lock,
                jab_lock,
                fill_header,
                jab,
                business,
                after_field=after_header_field,
                scope_hwnd=entry_scope_hwnd,
                dynamic_index=entry_dynamic_index,
                anchor_path=entry_anchor_path,
                recover_after_failure=recover_modal_after_failure,
            )
        if header_pause_reports:
            row_report["header_pause_diagnostics"] = header_pause_reports
        row_report["header_steps"] = header_steps
        if any(not step.get("ok") for step in header_steps):
            header_error = summarize_header_failure(header_steps)
            _event("header-fill-failed", excel_row=row.row, error=header_error)
            return fail(row_report, "header-fill", timings, header_error)
        customer_name = timings.measure(
            "header.customer-name-readback",
            run_with_jab_lock,
            jab_lock,
            read_customer_name_after_header,
            jab,
            header_steps,
            entry_dynamic_index,
            entry_scope_hwnd,
        )
        row_report["customer_name_readback"] = customer_name
        row_report["nc_customer_name"] = str(customer_name.get("value") or "").strip()
        if not row_report["nc_customer_name"]:
            reason = customer_name.get("reason") or "客户名称未确认"
            _event(
                "header-customer-readback-failed",
                excel_row=row.row,
                error=reason,
            )
            return fail(row_report, "header-customer-name", timings, reason)
        _stage("明细主行", excel_row=row.row)
        located = timings.measure(
            "body.locate",
            run_with_jab_lock,
            jab_lock,
            resolve_body_table_by_dynamic_prefix,
            jab,
            entry_dynamic_index,
            entry_scope_hwnd,
        )
        row_report["body_locate"] = {
            "source": located.get("source"),
            "cache_hit": located.get("cache_hit"),
            "fallback_used": located.get("fallback_used"),
            "status": located.get("status"),
            "seconds": located.get("seconds"),
            "started_offset_seconds": located.get("started_offset_seconds"),
        }
        row_report["table_candidates"] = located.get("candidates", [])[:5]
        if not located.get("best"):
            _event("locate-table-failed", excel_row=row.row, error="未定位到明细表")
            return fail(row_report, "locate-body-table", timings, "未定位到明细表")
        row_report["before_table"] = {
            "ok": True,
            "skipped": True,
            "reason": "明细表 path 已定位；后台 pipeline verifier 负责预热 path 和并发读回",
        }
        pipeline_verifier = DetailPipelineVerifier(
            config,
            located,
            flow_started_at=flow_started_at,
            jab=jab,
            jab_lock=jab_lock,
        )
        pipeline_verifier.start()
        pipeline_field_task_ids = []
        pipeline_field_tasks = {}
        pipeline_snapshot_task_ids = []
        pipeline_row_count_task_id = None

        def submit_detail_verify(row_index, field, business_values, _step):
            task_id = pipeline_verifier.submit_field(
                row_index,
                field,
                business_values,
            )
            pipeline_field_task_ids.append(task_id)
            pipeline_field_tasks[task_id] = {
                "row_index": int(row_index),
                "field": dict(field),
                "business": business_values,
            }
            return task_id

        detail_steps = timings.measure(
            "detail.main-line",
            run_with_jab_lock,
            jab_lock,
            write_detail_line_by_screen,
            jab,
            business,
            located,
            after_field=submit_detail_verify,
            recover_after_failure=recover_modal_after_failure,
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
                run_with_jab_lock,
                jab_lock,
                run_fee_only,
                jab,
                located,
                str(row.fee),
                after_field=submit_detail_verify,
                recover_after_failure=recover_modal_after_failure,
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
                run_with_jab_lock,
                jab_lock,
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
        expected_detail_rows = 2 if row.fee > 0 else 1
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
        if diagnose_detail_repair:
            row_report["detail_pipeline_verify_before_repair_drill"] = dict(
                row_report["detail_pipeline_verify"]
            )
            row_report["detail_pipeline_verify"] = force_one_detail_field_pending(
                row_report["detail_pipeline_verify"],
                pipeline_field_task_ids,
            )
        row_report["detail_pipeline_state"] = verifier_snapshot(pipeline_verifier)
        timings.add(
            "detail.pipeline-final-wait",
            time.perf_counter() - pipeline_wait_started,
        )
        row_report["detail_pipeline_snapshots"] = pipeline_snapshot_task_ids
        detail_pipeline_ok = bool(row_report["detail_pipeline_verify"].get("ok"))
        if not detail_pipeline_ok:
            repair_report = timings.measure(
                "detail.pipeline-repair",
                repair_detail_pipeline_failures,
                jab,
                jab_lock,
                located,
                pipeline_verifier,
                row_report["detail_pipeline_verify"],
                pipeline_field_tasks,
                pipeline_row_count_task_id,
                expected_detail_rows,
                entry_scope_hwnd,
                recover_modal_after_failure,
            )
            row_report["detail_pipeline_repair"] = repair_report
            if repair_report.get("snapshot_task_id"):
                pipeline_snapshot_task_ids.append(repair_report["snapshot_task_id"])
            repair_wait_ids = repair_report.get("wait_ids") or []
            if repair_wait_ids:
                repair_wait_started = time.perf_counter()
                row_report["detail_pipeline_verify_after_repair"] = (
                    pipeline_verifier.wait(
                        repair_wait_ids,
                        timeout=2.0,
                    )
                )
                row_report["detail_pipeline_state_after_repair"] = verifier_snapshot(
                    pipeline_verifier
                )
                timings.add(
                    "detail.pipeline-repair-wait",
                    time.perf_counter() - repair_wait_started,
                )
                detail_pipeline_ok = bool(
                    row_report["detail_pipeline_verify_after_repair"].get("ok")
                )
        if not detail_pipeline_ok:
            row_report["after_table"] = timings.measure(
                "body.read-after-fallback",
                run_with_jab_lock,
                jab_lock,
                read_body_table,
                jab,
                "after_detail_fill",
                entry_scope_hwnd,
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
            run_with_jab_lock,
            jab_lock,
            wait_header_account_description,
            jab,
            0.0,
            scope=build_header_scope_for_followup(
                entry_scope_hwnd,
                entry_dynamic_index,
            ),
        )
        row_report["header_account"] = account_check
        if not account_check.get("accepted"):
            row_report["header_account_readback_warning"] = {
                "ok": False,
                "reason": "表头收款银行账户未从 JAB 后端读回；明细账号已由后台 pipeline 校验，继续保存/后验查询闭包",
                "account_check": account_check,
            }
            _event(
                "account-readback-warning",
                excel_row=row.row,
                warning="表头收款银行账户未从 JAB 后端读回，继续执行",
            )
        if save_enabled:
            _stage("保存", excel_row=row.row)
            save_result = timings.measure(
                "save.ctrl-s",
                run_with_jab_lock,
                jab_lock,
                save_receipt_by_ctrl_s,
                jab,
                entry_scope_hwnd,
            )
            if not save_result.get("ok"):
                recovery = timings.measure(
                    "save.modal-recovery-after-failure",
                    recover_modal_after_failure,
                )
                if recovery.get("attempted") and recovery.get("ok"):
                    save_result = timings.measure(
                        "save.ctrl-s-retry-after-modal",
                        run_with_jab_lock,
                        jab_lock,
                        save_receipt_by_ctrl_s,
                        jab,
                        entry_scope_hwnd,
                    )
                    save_result["retried_after_modal_recovery"] = True
                    save_result["modal_recovery"] = recovery
                else:
                    save_result["modal_recovery"] = recovery
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
                "reason": "no-save 模式：已停在保存前，未触发 Ctrl+S",
            }
        row_report["ok"] = True
        row_report["timings"] = timings.items
        return row_report
    except Exception as exc:
        row_report["exception"] = {
            "type": type(exc).__name__,
            "message": str(exc),
            "stage": current_stage.get("name") or "",
        }
        _event(
            "row-exception",
            excel_row=row.row,
            stage=current_stage.get("name") or "",
            error=f"{type(exc).__name__}: {exc}",
        )
        return fail(
            row_report,
            "exception",
            timings,
            f"{type(exc).__name__}: {exc}",
        )
    finally:
        if pipeline_verifier is not None:
            pipeline_verifier.close(timeout=0.2)
        row_report["modal_recovery"] = {"events": modal_events}
        jab.close()


def open_self_made_entry(config, jab=None):
    opened = (
        run_receipt_new_probe_with_jab(jab)
        if jab is not None
        else run_receipt_new_probe()
    )
    if opened.get("ok"):
        return opened
    opened["reason"] = opened.get("reason") or "未能进入收款单自制录入态"
    return opened


def run_with_jab_lock(jab_lock, func, *args, **kwargs):
    if jab_lock is None:
        return func(*args, **kwargs)
    with jab_lock:
        return func(*args, **kwargs)


def verifier_snapshot(verifier):
    if verifier is None or not hasattr(verifier, "snapshot"):
        return None
    return verifier.snapshot()


def wait_receipt_header_anchor_in_current_canvas(
    jab,
    scope_hwnd,
    timeout=1.2,
    interval=0.2,
):
    started_at = time.perf_counter()
    deadline = started_at + max(float(timeout or 0), 0.0)
    interval = max(float(interval or 0.2), 0.01)
    attempts = []
    while True:
        remaining = max(deadline - time.perf_counter(), 0.0)
        attempt = resolve_receipt_header_anchor_in_canvas(
            jab,
            scope_hwnd,
            timeout=min(0.05, remaining) if remaining > 0 else 0.05,
        )
        attempts.append(attempt)
        if attempt.get("ok"):
            return {
                **attempt,
                "attempts": attempts,
                "poll_interval": interval,
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        if time.perf_counter() >= deadline:
            return {
                "ok": False,
                "reason": attempt.get("reason") or "当前 canvas 未找到财务组织(O) 锚点",
                "scope_hwnd": scope_hwnd,
                "attempts": attempts,
                "poll_interval": interval,
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        time.sleep(min(interval, max(deadline - time.perf_counter(), 0.0)))


def force_one_detail_field_pending(report, field_task_ids):
    forced = dict(report or {})
    results = dict(forced.get("results") or {})
    target_id = next(
        (task_id for task_id in field_task_ids if task_id in results), None
    )
    if target_id is None and field_task_ids:
        target_id = field_task_ids[-1]
    if target_id and target_id in results:
        results.pop(target_id, None)
    submitted = list(forced.get("submitted") or [])
    if target_id and target_id not in submitted:
        submitted.append(target_id)
    pending = int(forced.get("pending") or 0)
    forced.update(
        {
            "ok": False,
            "pending": max(1, pending),
            "results": results,
            "submitted": submitted,
            "forced_detail_repair_drill": True,
            "forced_pending_field_task_id": target_id,
        }
    )
    return forced


def repair_detail_pipeline_failures(
    jab,
    jab_lock,
    located,
    pipeline_verifier,
    pipeline_report,
    pipeline_field_tasks,
    pipeline_row_count_task_id,
    expected_rows,
    scope_hwnd,
    recover_after_failure=None,
):
    results = (pipeline_report or {}).get("results") or {}
    best = (located or {}).get("best") or {}
    table_window = best.get("window") or {}
    row_count = int(best.get("row_count") or 0)
    repair = {
        "ok": False,
        "policy": (
            "只用当前已定位的明细表 path 修复一次；不重扫表格，不切换到旧语义兜底"
        ),
        "field_repairs": [],
        "row_count_repair": None,
        "wait_ids": [],
        "snapshot_task_id": None,
    }
    repair_field_ids = []
    for task_id, task in (pipeline_field_tasks or {}).items():
        result = results.get(task_id)
        if result and result.get("ok"):
            continue
        field = task["field"]
        if not best or not table_window:
            attempt = {
                "ok": False,
                "reason": "明细表缓存窗口不可用，不能执行字段修复",
            }
        else:
            attempt = run_with_jab_lock(
                jab_lock,
                write_field_once,
                jab,
                located,
                table_window,
                int(task["row_index"]),
                row_count,
                field,
                field["col"],
                task["business"],
                2,
                current_col=None,
                recover_after_failure=recover_after_failure,
            )
        field_report = {
            "original_task_id": task_id,
            "name": field.get("name"),
            "row_index": int(task["row_index"]),
            "col": field.get("col"),
            "attempt": attempt,
        }
        if attempt.get("ok"):
            verify_task_id = pipeline_verifier.submit_field(
                int(task["row_index"]),
                field,
                task["business"],
            )
            repair_field_ids.append(verify_task_id)
            field_report["verify_task_id"] = verify_task_id
        else:
            field_report["reason"] = (
                attempt.get("input_reason")
                or attempt.get("commit_reason")
                or attempt.get("reason")
            )
        repair["field_repairs"].append(field_report)

    row_count_result = results.get(pipeline_row_count_task_id)
    row_count_needs_repair = bool(pipeline_row_count_task_id) and (
        row_count_result is None or not row_count_result.get("ok")
    )
    row_count_wait_id = None
    if row_count_needs_repair:
        row_repair = run_with_jab_lock(
            jab_lock,
            delete_extra_row_if_present,
            jab,
            located,
            int(expected_rows),
            scope_hwnd=scope_hwnd,
        )
        repair["row_count_repair"] = row_repair
        if row_repair.get("ok"):
            row_count_wait_id = pipeline_verifier.submit_row_count(int(expected_rows))

    if repair_field_ids:
        repair["snapshot_task_id"] = pipeline_verifier.submit_snapshot(
            "after-detail-repair",
            max_rows=max(3, int(expected_rows) + 1),
            min_matches=len(repair_field_ids),
        )
    wait_ids = [*repair_field_ids]
    if row_count_wait_id:
        wait_ids.append(row_count_wait_id)
    repair["wait_ids"] = wait_ids
    attempted = bool(repair["field_repairs"]) or bool(repair["row_count_repair"])
    if not attempted:
        repair["ok"] = False
        repair["reason"] = "pipeline 失败但没有可修复的字段或行数任务"
    elif not wait_ids:
        repair["ok"] = False
        repair["reason"] = "已尝试修复，但没有成功提交二次校验任务"
    else:
        repair["ok"] = True
    return repair


def build_body_table_cached_path(dynamic_index, scope_hwnd=None):
    if dynamic_index is None:
        return None
    path = f"{receipt_header_dynamic_prefix(dynamic_index)}.{BODY_TABLE_SUFFIX}"
    return {
        "best": {
            "path": path,
            "window": {
                "hwnd": scope_hwnd,
                "class_name": "SunAwtCanvas",
            },
        }
    }


def resolve_body_table_by_dynamic_prefix(jab, dynamic_index, scope_hwnd=None):
    cached = build_body_table_cached_path(dynamic_index, scope_hwnd=scope_hwnd)
    located = locate_receipt_body_table_cached(
        jab,
        cached=cached,
        max_rows=5,
        scope_hwnd=scope_hwnd,
    )
    source = (
        "dynamic-prefix-body-path"
        if located.get("cache_hit")
        else "dynamic-prefix-body-path-fallback-scan"
    )
    return {**located, "source": source, "cached_path": (cached or {}).get("best")}


def build_header_scope_for_followup(scope_hwnd, dynamic_index):
    if not scope_hwnd or dynamic_index is None:
        return None
    return {
        "ok": True,
        "scope_hwnd": scope_hwnd,
        "dynamic_index": dynamic_index,
        "mode": "provided-canvas-anchor",
    }


def probe_receipt_entry_page(jab):
    windows = collect_receipt_new_windows(jab)
    state = detect_self_made_entry_state(windows)
    scope_hwnd = extract_entry_scope_hwnd({"entry_state": state, "windows": windows})
    return {
        "ok": bool(state.get("ok")),
        "method": "entry-state",
        "scope_hwnd": scope_hwnd,
        "windows": windows,
        "entry_state": state,
    }


def extract_entry_scope_hwnd(report):
    state = (report or {}).get("entry_state") or {}
    hwnd = extract_entry_state_hwnd(state, prefer_canvas=True)
    if hwnd:
        return hwnd
    parsed = (report or {}).get("parsed") or {}
    for key in ("entry_state", "quick_entry_state"):
        hwnd = extract_entry_state_hwnd(parsed.get(key) or {}, prefer_canvas=True)
        if hwnd:
            return hwnd
    hwnd = extract_entry_state_hwnd(state, prefer_canvas=False)
    if hwnd:
        return hwnd
    for key in ("entry_state", "quick_entry_state"):
        hwnd = extract_entry_state_hwnd(parsed.get(key) or {}, prefer_canvas=False)
        if hwnd:
            return hwnd
    for key in (
        "windows_after_choose",
        "windows_after_open",
        "windows",
        "after_windows",
    ):
        for window in (report or {}).get(key) or parsed.get(key) or []:
            if (
                window.get("is_java")
                and window.get("visible")
                and window.get("hwnd")
                and window.get("class_name") == "SunAwtCanvas"
            ):
                return int(window["hwnd"])
    return None


def extract_entry_dynamic_index(report):
    state = (report or {}).get("entry_state") or {}
    dynamic_index = extract_dynamic_index_from_entry_state(state)
    if dynamic_index is not None:
        return dynamic_index
    parsed = (report or {}).get("parsed") or {}
    for key in ("entry_state", "quick_entry_state"):
        dynamic_index = extract_dynamic_index_from_entry_state(parsed.get(key) or {})
        if dynamic_index is not None:
            return dynamic_index
    return None


def extract_entry_anchor_path(report):
    state = (report or {}).get("entry_state") or {}
    path = extract_anchor_path_from_entry_state(state)
    if path:
        return path
    parsed = (report or {}).get("parsed") or {}
    for key in ("entry_state", "quick_entry_state"):
        path = extract_anchor_path_from_entry_state(parsed.get(key) or {})
        if path:
            return path
    return None


def extract_anchor_path_from_entry_state(state):
    for hit in (state or {}).get("hits") or []:
        control = hit.get("control") or {}
        if (
            control.get("name") == "财务组织(O)"
            or control.get("description") == "财务组织(O)"
        ):
            path = control.get("path")
            if path:
                return str(path)
    return None


def extract_dynamic_index_from_entry_state(state):
    for hit in (state or {}).get("hits") or []:
        control = hit.get("control") or {}
        direct = control.get("dynamic_index")
        if direct is not None:
            try:
                return int(direct)
            except (TypeError, ValueError):
                pass
        path = control.get("path") or ""
        dynamic_index = extract_receipt_module_dynamic_index(path)
        if dynamic_index is not None:
            return dynamic_index
    return None


def extract_receipt_module_dynamic_index(path):
    prefix = "0.0.1.0.0.0.0."
    text = str(path or "")
    if not text.startswith(prefix):
        return None
    part = text[len(prefix) :].split(".", 1)[0]
    try:
        return int(part)
    except ValueError:
        return None


def extract_entry_state_hwnd(state, prefer_canvas=False):
    if prefer_canvas:
        for hit in (state or {}).get("hits") or []:
            window = hit.get("window") or {}
            if window.get("class_name") == "SunAwtCanvas" and window.get("hwnd"):
                return int(window["hwnd"])
    for hit in (state or {}).get("hits") or []:
        window = hit.get("window") or {}
        hwnd = window.get("hwnd")
        if hwnd:
            return int(hwnd)
    return None


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


def read_customer_name_after_header(jab, header_steps, dynamic_index, scope_hwnd):
    step = next(
        (item for item in header_steps or [] if item.get("label") == "客户"),
        None,
    )
    attempts = []
    if step and step.get("path"):
        found = find_receipt_header_field_by_dynamic_path(
            jab,
            "客户",
            step.get("dynamic_index") or dynamic_index,
            scope_hwnd=scope_hwnd,
            require_showing=False,
            require_valid_bounds=False,
            path_template=(step.get("path_attempt") or {}).get("path_template"),
        )
        attempts.append(
            read_customer_name_from_found_field(jab, found, source="path-readback")
        )
    if step:
        attempts.append(
            {
                "ok": True,
                "source": "header-step-snapshot",
                "value": extract_header_accepted_text([step], "客户"),
                "snapshot": step.get("post_write_snapshot")
                or step.get("backend_state")
                or {},
            }
        )
    for attempt in attempts:
        value = str(attempt.get("value") or "").strip()
        if is_valid_customer_name_candidate(value):
            return {
                "ok": True,
                "value": value,
                "source": attempt.get("source"),
                "attempts": attempts,
            }
    return {
        "ok": False,
        "value": "",
        "attempts": attempts,
        "reason": "客户名称未确认：客户字段 description 未读到有效 NC 客户名称",
    }


def read_customer_name_from_found_field(jab, found, source):
    if not found.get("ok"):
        return {
            "ok": False,
            "source": source,
            "reason": found.get("reason"),
            "path": found.get("path"),
        }
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    try:
        info = jab.get_context_info(vm_id, context)
        text = jab.get_text_context_value(vm_id, context)
        description = info.description.strip() if info else ""
        name = info.name.strip() if info else ""
        value = first_valid_text(description, text, name)
        return {
            "ok": bool(value),
            "source": source,
            "value": value,
            "path": found.get("path"),
            "label_path": found.get("label_path"),
            "text": text,
            "name": name,
            "description": description,
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def first_valid_text(*values):
    for value in values:
        text = str(value or "").strip()
        if is_valid_customer_name_candidate(text):
            return text
    return ""


def extract_header_accepted_text(header_steps, label):
    for step in header_steps or []:
        if step.get("label") != label:
            continue
        text = str(step.get("accepted_text") or "").strip()
        if is_valid_customer_name_candidate(text):
            return text
        backend = step.get("post_write_snapshot") or step.get("backend_state") or {}
        for key in ("description", "text", "name"):
            value = str(backend.get(key) or "").strip()
            if (
                value
                and value != str(step.get("value") or "").strip()
                and is_valid_customer_name_candidate(value)
            ):
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


def post_query_failure_reasons(post_query):
    if not post_query or not post_query.get("ok"):
        return {"*": (post_query or {}).get("reason") or "后验查询失败"}
    issues = {}
    for group in post_query.get("groups") or []:
        match = group.get("match") or {}
        for row, reason in (match.get("issues") or {}).items():
            issues[str(row)] = reason or "后验未匹配"
        if not group.get("ok"):
            reason = group.get("reason") or "后验查询失败"
            for row in group.get("target_rows") or []:
                issues.setdefault(str(row), reason)
    return issues


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


def save_receipt_by_ctrl_s(jab, scope_hwnd=None, timeout=1.0):
    page = None
    if not scope_hwnd:
        page = probe_receipt_entry_page(jab)
        if not page.get("ok"):
            return {
                "ok": False,
                "triggered": False,
                "reason": "Ctrl+S 保存前未确认当前是收款单自制录入页",
                "page": page,
            }
        scope = page.get("scope") or {}
        scope_hwnd = scope.get("scope_hwnd") or page.get("scope_hwnd")
    if not scope_hwnd:
        return {
            "ok": False,
            "triggered": False,
            "reason": "Ctrl+S 保存前未取得收款单窗口句柄",
            "page": page,
        }
    guard = foreground_matches_window({"hwnd": scope_hwnd})
    if not guard.get("ok"):
        return {
            "ok": False,
            "triggered": False,
            "reason": guard.get("reason") or "当前前台窗口不是目标 NC 窗口",
            "guard": guard,
            "page": page,
        }
    try:
        jab.press_hotkey("ctrl", "s", wait=0)
    except Exception as exc:
        return {
            "ok": False,
            "triggered": False,
            "reason": f"Ctrl+S 键盘热键触发失败：{type(exc).__name__}: {exc}",
            "guard": guard,
            "page": page,
        }
    started = time.perf_counter()
    last_state = None
    while time.perf_counter() - started < timeout:
        windows = collect_receipt_new_windows(jab)
        state = detect_self_made_entry_state(windows)
        last_state = state
        parent_new_state = detect_receipt_parent_new_ready(windows)
        if parent_new_state.get("ok") and not state.get("ok"):
            return {
                "ok": True,
                "triggered": True,
                "hotkey": {"ok": True, "mode": "jab.press_hotkey", "key": "Ctrl+S"},
                "precondition": {
                    "page": page,
                    "foreground_guard": guard,
                },
                "seconds": round(time.perf_counter() - started, 3),
                "oracle": {
                    "name": "receipt_parent_new_ready_after_save",
                    "ok": True,
                    "evidence": "保存后重新检测到收款单录入父页前台【新增】按钮，且保存/暂存/取消三按钮不再同时存在",
                    "parent_new_state": parent_new_state,
                    "self_made_entry_state": state,
                },
                "entry_state": state,
                "parent_new_state": parent_new_state,
            }
        time.sleep(0.2)
    final_windows = collect_receipt_new_windows(jab)
    parent_new_state = detect_receipt_parent_new_ready(final_windows)
    return {
        "ok": False,
        "triggered": True,
        "hotkey": {"ok": True, "mode": "jab.press_hotkey", "key": "Ctrl+S"},
        "precondition": {
            "page": page,
            "foreground_guard": guard,
        },
        "seconds": round(time.perf_counter() - started, 3),
        "reason": "保存后未确认收款单父页【新增】已恢复，不能证明保存成功",
        "oracle": {
            "name": "receipt_parent_new_ready_after_save",
            "ok": False,
            "evidence": "需要同时满足：前台收款单父页【新增】按钮可用，且保存/暂存/取消三按钮不再同时存在",
            "parent_new_state": parent_new_state,
            "self_made_entry_state": last_state,
        },
        "entry_state": last_state,
        "parent_new_state": parent_new_state,
    }


def detect_receipt_parent_new_ready(windows):
    foreground = foreground_info()
    buttons = find_named_controls_in_windows(
        windows,
        "新增",
        role=None,
        class_name="SunAwtFrame",
        require_action=True,
    )
    annotate_foreground_root_for_targets(buttons, foreground)
    usable = filter_usable_new_buttons(buttons, foreground)
    return {
        "ok": bool(usable),
        "foreground": foreground,
        "usable_new_button_count": len(usable),
        "usable_new_buttons": [
            {
                "window": item.get("window"),
                "control": {
                    key: (item.get("control") or {}).get(key)
                    for key in ("name", "description", "role", "states", "path")
                },
            }
            for item in usable[:3]
        ],
        "candidate_count": len(buttons),
    }


def diagnose_written_header_fields(
    jab,
    labels,
    dynamic_index,
    dynamic_prefix,
    scope_hwnd,
):
    results = []
    for label in labels or []:
        found = find_receipt_header_field_by_dynamic_path(
            jab,
            label,
            dynamic_index,
            scope_hwnd=scope_hwnd,
            require_showing=False,
            require_valid_bounds=False,
        )
        if not found.get("ok"):
            results.append(
                {
                    "label": label,
                    "ok": False,
                    "present": False,
                    "reason": found.get("reason") or "field path not found",
                    "dynamic_prefix": dynamic_prefix,
                    "path": found.get("path"),
                }
            )
            continue
        context = found["context"]
        vm_id = found["vm_id"]
        owned_contexts = found["owned_contexts"]
        try:
            info = jab.get_context_info(vm_id, context)
            text = jab.get_text_context_value(vm_id, context)
            description = info.description.strip() if info else ""
            name = info.name.strip() if info else ""
            present = bool(str(text or "").strip() or description or name)
            results.append(
                {
                    "label": label,
                    "ok": True,
                    "present": present,
                    "text": text,
                    "description": description,
                    "name": name,
                    "path": found.get("path"),
                    "dynamic_prefix": found.get("dynamic_prefix") or dynamic_prefix,
                }
            )
        finally:
            jab.release_contexts(vm_id, owned_contexts)
    return results


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


def summarize_header_failure(header_steps):
    for step in header_steps or []:
        if step.get("ok"):
            continue
        label = step.get("label")
        reason = (
            step.get("reason")
            or step.get("stage")
            or ((step.get("scope") or {}).get("reason"))
            or ((step.get("path_attempt") or {}).get("reason"))
            or "表头字段写入失败"
        )
        return f"表头字段写入失败: {label or step.get('step') or '未知字段'} - {reason}"
    return "表头字段写入失败"


def serializable(value):
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, dict):
        return {key: serializable(item) for key, item in value.items()}
    if isinstance(value, list | tuple):
        return [serializable(item) for item in value]
    return value


def write_last_report(report):
    path = ROOT / "logs" / "last_receipt_full_flow_report.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(".json.tmp")
    with tmp_path.open("w", encoding="utf-8") as file:
        json.dump(report, file, ensure_ascii=False, indent=2, default=str)
        file.write("\n")
    tmp_path.replace(path)
    summary_path = ROOT / "logs" / "last_receipt_failure_summary.txt"
    summary_path.write_text(
        "\n".join(build_console_report_lines(report, path, summary_path)) + "\n",
        encoding="utf-8",
    )
    return path


def print_report(report, args):
    if args.json:
        text = json.dumps(report, ensure_ascii=False, indent=2, default=str)
        print(text)
        return
    report_path = ROOT / "logs" / "last_receipt_full_flow_report.json"
    summary_path = ROOT / "logs" / "last_receipt_failure_summary.txt"
    for line in build_console_report_lines(report, report_path, summary_path):
        print(line)


def build_console_report_lines(report, report_path=None, summary_path=None):
    lines = ["收款单完整流程结果摘要"]
    lines.append(f"结果：{'成功' if report.get('ok') else '失败'}")
    lines.append(f"总用时：{report.get('total_seconds')}s")
    rows = report.get("rows") or []
    failed_row = next((row for row in rows if not row.get("ok")), None)
    if failed_row:
        lines.append(f"失败行：Sheet 行 {failed_row.get('excel_row')}")
        lines.append(f"失败阶段：{failed_row.get('failed_step') or '未知'}")
        if failed_row.get("reason"):
            lines.append(f"失败原因：{failed_row.get('reason')}")
        entry_scope_hwnd = failed_row.get("entry_scope_hwnd")
        entry_dynamic_index = failed_row.get("entry_dynamic_index")
        if entry_scope_hwnd is not None or entry_dynamic_index is not None:
            lines.append(
                "入口上下文："
                f"scope_hwnd={entry_scope_hwnd}, "
                f"entry_dynamic_index={entry_dynamic_index}"
            )
        header_step = first_failed_step(failed_row.get("header_steps"))
        if header_step:
            lines.extend(format_header_failure_lines(header_step))
        modal_events = ((failed_row.get("modal_recovery") or {}).get("events")) or []
        if modal_events:
            last_modal = modal_events[-1]
            lines.append(
                "弹窗恢复："
                f"attempted={last_modal.get('attempted')}, "
                f"ok={last_modal.get('ok')}, "
                f"reason={last_modal.get('reason') or ''}"
            )
        else:
            lines.append("弹窗恢复：本次失败点没有检测到可取消弹窗")
        timings = failed_row.get("timings") or []
        if timings:
            lines.append("关键耗时：" + format_timings(timings))
    elif report.get("ok"):
        ok_rows = [row.get("excel_row") for row in rows if row.get("ok")]
        lines.append(f"通过行：{ok_rows}")
    elif report.get("post_query_failed_rows"):
        ok_rows = [row.get("excel_row") for row in rows if row.get("ok")]
        lines.append(f"录入保存通过行：{ok_rows}")
        lines.append("失败阶段：post-query")
        for row, reason in (report.get("post_query_failed_rows") or {}).items():
            lines.append(f"后验未匹配行 {row}：{reason}")
    elif report.get("reason"):
        lines.append(f"失败原因：{report.get('reason')}")
    if report_path:
        lines.append(f"完整报告：{report_path}")
    if summary_path:
        lines.append(f"摘要文件：{summary_path}")
    return lines


def first_failed_step(steps):
    for step in steps or []:
        if step.get("ok"):
            continue
        return step
    return None


def format_header_failure_lines(step):
    lines = []
    label = step.get("label") or step.get("step") or "未知字段"
    lines.append(f"表头失败字段：{label}")
    if step.get("stage"):
        lines.append(f"表头失败阶段：{step.get('stage')}")
    scope = step.get("scope") or {}
    if scope:
        lines.append(
            "表头 scope："
            f"mode={scope.get('mode')}, "
            f"dynamic_index={scope.get('dynamic_index')}, "
            f"dynamic_prefix={scope.get('dynamic_prefix')}"
        )
    path_attempt = step.get("path_attempt") or {}
    if path_attempt:
        lines.append(
            "path 尝试："
            f"{path_attempt.get('path') or ''} "
            f"({path_attempt.get('reason') or '无原因'})"
        )
    modal_recovery = step.get("modal_recovery") or {}
    if modal_recovery:
        lines.append(
            "字段级弹窗恢复："
            f"attempted={modal_recovery.get('attempted')}, "
            f"ok={modal_recovery.get('ok')}, "
            f"reason={modal_recovery.get('reason') or ''}"
        )
    return lines


def format_timings(timings):
    chunks = []
    for item in timings:
        name = item.get("name")
        seconds = item.get("seconds")
        if name is None or seconds is None:
            continue
        chunks.append(f"{name}={seconds}s")
    return ", ".join(chunks)


if __name__ == "__main__":
    raise SystemExit(main())
