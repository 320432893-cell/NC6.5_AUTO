# 职责：编排收款单完整流程测试入口，消费 ReceiptPlanRow 跑开单/表头/明细/手续费/可选保存
# 不做什么：不做录入前 NC 查重，不复用历史 T0 保存脚本，不把保存设为默认行为
# 允许依赖层：core 收款计划/配置/JAB、tools 下已正式化的收款开单/明细/查询组件
# 谁不应该 import：core 层模块不应 import 本入口；凭证批量模块不应 import

import argparse
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402, F401
from core.receipt_entry import ReceiptEntryWorkbook  # noqa: E402
from core.receipt_models import ReceiptBatchResultRow  # noqa: E402
from core.run_state import RunStateRecorder  # noqa: E402
from core.utils import check_abort, load_config  # noqa: E402

# 以下外部协作者在本入口命名空间内保留为模块属性：拆分出去的 flow 子模块通过
# `import tools.receipt_full_flow_entry as _flow` 在调用时读取它们，使测试对
# tools.receipt_full_flow_entry.<name> 的 monkeypatch 与拆分前完全一致地生效。
from tools.receipt_body_table_locator import (  # noqa: E402, F401
    locate_receipt_body_table_cached,
)
from tools.receipt_detail_async_verifier import (  # noqa: E402, F401
    DetailPipelineVerifier,
)
from tools.receipt_detail_row_cleanup import (  # noqa: E402, F401
    delete_extra_row_if_present,
)
from tools.receipt_detail_writer import (  # noqa: E402, F401
    write_detail_line_by_screen,
    write_field_once,
)
from tools.receipt_flow_detail_repair import (  # noqa: E402, F401
    force_one_detail_field_pending,
    repair_detail_pipeline_failures,
)
from tools.receipt_flow_entry_state import (  # noqa: E402, F401
    BODY_TABLE_SUFFIX,
    build_body_table_cached_path,
    build_header_scope_for_followup,
    extract_anchor_path_from_entry_state,
    extract_dynamic_index_from_entry_state,
    extract_entry_anchor_path,
    extract_entry_dynamic_index,
    extract_entry_scope_hwnd,
    extract_entry_state_hwnd,
    extract_receipt_module_dynamic_index,
    resolve_body_table_by_dynamic_prefix,
    run_with_jab_lock,
    wait_receipt_header_anchor_in_current_canvas,
)
from tools.receipt_flow_header_diag import (  # noqa: E402, F401
    diagnose_written_header_fields,
    extract_header_accepted_text,
    first_valid_text,
    read_customer_name_after_header,
    read_customer_name_from_found_field,
    summarize_header_failure,
)
from tools.receipt_flow_report import (  # noqa: E402, F401
    build_console_report_lines,
    first_failed_step,
    format_header_failure_lines,
    format_timings,
    print_report,
    serializable,
    write_last_report,
)
from tools.receipt_flow_row_runner import (  # noqa: E402, F401
    business_from_plan_row,
    fail,
    open_self_made_entry,
    run_one_row,
    verifier_snapshot,
)
from tools.receipt_flow_save import (  # noqa: E402, F401
    detect_receipt_parent_new_ready,
    probe_receipt_entry_page,
    save_receipt_by_ctrl_s,
)
from tools.receipt_keyboard_utils import (  # noqa: E402, F401
    foreground_matches_window,
)
from tools.receipt_modal_guard import (  # noqa: E402, F401
    recover_cancelable_modal_now,
)
from tools.receipt_new_probe import (  # noqa: E402, F401
    collect_receipt_new_windows,
    detect_self_made_entry_state,
)
from tools.receipt_post_save_query import run_post_save_batch_query  # noqa: E402
from tools.receipt_self_made_flow import (  # noqa: E402, F401
    fill_header,
    find_receipt_header_field_by_dynamic_path,
    read_body_table,
    resolve_receipt_header_anchor_in_canvas,
    run_receipt_new_probe,
    run_receipt_new_probe_with_jab,
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
            check_abort()  # 行边界轮询外部停止标志：停在可续跑位置
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
    except BaseException as _exc:
        # 外部停止标志/紧急停止经 JAB 原语的 check_abort 抛出 SystemExit；
        # 按 ENGINE_CONTRACT.md §1.4 收尾为 aborted、退出码 3、可续跑，
        # GUI 由此把"用户停止"与"崩溃"区分开（对齐 jab_batch.py 的处理）。
        if isinstance(_exc, (KeyboardInterrupt, SystemExit)):
            recorder.finish("aborted", error=f"{type(_exc).__name__}: {_exc}")
            report["ok"] = False
            report["aborted"] = True
            report["reason"] = f"已停止：{_exc}"
            report["resumable"] = {
                "can_resume": True,
                "resume_hint": "重跑本入口并用 --excel-rows 指定未完成行",
            }
            report["total_seconds"] = round(time.perf_counter() - started, 3)
            write_last_report(report)
            print_report(report, args)
            return 3
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


if __name__ == "__main__":
    raise SystemExit(main())
