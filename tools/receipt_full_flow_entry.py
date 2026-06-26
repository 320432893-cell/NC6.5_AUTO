# 职责：编排收款单完整流程入口，消费 ReceiptPlanRow 跑开单/表头/明细/手续费/保存闸/后验查询
# 不做什么：不做录入前 NC 查重，不复用历史 T0 保存脚本，不把保存设为默认行为
# 允许依赖层：core 收款计划/配置/JAB、tools 下已正式化的收款开单/明细/查询组件
# 谁不应该 import：core 层模块不应 import 本入口；凭证批量模块不应 import

import argparse
import ctypes
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

from core.errors import ExcelLockedError  # noqa: E402
from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_entry import ReceiptEntryWorkbook  # noqa: E402
from core.receipt_models import ReceiptBatchResultRow  # noqa: E402
from core.run_state import RunStateRecorder  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table_cached  # noqa: E402
from tools.receipt_detail_async_verifier import DetailPipelineVerifier  # noqa: E402
from tools.receipt_detail_fields import validate_exchange_rate_not_polluted  # noqa: E402
from tools.receipt_detail_row_cleanup import delete_extra_row_if_present  # noqa: E402
from tools.receipt_detail_rows import StepTimer, run_fee_only  # noqa: E402
from tools.receipt_detail_writer import (  # noqa: E402
    write_detail_line_by_screen,
    write_field_once,
)
from tools.receipt_keyboard_utils import (  # noqa: E402
    foreground_matches_window,
    send_hotkey_alt_y,
    send_hotkey_ctrl_q,
    send_hotkey_ctrl_s,
)
from tools.receipt_modal_guard import (  # noqa: E402
    collect_visible_java_dialogs,
    focus_window,
    recover_cancelable_modal_now,
)
from tools.receipt_new_probe import (  # noqa: E402
    annotate_foreground_root_for_targets,
    collect_receipt_new_windows,
    detect_self_made_entry_state,
    filter_usable_new_buttons,
    find_named_controls_in_windows,
    foreground_info,
    root_hwnd,
)
from tools.receipt_post_save_query import run_post_save_batch_query  # noqa: E402
from tools.jab_probe import JOBJECT  # noqa: E402
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    HEADER_SCOPE_ANCHOR_LABEL,
    fill_header,
    find_finance_org_header_scope_by_paths,
    find_receipt_header_field_by_dynamic_path,
    find_receipt_header_field_by_live_semantic,
    get_receipt_header_path_template,
    infer_header_path_template_from_field,
    is_valid_customer_name_candidate,
    read_body_table,
    receipt_header_dynamic_prefix,
    resolve_receipt_header_anchor_in_canvas,
    run_receipt_new_probe,
    run_receipt_new_probe_with_jab,
    set_receipt_header_path_template,
    wait_header_account_description,
    guarded_paste_header_value,
    describe_backend_field_state,
)

COUNTERPARTY_LABEL = "往来对象"
COUNTERPARTY_EXPECTED = "客户"
COUNTERPARTY_KNOWN_OPTIONS = {"客户", "部门", "业务员", "供应商"}
COUNTERPARTY_NEARBY_MAX_VERTICAL_DISTANCE = 36
COUNTERPARTY_NEARBY_MAX_RIGHT_DISTANCE = 700
COUNTERPARTY_STATE_OK = "ok"
COUNTERPARTY_STATE_REPAIRABLE = "repairable-empty"
COUNTERPARTY_STATE_CONFLICT = "conflict"
COUNTERPARTY_STATE_DETAIL_UNREADABLE = "detail-unreadable"
_COUNTERPARTY_NEARBY_SUFFIX_CACHE = {}
_BODY_TABLE_SUFFIX_CACHE = {}
SLOW_STEP_THRESHOLD_SECONDS = 0.5
CIRCUIT_BREAKER_RETRY_STEPS = {
    "detail-extra-text-verify",
    "detail-pipeline-verify",
    "detail-exchange-rate-guard",
}


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "收款单完整流程入口：消费 ReceiptPlanRow，从指定 Sheet1 行开始往下处理；"
            "--limit 不传或 0 表示做到表尾。--save 是真实保存安全闸。"
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--excel-path", default=None)
    parser.add_argument(
        "--start-row",
        type=int,
        default=None,
        help="从指定 Sheet1 行开始往下处理，包含该行",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="从起始行往下做几条；不传或 0 表示做到表尾",
    )
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
        help="高风险：键盘触发 Ctrl+S 保存收款单；正式桌面入口默认开启。",
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
        "--excel-text-field-map",
        action="append",
        default=[],
        metavar="EXCEL列名=NC文本名",
        help="把 Sheet1 指定列的非空值写入 NC 同名/近邻文本框；可重复。",
    )
    parser.add_argument(
        "--excel-column",
        action="append",
        default=[],
        metavar="配置键=EXCEL列名",
        help="覆盖 receipt_entry.excel 下固定列名配置；例如 raw_amount_column=🟪原始金额。",
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
    apply_excel_column_overrides(config, args.excel_column)
    apply_excel_text_field_mappings(config, args.excel_text_field_map)
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
    if args.start_row is None or args.start_row <= 0:
        raise SystemExit("--start-row 必填且必须大于 0")
    if args.limit is not None and args.limit < 0:
        raise SystemExit("--limit 必须大于等于 0")
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
    try:
        if args.write_plan_sheet:
            workbook.write_plan_sheet(plan_rows, issues)
        elif args.write_selected_plan_sheet:
            workbook.write_plan_sheet(
                selected_rows,
                filter_issues_for_rows(issues, selected_rows),
            )
    except ExcelLockedError as exc:
        reason = user_excel_locked_message(exc)
        report.update(
            {
                "ok": False,
                "reason": reason,
                "error_category": "excel_locked",
                "total_seconds": round(time.perf_counter() - started, 3),
            }
        )
        recorder.event("excel-write-failed", reason=reason)
        recorder.finish("failed", error=reason)
        write_last_report(report)
        print_report(report, args)
        return 1

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
    header_scope_cache = {}
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
                header_scope_cache=header_scope_cache,
            )
            if should_retry_row_by_cancel_reopen(row_report):
                recorder.event(
                    "row-circuit-breaker-start",
                    excel_row=row.row,
                    failed_step=row_report.get("failed_step"),
                    reason=row_report.get("reason"),
                )
                cancel_report = cancel_current_receipt_entry(config)
                recorder.event(
                    "row-circuit-breaker-cancel",
                    excel_row=row.row,
                    ok=bool(cancel_report.get("ok")),
                    reason=cancel_report.get("reason") or "",
                )
                if cancel_report.get("ok"):
                    retry_report = run_one_row(
                        config,
                        row,
                        save_enabled=args.save,
                        recorder=recorder,
                        pause_after_header_field=args.pause_after_header_field,
                        diagnose_header_after_pause=args.diagnose_header_after_pause,
                        diagnose_detail_repair=args.diagnose_detail_repair,
                        header_scope_cache=header_scope_cache,
                    )
                    retry_report["circuit_breaker"] = {
                        "triggered": True,
                        "retry_count": 1,
                        "retryable_steps": sorted(CIRCUIT_BREAKER_RETRY_STEPS),
                        "first_attempt": summarize_retry_attempt(row_report),
                        "cancel": cancel_report,
                    }
                    row_report = retry_report
                    recorder.event(
                        "row-circuit-breaker-retry-done",
                        excel_row=row.row,
                        ok=bool(row_report.get("ok")),
                        failed_step=row_report.get("failed_step") or "",
                    )
                else:
                    row_report["circuit_breaker"] = {
                        "triggered": True,
                        "retry_count": 0,
                        "retryable_steps": sorted(CIRCUIT_BREAKER_RETRY_STEPS),
                        "first_attempt": summarize_retry_attempt(row_report),
                        "cancel": cancel_report,
                        "reason": "取消当前未保存收款单失败，拒绝重试避免叠单",
                    }
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
        report["post_query_requested"] = bool(args.query_after_save)
        report["post_query_executed"] = False
        if args.query_after_save and report["rows"] and exit_code == 0:
            recorder.set_stage("后验查询", step_index=len(report["rows"]), total_steps=len(selected_rows))
            recorder.event(
                "post-query-start",
                saved_rows=[row.get("excel_row") for row in report["rows"] if row.get("ok")],
            )
            batch_results, post_query = run_post_save_batch_query(
                config,
                selected_rows,
                report["rows"],
            )
            report["post_query"] = post_query
            report["post_query_executed"] = True
            recorder.event(
                "post-query-done",
                ok=bool(post_query.get("ok")),
                reason=post_query.get("reason") or "",
            )
            post_query_issues = post_query_failure_reasons(post_query)
            if post_query_issues:
                report["post_query_failed_rows"] = post_query_issues
                exit_code = 1
        elif args.query_after_save:
            report["post_query_skipped"] = {
                "reason": post_query_skip_reason(report["rows"], exit_code),
                "rows": len(report["rows"]),
                "exit_code": exit_code,
            }
            recorder.event(
                "post-query-skipped",
                reason=report["post_query_skipped"]["reason"],
                rows=len(report["rows"]),
                exit_code=exit_code,
            )
        if args.write_selected_plan_sheet:
            try:
                workbook.write_batch_result_sheet(batch_results)
            except ExcelLockedError as exc:
                report["excel_result_write_error"] = user_excel_locked_message(exc)

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


def user_excel_locked_message(exc):
    return (
        "Excel 文件无法写入。请先关闭正在打开的 Excel/WPS 文件、关闭资源管理器预览窗格，"
        "或取消“写入选中计划 Sheet2”后重试；原始错误："
        f"{exc}"
    )


def cache_receipt_header_scope(jab, shared_cache, scope):
    if not isinstance(scope, dict) or not scope.get("ok"):
        return
    cached = dict(scope)
    try:
        setattr(jab, "_receipt_header_scope_cache", cached)
    except AttributeError:
        pass
    if isinstance(shared_cache, dict):
        shared_cache.clear()
        shared_cache.update(cached)


def select_plan_rows(plan_rows, issues, args):
    issue_rows = {issue.excel_row for issue in issues if issue.excel_row is not None}
    runnable = [row for row in plan_rows if row.row not in issue_rows]
    start_row = getattr(args, "start_row", None)
    if start_row is None or start_row <= 0:
        raise SystemExit("--start-row 必填且必须大于 0")
    runnable = [row for row in runnable if row.row >= start_row]
    limit = max(int(args.limit or 0), 0)
    if limit:
        runnable = runnable[:limit]
    return runnable


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
    header_scope_cache=None,
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
        entry_dynamic_index_source = (
            "entry-state" if entry_dynamic_index is not None else None
        )
        if entry_scope_hwnd and entry_dynamic_index is None:
            preferred_cached_dynamic_index = None
            cached_scope = {}
            if isinstance(header_scope_cache, dict):
                cached_scope = header_scope_cache
            if not cached_scope.get("ok"):
                cached_scope = getattr(jab, "_receipt_header_scope_cache", None) or {}
            if cached_scope.get("ok") and cached_scope.get("dynamic_index") is not None:
                preferred_cached_dynamic_index = cached_scope.get("dynamic_index")
            if (
                cached_scope.get("ok")
                and cached_scope.get("scope_hwnd") == entry_scope_hwnd
                and cached_scope.get("dynamic_index") is not None
            ):
                entry_dynamic_index = cached_scope.get("dynamic_index")
                entry_dynamic_index_source = "header-scope-cache"
                entry_anchor_path = (
                    cached_scope.get("label_path")
                    or cached_scope.get("semantic_label_path")
                    or entry_anchor_path
                )
                row_report["entry_header_scope_cache"] = {
                    "ok": True,
                    "scope_hwnd": entry_scope_hwnd,
                    "dynamic_index": entry_dynamic_index,
                    "dynamic_prefix": cached_scope.get("dynamic_prefix"),
                    "label_path": entry_anchor_path,
                    "source": cached_scope.get("mode") or "receipt-header-scope-cache",
                }
            else:
                finance_scope = timings.measure(
                    "header.finance-org-fast-scope",
                    run_with_jab_lock,
                    jab_lock,
                    find_finance_org_header_scope_by_paths,
                    jab,
                    entry_scope_hwnd,
                    preferred_dynamic_index=preferred_cached_dynamic_index,
                    min_index=1,
                    max_index=10,
                )
                row_report["entry_finance_org_fast_scope"] = {
                    **finance_scope,
                    "purpose": (
                        "开单快速确认只提供当前 Canvas；优先用财务组织(O)"
                        "稳定 path 解析表头 dynamic_index"
                    )
                }
                if finance_scope.get("ok"):
                    entry_dynamic_index = finance_scope.get("dynamic_index")
                    entry_dynamic_index_source = "finance-org-fast-scope"
                    finance_semantic_label_path = finance_scope.get("semantic_label_path")
                    finance_label_path = finance_scope.get("label_path")
                    entry_anchor_path = (
                        finance_label_path
                        or finance_semantic_label_path
                        or entry_anchor_path
                    )
                    if entry_dynamic_index is not None:
                        cache_receipt_header_scope(
                            jab,
                            header_scope_cache,
                            {
                                "ok": True,
                                "scope_hwnd": entry_scope_hwnd,
                                "mode": "finance-org-fast-scope",
                                "dynamic_index": entry_dynamic_index,
                                "dynamic_prefix": finance_scope.get("dynamic_prefix"),
                                "matched_labels": [HEADER_SCOPE_ANCHOR_LABEL],
                                "semantic_label_path": (
                                    finance_semantic_label_path
                                    or finance_label_path
                                    or entry_anchor_path
                                ),
                                "label_path": (
                                    finance_label_path
                                    or finance_semantic_label_path
                                    or entry_anchor_path
                                ),
                                "text_path": finance_scope.get("text_path"),
                                "variant": finance_scope.get("variant"),
                            },
                        )
        row_report["entry_scope_hwnd"] = entry_scope_hwnd
        row_report["entry_dynamic_index"] = entry_dynamic_index
        row_report["entry_dynamic_index_source"] = entry_dynamic_index_source
        row_report["entry_anchor_path"] = entry_anchor_path
        row_report["locator_policy"] = {
            "header": (
                "财务组织用于确认当前 canvas/scope 并缓存语义锚点；其它表头字段优先复用"
                "该 scope 做容器内标签定位，失败才单字段语义兜底"
            ),
            "body": "明细表优先复用已定位表格，必要时按表格语义扫描重新定位",
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
                trust_provided_scope=True,
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
                trust_provided_scope=True,
                recover_after_failure=recover_modal_after_failure,
            )
        if header_pause_reports:
            row_report["header_pause_diagnostics"] = header_pause_reports
        row_report["header_steps"] = header_steps
        cached_after_header = getattr(jab, "_receipt_header_scope_cache", None) or {}
        if cached_after_header.get("ok"):
            cache_receipt_header_scope(jab, header_scope_cache, cached_after_header)
        if any(not step.get("ok") for step in header_steps):
            header_error = summarize_header_failure(header_steps)
            _event("header-fill-failed", excel_row=row.row, error=header_error)
            return fail(row_report, "header-fill", timings, header_error)
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
        counterparty_report = timings.measure(
            "header.counterparty-type",
            run_with_jab_lock,
            jab_lock,
            ensure_header_counterparty_customer,
            jab,
            entry_dynamic_index,
            entry_scope_hwnd,
            located=located,
            recover_after_failure=recover_modal_after_failure,
        )
        row_report["header_counterparty"] = counterparty_report
        if not counterparty_report.get("ok"):
            _event(
                "header-counterparty-failed",
                excel_row=row.row,
                error=counterparty_report.get("reason"),
            )
            return fail(
                row_report,
                "header-counterparty-type",
                timings,
                counterparty_report.get("reason") or "往来对象未确认",
            )
        extra_text_report = timings.measure(
            "header.extra-text-fields",
            run_with_jab_lock,
            jab_lock,
            write_extra_text_fields,
            jab,
            row.extra_text_fields,
            entry_dynamic_index,
            entry_scope_hwnd,
            recover_after_failure=recover_modal_after_failure,
        )
        row_report["extra_text_fields"] = extra_text_report
        if not extra_text_report.get("ok"):
            _event(
                "header-extra-text-failed",
                excel_row=row.row,
                error=extra_text_report.get("reason"),
            )
            return fail(
                row_report,
                "header-extra-text-fields",
                timings,
                extra_text_report.get("reason") or "扩展文本字段写入失败",
            )
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
        pipeline_text_task_ids = []
        pipeline_snapshot_task_ids = []
        pipeline_row_count_task_id = None

        for field_report in extra_text_report.get("fields") or []:
            if not field_report.get("ok") or not field_report.get("value"):
                continue
            if not field_report.get("path"):
                return fail(
                    row_report,
                    "header-extra-text-fields",
                    timings,
                    f"扩展文本字段 {field_report.get('label')!r} 缺少后台验证 path",
                )
            pipeline_text_task_ids.append(
                pipeline_verifier.submit_path_text(
                    field_report.get("label"),
                    field_report.get("path"),
                    field_report.get("value"),
                    scope_hwnd=entry_scope_hwnd,
                )
            )
        row_report["extra_text_verify_tasks"] = pipeline_text_task_ids

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
                scope_hwnd=entry_scope_hwnd,
                defer_wait=True,
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
        pipeline_wait_ids.extend(pipeline_text_task_ids)
        pipeline_wait_started = time.perf_counter()
        row_report["detail_pipeline_verify"] = pipeline_verifier.wait(
            pipeline_wait_ids,
            timeout=2.0,
        )
        row_report["detail_pipeline_state"] = verifier_snapshot(pipeline_verifier)
        timings.add(
            "detail.pipeline-final-wait",
            time.perf_counter() - pipeline_wait_started,
        )
        extra_text_failures = []
        if pipeline_text_task_ids:
            pipeline_results = row_report["detail_pipeline_verify"].get("results") or {}
            for task_id in pipeline_text_task_ids:
                result = pipeline_results.get(task_id)
                if result is None or not result.get("ok"):
                    extra_text_failures.append(
                        {
                            "task_id": task_id,
                            "result": result,
                        }
                    )
        if extra_text_failures:
            row_report["extra_text_verify_failures"] = extra_text_failures
            _event(
                "extra-text-verify-failed",
                excel_row=row.row,
                error="扩展文本字段后台验证未通过",
            )
            return fail(
                row_report,
                "detail-extra-text-verify",
                timings,
                "扩展文本字段后台验证未通过",
            )
        if diagnose_detail_repair:
            row_report["detail_pipeline_verify_before_repair_drill"] = dict(
                row_report["detail_pipeline_verify"]
            )
            row_report["detail_pipeline_verify"] = force_one_detail_field_pending(
                row_report["detail_pipeline_verify"],
                pipeline_field_task_ids,
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
        guard_source = (
            row_report.get("detail_pipeline_verify_after_repair")
            or row_report.get("detail_pipeline_verify")
        )
        main_row_cells = latest_snapshot_row_cells(guard_source, row_index=0)
        exchange_rate_guard = validate_exchange_rate_not_polluted(
            main_row_cells,
            row.currency,
            row.raw_amount,
            row_index=0,
        )
        row_report["detail_exchange_rate_guard"] = exchange_rate_guard
        if not exchange_rate_guard.get("ok"):
            _event(
                "exchange-rate-guard-failed",
                excel_row=row.row,
                error=exchange_rate_guard.get("reason"),
            )
            return fail(
                row_report,
                "detail-exchange-rate-guard",
                timings,
                exchange_rate_guard.get("reason") or "汇率列污染守卫未通过",
            )
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
            row_report["save_attempted"] = True
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
        attach_slow_step_summary(row_report, timings)
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


def latest_snapshot_row_cells(report, row_index=0):
    for item in reversed((report or {}).get("snapshots") or []):
        snapshot = item.get("snapshot") or {}
        if not snapshot.get("ok"):
            continue
        for row in snapshot.get("rows") or []:
            try:
                current_index = int(row.get("row_index"))
            except (TypeError, ValueError):
                continue
            if current_index == int(row_index):
                return row.get("cells") or {}
    for result in reversed(list(((report or {}).get("results") or {}).values())):
        snapshot = result.get("snapshot") or {}
        if not snapshot.get("ok"):
            continue
        for row in snapshot.get("rows") or []:
            try:
                current_index = int(row.get("row_index"))
            except (TypeError, ValueError):
                continue
            if current_index == int(row_index):
                return row.get("cells") or {}
    return {}


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
    cached = _BODY_TABLE_SUFFIX_CACHE.get(body_table_cache_key(scope_hwnd))
    if not cached:
        cached = _BODY_TABLE_SUFFIX_CACHE.get("last")
    suffix = (cached or {}).get("suffix")
    if not suffix:
        return None
    path = f"{receipt_header_dynamic_prefix(dynamic_index)}.{suffix}"
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
    if cached:
        located = locate_receipt_body_table_cached(
            jab,
            cached=cached,
            max_rows=5,
            scope_hwnd=scope_hwnd,
        )
        if located.get("cache_hit"):
            return {
                **located,
                "source": "learned-body-table-path",
                "cached_path": (cached or {}).get("best"),
            }

    located = locate_receipt_body_table_cached(
        jab,
        cached=None,
        max_rows=5,
        scope_hwnd=scope_hwnd,
    )
    learned = cache_body_table_suffix(
        dynamic_index,
        scope_hwnd,
        ((located or {}).get("best") or {}).get("path"),
    )
    return {
        **located,
        "source": "semantic-body-table-scan",
        "cached_path": (cached or {}).get("best"),
        "learned_suffix": learned,
    }


def cache_body_table_suffix(dynamic_index, scope_hwnd, path):
    prefix = receipt_header_dynamic_prefix(dynamic_index)
    if not prefix or not path or not str(path).startswith(f"{prefix}."):
        return None
    suffix = str(path)[len(prefix) + 1 :]
    cached = {
        "dynamic_index": dynamic_index,
        "scope_hwnd": scope_hwnd,
        "suffix": suffix,
        "path": path,
        "source": "semantic-body-table-scan",
    }
    _BODY_TABLE_SUFFIX_CACHE[body_table_cache_key(scope_hwnd)] = cached
    _BODY_TABLE_SUFFIX_CACHE["last"] = cached
    return cached


def body_table_cache_key(scope_hwnd):
    return int(scope_hwnd) if scope_hwnd is not None else None


def build_header_scope_for_followup(scope_hwnd, dynamic_index):
    if not scope_hwnd or dynamic_index is None:
        return None
    return {
        "ok": True,
        "scope_hwnd": scope_hwnd,
        "dynamic_index": dynamic_index,
        "mode": "provided-canvas-anchor",
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


def apply_excel_text_field_mappings(config, raw_mappings):
    mappings = []
    for raw in raw_mappings or []:
        text = str(raw or "").strip()
        if not text:
            continue
        if "=" not in text:
            raise SystemExit("--excel-text-field-map 格式必须是 EXCEL列名=NC文本名")
        excel_column, nc_field = [part.strip() for part in text.split("=", 1)]
        if not excel_column or not nc_field:
            raise SystemExit("--excel-text-field-map 两侧都不能为空")
        mappings.append({"excel_column": excel_column, "nc_field": nc_field})
    if mappings:
        config.setdefault("receipt_entry", {})["excel_text_field_mappings"] = mappings


def apply_excel_column_overrides(config, raw_overrides):
    allowed = {
        "date_column",
        "payer_name_column",
        "raw_amount_column",
        "bank_column",
        "currency_column",
        "customer_code_column",
        "fee_column",
        "organization_column",
    }
    excel_cfg = config.setdefault("receipt_entry", {}).setdefault("excel", {})
    for raw in raw_overrides or []:
        text = str(raw or "").strip()
        if not text:
            continue
        if "=" not in text:
            raise SystemExit("--excel-column 格式必须是 配置键=EXCEL列名")
        key, excel_column = [part.strip() for part in text.split("=", 1)]
        if key not in allowed:
            raise SystemExit(f"--excel-column 不支持配置键 {key!r}")
        if not excel_column:
            raise SystemExit("--excel-column 的 EXCEL列名不能为空")
        excel_cfg[key] = excel_column


def write_extra_text_fields(
    jab,
    fields,
    dynamic_index,
    scope_hwnd=None,
    recover_after_failure=None,
):
    values = {
        str(label or "").strip(): str(value or "").strip()
        for label, value in (fields or {}).items()
        if str(label or "").strip() and str(value or "").strip()
    }
    if not values:
        return {"ok": True, "skipped": True, "fields": []}
    results = []
    for label, value in values.items():
        result = write_extra_text_field_by_dynamic_path(
            jab,
            label,
            value,
            dynamic_index,
            scope_hwnd=scope_hwnd,
            recover_after_failure=recover_after_failure,
        )
        results.append(result)
        if not result.get("ok"):
            return {
                "ok": False,
                "fields": results,
                "reason": f"未能写入 NC 文本字段 {label!r}",
            }
    return {"ok": True, "fields": results}


def ensure_header_counterparty_customer(
    jab,
    dynamic_index,
    scope_hwnd=None,
    located=None,
    recover_after_failure=None,
):
    started_at = time.perf_counter()
    if dynamic_index is None:
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "dynamic_index": dynamic_index,
            "reason": "往来对象 dynamic_index 未配置",
            "seconds": round(time.perf_counter() - started_at, 3),
        }

    detail = read_detail_counterparty_value(
        jab,
        dynamic_index,
        scope_hwnd=scope_hwnd,
        located=located,
        row=0,
        col=0,
    )
    detail_value = normalize_counterparty_value(
        detail.get("value"),
        detail.get("text"),
    )
    if detail_value == COUNTERPARTY_EXPECTED:
        return {
            "ok": True,
            "skipped": True,
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "actual": COUNTERPARTY_EXPECTED,
            "path": None,
            "dynamic_index": dynamic_index,
            "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
            "detail": detail,
            "state": {
                "state": COUNTERPARTY_STATE_OK,
                "actual": detail_value,
                "source": "detail-row0-col0",
                "repairable": False,
            },
            "source": "detail-row0-col0",
            "reason": "明细表第 0 行往来对象为客户，跳过",
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    if detail_value in COUNTERPARTY_KNOWN_OPTIONS:
        snapshot = {
            "combo": {},
            "embedded": {},
            "detail": detail,
            "selected": "",
            "combo_text": "",
            "detail_value": detail_value,
            "state": {
                "state": COUNTERPARTY_STATE_CONFLICT,
                "actual": detail_value,
                "source": "detail-row0-col0",
                "repairable": False,
            },
        }
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "actual": detail_value,
            "path": None,
            "dynamic_index": dynamic_index,
            "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
            "before": {},
            "embedded": {},
            "detail": detail,
            "state": snapshot["state"],
            "readback_trusted": False,
            "reason": summarize_counterparty_failure(snapshot),
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    if detail.get("ok") is False and not detail_value:
        reason = str(detail.get("reason") or "")
        unreadable = any(
            marker in reason
            for marker in ("未定位", "命中失败", "行列不足", "异常")
        )
        if unreadable:
            snapshot = {
                "combo": {},
                "embedded": {},
                "detail": detail,
                "selected": "",
                "combo_text": "",
                "detail_value": "",
                "state": {
                    "state": COUNTERPARTY_STATE_DETAIL_UNREADABLE,
                    "actual": "",
                    "source": "detail-row0-col0",
                    "repairable": False,
                },
            }
            return {
                "ok": False,
                "label": COUNTERPARTY_LABEL,
                "expected": COUNTERPARTY_EXPECTED,
                "actual": "",
                "path": None,
                "dynamic_index": dynamic_index,
                "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
                "before": {},
                "embedded": {},
                "detail": detail,
                "state": snapshot["state"],
                "readback_trusted": False,
                "reason": summarize_counterparty_failure(snapshot),
                "seconds": round(time.perf_counter() - started_at, 3),
            }

    found = find_counterparty_combo(jab, dynamic_index, scope_hwnd=scope_hwnd)
    found_path = found.get("path")
    recovery_after_find = None
    if not found.get("ok") and recover_after_failure is not None:
        recovery_after_find = recover_after_failure()
        if recovery_after_find.get("attempted") and recovery_after_find.get("ok"):
            found = find_counterparty_combo(jab, dynamic_index, scope_hwnd=scope_hwnd)
            found_path = found.get("path")
    if not found.get("ok"):
        return {
            **found,
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "dynamic_index": dynamic_index,
            "detail": detail,
            "modal_recovery": recovery_after_find,
            "seconds": round(time.perf_counter() - started_at, 3),
        }

    try:
        combo = read_counterparty_combo_state(
            jab,
            found["vm_id"],
            found["context"],
        )
        embedded = read_counterparty_selected_option(
            jab,
            found["vm_id"],
            found["context"],
        )
        snapshot = {
            "combo": combo,
            "embedded": embedded,
            "detail": detail,
            "selected": normalize_counterparty_value(embedded.get("selected")),
            "combo_text": normalize_counterparty_value(
                combo.get("description"),
                combo.get("text"),
                combo.get("name"),
            ),
            "detail_value": detail_value,
        }
        state = classify_counterparty_snapshot(snapshot)
        snapshot["state"] = state
        if state["state"] == COUNTERPARTY_STATE_OK:
            return {
                "ok": True,
                "skipped": True,
                "label": COUNTERPARTY_LABEL,
                "expected": COUNTERPARTY_EXPECTED,
                "actual": COUNTERPARTY_EXPECTED,
                "path": found_path,
                "dynamic_index": dynamic_index,
                "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
                "before": snapshot["combo"],
                "embedded": snapshot["embedded"],
                "detail": snapshot["detail"],
                "state": state,
                "source": "detail-row0-col0",
                "reason": "明细表第 0 行往来对象为客户，跳过",
                "seconds": round(time.perf_counter() - started_at, 3),
            }

        if state["state"] == COUNTERPARTY_STATE_REPAIRABLE:
            repair = select_counterparty_customer_embedded(
                jab,
                found["vm_id"],
                found["context"],
                press_enter=True,
            )
            time.sleep(0.12)
            after_detail = read_detail_counterparty_value(
                jab,
                dynamic_index,
                scope_hwnd=scope_hwnd,
                located=located,
                row=0,
                col=0,
            )
            after_value = normalize_counterparty_value(
                after_detail.get("value"),
                after_detail.get("text"),
            )
            if after_value == COUNTERPARTY_EXPECTED:
                return {
                    "ok": True,
                    "repaired": True,
                    "label": COUNTERPARTY_LABEL,
                    "expected": COUNTERPARTY_EXPECTED,
                    "actual": after_value,
                    "path": found_path,
                    "dynamic_index": dynamic_index,
                    "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
                    "before": snapshot["combo"],
                    "embedded": snapshot["embedded"],
                    "detail": snapshot["detail"],
                    "after_detail": after_detail,
                    "repair": repair,
                    "state": state,
                    "source": "embedded-selection-api",
                    "reason": "往来对象为空，已通过子列表 selection API 选择客户并验证明细表",
                    "seconds": round(time.perf_counter() - started_at, 3),
                }
            return {
                "ok": False,
                "label": COUNTERPARTY_LABEL,
                "expected": COUNTERPARTY_EXPECTED,
                "actual": after_value,
                "path": found_path,
                "dynamic_index": dynamic_index,
                "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
                "before": snapshot["combo"],
                "embedded": snapshot["embedded"],
                "detail": snapshot["detail"],
                "after_detail": after_detail,
                "repair": repair,
                "state": state,
                "readback_trusted": False,
                "reason": summarize_counterparty_failure(snapshot, after_detail),
                "seconds": round(time.perf_counter() - started_at, 3),
            }

        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "expected": COUNTERPARTY_EXPECTED,
            "actual": state.get("actual") or "",
            "path": found_path,
            "dynamic_index": dynamic_index,
            "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
            "before": snapshot["combo"],
            "embedded": snapshot["embedded"],
            "detail": snapshot["detail"],
            "state": state,
            "readback_trusted": False,
            "reason": summarize_counterparty_failure(snapshot),
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    finally:
        jab.release_contexts(found["vm_id"], found["owned_contexts"])


def read_counterparty_snapshot_from_found(jab, found, dynamic_index, scope_hwnd=None):
    combo = read_counterparty_combo_state(
        jab,
        found["vm_id"],
        found["context"],
    )
    embedded = read_counterparty_selected_option(
        jab,
        found["vm_id"],
        found["context"],
    )
    detail = read_detail_counterparty_value(
        jab,
        dynamic_index,
        scope_hwnd=scope_hwnd,
        row=0,
        col=0,
    )
    return {
        "combo": combo,
        "embedded": embedded,
        "detail": detail,
        "selected": normalize_counterparty_value(embedded.get("selected")),
        "combo_text": normalize_counterparty_value(
            combo.get("description"),
            combo.get("text"),
            combo.get("name"),
        ),
        "detail_value": normalize_counterparty_value(
            detail.get("value"),
            detail.get("text"),
        ),
    }


def classify_counterparty_snapshot(snapshot):
    selected = (snapshot or {}).get("selected") or ""
    combo_text = (snapshot or {}).get("combo_text") or ""
    detail_value = (snapshot or {}).get("detail_value") or ""
    detail = (snapshot or {}).get("detail") or {}

    if detail_value == COUNTERPARTY_EXPECTED:
        return {
            "state": COUNTERPARTY_STATE_OK,
            "actual": detail_value,
            "source": "detail-row0-col0",
            "repairable": False,
        }

    for source, value in (
        ("detail-row0-col0", detail_value),
        ("combo-text", combo_text),
        ("embedded-selected-option", selected),
    ):
        if value in COUNTERPARTY_KNOWN_OPTIONS and value != COUNTERPARTY_EXPECTED:
            return {
                "state": COUNTERPARTY_STATE_CONFLICT,
                "actual": value,
                "source": source,
                "repairable": False,
            }

    if detail.get("ok") is False and not detail_value:
        reason = str(detail.get("reason") or "")
        unreadable = any(
            marker in reason
            for marker in ("未定位", "命中失败", "行列不足", "异常")
        )
        if unreadable:
            return {
                "state": COUNTERPARTY_STATE_DETAIL_UNREADABLE,
                "actual": "",
                "source": "detail-row0-col0",
                "repairable": False,
            }

    return {
        "state": COUNTERPARTY_STATE_REPAIRABLE,
        "actual": detail_value,
        "source": "detail-row0-col0",
        "repairable": True,
    }


def normalize_counterparty_value(*values):
    for value in values:
        text = str(value or "").strip()
        if text in COUNTERPARTY_KNOWN_OPTIONS:
            return text
    return ""


def select_counterparty_customer_embedded(jab, vm_id, combo_context, press_enter=True):
    target = find_counterparty_embedded_list(jab, vm_id, combo_context)
    try:
        if not target.get("ok"):
            return {
                "ok": False,
                "method": "embedded-selection-api",
                "reason": target.get("reason") or "往来对象子列表未找到",
            }
        if not hasattr(jab.dll, "addAccessibleSelectionFromContext"):
            return {
                "ok": False,
                "method": "embedded-selection-api",
                "reason": "JAB selection API unavailable",
                "target": embedded_counterparty_target_summary(target),
            }
        list_context = target["list_context"]
        customer_index = next(
            (
                int(item.get("index"))
                for item in target.get("labels") or []
                if item.get("name") == COUNTERPARTY_EXPECTED
            ),
            None,
        )
        if customer_index is None:
            return {
                "ok": False,
                "method": "embedded-selection-api",
                "reason": "往来对象子列表没有客户选项",
                "target": embedded_counterparty_target_summary(target),
            }
        if hasattr(jab.dll, "clearAccessibleSelectionFromContext"):
            jab.dll.clearAccessibleSelectionFromContext(vm_id, list_context)
        selected_ok = bool(
            jab.dll.addAccessibleSelectionFromContext(
                vm_id,
                list_context,
                customer_index,
            )
        )
        focus_ok = request_focus_context(jab, vm_id, list_context)
        enter_ok = None
        if press_enter:
            enter_ok = press_counterparty_commit_keys(jab)
        return {
            "ok": bool(selected_ok),
            "method": "embedded-selection-api",
            "selected_index": customer_index,
            "target": embedded_counterparty_target_summary(target),
            "request_focus_list": focus_ok,
            "commit": enter_ok,
        }
    except Exception as exc:
        return {
            "ok": False,
            "method": "embedded-selection-api",
            "error": repr(exc),
        }
    finally:
        if target.get("owned_contexts"):
            jab.release_contexts(vm_id, target["owned_contexts"])


def request_focus_context(jab, vm_id, context):
    if not hasattr(jab, "dll") or not hasattr(jab.dll, "requestFocus"):
        return {"ok": None, "reason": "requestFocus unavailable"}
    try:
        return {"ok": bool(jab.dll.requestFocus(vm_id, context))}
    except Exception as exc:
        return {"ok": False, "error": repr(exc)}


def press_counterparty_commit_keys(jab):
    sent = []
    try:
        jab.press_key("home", wait=0.02)
        sent.append("home")
        jab.press_key("enter", wait=0)
        sent.append("enter")
        return {"ok": True, "keys": sent}
    except Exception as exc:
        return {"ok": False, "keys": sent, "error": repr(exc)}


def embedded_counterparty_target_summary(target):
    return {
        "list": target.get("list"),
        "popup": target.get("popup"),
        "labels": [
            {
                "index": item.get("index"),
                "name": item.get("name"),
                "states": item.get("states"),
            }
            for item in target.get("labels") or []
        ],
    }


def first_non_empty_counterparty_text(*values):
    return normalize_counterparty_value(*values)


def read_detail_counterparty_value(
    jab,
    dynamic_index,
    scope_hwnd=None,
    row=0,
    col=0,
    located=None,
):
    try:
        if located is None:
            located = resolve_body_table_by_dynamic_prefix(
                jab,
                dynamic_index,
                scope_hwnd=scope_hwnd,
            )
        best = (located or {}).get("best") or {}
        path = best.get("path")
        if not path:
            return {
                "ok": False,
                "source": "detail-row0-col0",
                "row": row,
                "col": col,
                "located": slim_counterparty_located(located),
                "reason": "明细表 path 未定位",
            }

        window = best.get("window") or {}
        context, vm_id, owned, window_info = jab.find_context_by_path_once(
            path,
            class_name=window.get("class_name") or "SunAwtCanvas",
            scope_hwnd=scope_hwnd or window.get("hwnd"),
            role="table",
            require_showing=False,
            require_valid_bounds=False,
        )
        if not context:
            return {
                "ok": False,
                "source": "detail-row0-col0",
                "row": row,
                "col": col,
                "path": path,
                "located": slim_counterparty_located(located),
                "reason": "明细表 path 命中失败",
            }
        try:
            table_info = jab.get_table_info(vm_id, context)
            table = {
                "path": path,
                "window": window_info or window,
                "row_count": int(getattr(table_info, "rowCount", 0) or 0)
                if table_info
                else None,
                "col_count": int(getattr(table_info, "columnCount", 0) or 0)
                if table_info
                else None,
            }
            schema = detail_table_schema_snapshot(best)
            table["schema"] = schema
            if table.get("col_count") is not None and table.get("col_count") < 12:
                return {
                    "ok": False,
                    "source": "detail-row0-col0",
                    "row": row,
                    "col": col,
                    "path": path,
                    "table": table,
                    "located": slim_counterparty_located(located),
                    "reason": "明细表列数不足，不像收款单明细表",
                }
            if table_info and (
                int(table_info.rowCount) <= row or int(table_info.columnCount) <= col
            ):
                return {
                    "ok": False,
                    "source": "detail-row0-col0",
                    "row": row,
                    "col": col,
                    "path": path,
                    "table": table,
                    "located": slim_counterparty_located(located),
                    "reason": "明细表行列不足，无法读取往来对象",
                }
            text, is_selected = jab.get_table_cell_text_and_selection(
                vm_id,
                context,
                row,
                col,
            )
            value = first_non_empty_counterparty_text(text)
            return {
                "ok": bool(value),
                "source": "detail-row0-col0",
                "row": row,
                "col": col,
                "value": value,
                "text": str(text or "").strip(),
                "is_selected": bool(is_selected),
                "path": path,
                "table": table,
                "located": slim_counterparty_located(located),
                "reason": None if value else "明细表往来对象单元格为空",
            }
        finally:
            jab.release_contexts(vm_id, owned)
    except Exception as exc:
        return {
            "ok": False,
            "source": "detail-row0-col0",
            "row": row,
            "col": col,
            "reason": "读取明细表往来对象异常",
            "error": repr(exc),
        }


def slim_counterparty_located(located):
    if not located:
        return None
    best = (located or {}).get("best") or {}
    return {
        "cache_hit": bool(located.get("cache_hit")),
        "fallback_used": bool(located.get("fallback_used")),
        "source": located.get("source"),
        "path": best.get("path"),
        "window": best.get("window"),
        "row_count": best.get("row_count"),
        "col_count": best.get("col_count"),
        "reason": located.get("reason"),
    }


def detail_table_schema_snapshot(best):
    rows = (best or {}).get("rows") or []
    first = rows[0] if rows else {}
    cells = (first or {}).get("cells") or first.get("values") or first
    if not isinstance(cells, dict):
        cells = {}
    key_cells = {
        str(index): str(cells.get(str(index), cells.get(index, "")) or "").strip()
        for index in (0, 1, 2, 3, 4, 5, 7, 11)
    }
    return {
        "row0_key_cells": key_cells,
        "looks_like_receipt_detail": (
            key_cells.get("0") in {"", *COUNTERPARTY_KNOWN_OPTIONS}
            and bool(key_cells.get("1") or key_cells.get("2") or key_cells.get("5"))
        ),
    }


def summarize_counterparty_failure(snapshot, after_detail=None):
    snapshot = snapshot or {}
    detail = after_detail or snapshot.get("detail") or {}
    detail_value = normalize_counterparty_value(
        detail.get("value"),
        detail.get("text"),
    )
    raw_detail = (detail or {}).get("text") or ""
    detail_reason = (detail or {}).get("reason") or ""
    state = snapshot.get("state") or {}
    parts = [
        f"header_selected={snapshot.get('selected') or ''}",
        f"combo_text={snapshot.get('combo_text') or ''}",
        f"detail_row0_col0={detail_value}",
    ]
    if raw_detail and raw_detail != detail_value:
        parts.append(f"detail_raw={raw_detail}")
    if detail_reason:
        parts.append(f"detail_reason={detail_reason}")
    if state.get("state"):
        parts.append(f"state={state.get('state')}")
    return f"往来对象未确认客户；{'; '.join(parts)}；已禁用旧下拉键盘方案"


def read_counterparty_selected_option(jab, vm_id, combo_context):
    found = find_counterparty_embedded_list(jab, vm_id, combo_context)
    try:
        if not found.get("ok"):
            return found
        labels = found.get("labels") or []
        selected = next(
            (
                item.get("name")
                for item in labels
                if "selected" in str(item.get("states") or "").lower()
            ),
            "",
        )
        return {
            "ok": True,
            "selected": selected,
            "options": [item.get("name") for item in labels if item.get("name")],
            "list": found.get("list"),
            "popup": found.get("popup"),
        }
    finally:
        if found.get("owned_contexts"):
            jab.release_contexts(vm_id, found["owned_contexts"])


def find_counterparty_embedded_list(jab, vm_id, combo_context):
    result = {
        "ok": False,
        "reason": "往来对象子列表未找到",
        "owned_contexts": [],
    }
    best = None

    def visit(context, path, depth, ancestors):
        nonlocal best
        info = jab.get_context_info(vm_id, context)
        if not info:
            return
        role = (info.role_en_US.strip() or info.role.strip()).lower()
        owned = []
        children = []
        if depth > 0:
            for index in range(min(info.childrenCount, getattr(jab, "max_children", 1000))):
                child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
                if not child:
                    continue
                owned.append(child)
                child_info = jab.get_context_info(vm_id, child)
                if not child_info:
                    continue
                child_role = (
                    child_info.role_en_US.strip() or child_info.role.strip()
                ).lower()
                child_name = child_info.name.strip()
                children.append((index, child, child_info, child_role, child_name))

        if role == "list":
            label_items = [
                {
                    "index": index,
                    "path": f"{path}.{index}",
                    "name": child_name,
                    "description": child_info.description.strip(),
                    "role": child_info.role_en_US.strip() or child_info.role.strip(),
                    "states": child_info.states_en_US.strip()
                    or child_info.states.strip(),
                    "bounds": [
                        child_info.x,
                        child_info.y,
                        child_info.width,
                        child_info.height,
                    ],
                }
                for index, _child, child_info, child_role, child_name in children
                if child_role == "label"
            ]
            names = [item["name"] for item in label_items if item.get("name")]
            if COUNTERPARTY_EXPECTED in names:
                popup = next(
                    (
                        info_to_counterparty_dict(ancestor_info, "ancestor")
                        for _ancestor_context, ancestor_info in reversed(ancestors)
                        if (
                            ancestor_info.role_en_US.strip()
                            or ancestor_info.role.strip()
                        ).lower()
                        == "popup menu"
                    ),
                    None,
                )
                keep = [context]
                keep.extend(child for _index, child, *_rest in children)
                best = {
                    "ok": True,
                    "list": info_to_counterparty_dict(info, path),
                    "list_context": context,
                    "popup": popup,
                    "labels": label_items,
                    "owned_contexts": unique_contexts(keep + owned),
                }
                return

        try:
            if depth > 0 and best is None:
                next_ancestors = ancestors + [(context, info)]
                for index, child, _child_info, _child_role, _child_name in children:
                    visit(child, f"{path}.{index}", depth - 1, next_ancestors)
                    if best is not None:
                        break
        finally:
            if best is None:
                jab.release_contexts(vm_id, owned)

    visit(combo_context, "target", 8, [])
    return best or result


def info_to_counterparty_dict(info, path):
    if not info:
        return None
    return {
        "path": path,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "role": info.role_en_US.strip() or info.role.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "bounds": [info.x, info.y, info.width, info.height],
        "children_count": info.childrenCount,
    }


def unique_contexts(contexts):
    result = []
    seen = set()
    for context in contexts or []:
        key = context_key(context)
        if key in seen:
            continue
        seen.add(key)
        result.append(context)
    return result


def context_key(context):
    try:
        value = getattr(context, "value", context)
        return ("int", int(value))
    except Exception:
        return ("repr", repr(context))


def find_counterparty_combo(jab, dynamic_index, scope_hwnd=None):
    prefix = receipt_header_dynamic_prefix(dynamic_index)
    cached_path = build_cached_counterparty_nearby_path(dynamic_index, scope_hwnd)
    if cached_path:
        cached = find_counterparty_combo_by_path(
            jab,
            cached_path,
            scope_hwnd=scope_hwnd,
        )
        if cached.get("ok"):
            return {
                **cached,
                "source": "nearby-cache-path",
                "cached_path": cached_path,
            }

    nearby = find_counterparty_combo_nearby(
        jab,
        dynamic_index,
        scope_hwnd=scope_hwnd,
    )
    if nearby.get("ok"):
        cache_counterparty_nearby_suffix(
            dynamic_index,
            scope_hwnd,
            prefix,
            nearby.get("path"),
        )
        return nearby

    return {
        "ok": False,
        "label": COUNTERPARTY_LABEL,
        "source": "not-found",
        "nearby_attempt": slim_found(nearby),
        "reason": nearby.get("reason") or "往来对象 nearby 定位失败",
    }


def build_cached_counterparty_nearby_path(dynamic_index, scope_hwnd=None):
    prefix = receipt_header_dynamic_prefix(dynamic_index)
    if not prefix:
        return None
    cached = _COUNTERPARTY_NEARBY_SUFFIX_CACHE.get(counterparty_cache_key(scope_hwnd))
    if not cached:
        cached = _COUNTERPARTY_NEARBY_SUFFIX_CACHE.get("last")
    suffix = (cached or {}).get("suffix")
    if not suffix:
        return None
    return f"{prefix}.{suffix}"


def cache_counterparty_nearby_suffix(dynamic_index, scope_hwnd, prefix, path):
    if not prefix or not path or not str(path).startswith(f"{prefix}."):
        return None
    suffix = str(path)[len(prefix) + 1 :]
    cached = {
        "dynamic_index": dynamic_index,
        "scope_hwnd": scope_hwnd,
        "suffix": suffix,
        "path": path,
        "source": "nearby",
    }
    _COUNTERPARTY_NEARBY_SUFFIX_CACHE[counterparty_cache_key(scope_hwnd)] = cached
    _COUNTERPARTY_NEARBY_SUFFIX_CACHE["last"] = cached
    return cached


def counterparty_cache_key(scope_hwnd):
    return int(scope_hwnd) if scope_hwnd is not None else None


def find_counterparty_combo_nearby(jab, dynamic_index, scope_hwnd=None):
    if scope_hwnd is None:
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "reason": "nearby 定位缺少 scope_hwnd",
        }
    dll = getattr(jab, "dll", None)
    if not dll or not hasattr(dll, "isJavaWindow") or not dll.isJavaWindow(scope_hwnd):
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "scope_hwnd": scope_hwnd,
            "reason": "nearby scope 不是 Java 窗口",
        }

    vm_id = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        int(scope_hwnd),
        ctypes.byref(vm_id),
        ctypes.byref(root_context),
    ):
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "scope_hwnd": scope_hwnd,
            "reason": "nearby 读取 scope root 失败",
        }

    controls = []
    owned = []
    selected_contexts = set()
    try:
        collect_counterparty_controls_for_bounds_scan(
            jab,
            vm_id.value,
            root_context.value,
            controls,
            owned,
            require_showing=True,
            depth=0,
        )
        labels = [
            (context, info, _path)
            for context, info, _path in controls
            if control_role(info) == "label"
            and info.name.strip() == COUNTERPARTY_LABEL
            and jab.context_info_has_valid_bounds(info)
        ]
        labels.sort(key=lambda item: (item[1].y, item[1].x))
        prefix = receipt_header_dynamic_prefix(dynamic_index)
        candidates = []
        for label_context, label_info, _label_path in labels:
            label_mid_y = label_info.y + label_info.height / 2
            label_right = label_info.x + label_info.width
            row_candidates = []
            for context, info, control_path in controls:
                if control_role(info) != "combo box":
                    continue
                if not jab.context_info_has_valid_bounds(info):
                    continue
                mid_y = info.y + info.height / 2
                right_distance = info.x - label_right
                dy = abs(mid_y - label_mid_y)
                if right_distance <= 0:
                    continue
                if dy > COUNTERPARTY_NEARBY_MAX_VERTICAL_DISTANCE:
                    continue
                if right_distance > COUNTERPARTY_NEARBY_MAX_RIGHT_DISTANCE:
                    continue
                actions = jab.get_action_names(vm_id.value, context)
                row_candidates.append(
                    {
                        "context": context,
                        "info": info,
                        "path": control_path,
                        "score": (right_distance, dy),
                        "actions": actions,
                    }
                )
            row_candidates.sort(key=lambda item: item["score"])
            for item in row_candidates:
                candidates.append(
                    {
                        "label": jab.info_to_dict(label_info),
                        "control": jab.info_to_dict(item["info"]),
                        "path": item["path"],
                        "actions": item["actions"],
                        "score": list(item["score"]),
                    }
                )
            if row_candidates:
                target = row_candidates[0]
                selected_contexts = {target["context"], label_context}
                release = [
                    context
                    for context in owned
                    if context not in selected_contexts
                ]
                jab.release_contexts(vm_id.value, release)
                return {
                    "ok": True,
                    "label": COUNTERPARTY_LABEL,
                    "source": "nearby",
                    "context": target["context"],
                    "vm_id": vm_id.value,
                    "owned_contexts": list(selected_contexts),
                    "window": {
                        "hwnd": int(scope_hwnd),
                        "class_name": "SunAwtCanvas",
                    },
                    "path": target["path"],
                    "dynamic_prefix": prefix,
                    "target": {
                        "label": jab.info_to_dict(label_info),
                        "control": jab.info_to_dict(target["info"]),
                        "actions": target["actions"],
                        "score": list(target["score"]),
                    },
                    "candidate_count": len(candidates),
                    "candidates": candidates[:8],
                }
    finally:
        if not selected_contexts:
            jab.release_contexts(vm_id.value, owned)

    return {
        "ok": False,
        "label": COUNTERPARTY_LABEL,
        "source": "nearby",
        "scope_hwnd": scope_hwnd,
        "candidate_count": len(candidates) if "candidates" in locals() else 0,
        "reason": "未在往来对象标签右侧找到 combo box",
    }


def collect_counterparty_controls_for_bounds_scan(
    jab,
    vm_id,
    context,
    controls,
    owned,
    require_showing=True,
    depth=0,
    path="0",
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    role = control_role(info)
    if role == "table" or depth >= jab.max_depth:
        return

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_path = f"{path}.{index}"
        child_info = jab.get_context_info(vm_id, child)
        if not child_info:
            jab.release_contexts(vm_id, [child])
            continue

        owned.append(child)
        states = (
            child_info.states_en_US.strip() or child_info.states.strip()
        ).lower()
        showing = "visible" in states and "showing" in states
        if not require_showing or showing:
            controls.append((child, child_info, child_path))

        collect_counterparty_controls_for_bounds_scan(
            jab,
            vm_id,
            child,
            controls,
            owned,
            require_showing=require_showing,
            depth=depth + 1,
            path=child_path,
        )


def control_role(info):
    return (info.role_en_US.strip() or info.role.strip()).lower()


def slim_found(found):
    return {
        key: value
        for key, value in (found or {}).items()
        if key not in {"context", "vm_id", "owned_contexts", "candidates"}
    }


def find_counterparty_combo_by_path(jab, path, scope_hwnd=None):
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="combo box",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "label": COUNTERPARTY_LABEL,
            "path": path,
            "reason": "往来对象下拉控件未找到",
        }
    return {
        "ok": True,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "window": window_info,
        "path": path,
    }


def read_counterparty_combo_state(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    text = jab.get_text_context_value(vm_id, context)
    return {
        "name": info.name.strip() if info else "",
        "description": info.description.strip() if info else "",
        "text": str(text or "").strip(),
        "role": (info.role_en_US.strip() or info.role.strip()) if info else "",
        "states": (info.states_en_US.strip() or info.states.strip()) if info else "",
    }


def first_non_empty_text(*values):
    for value in values:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def write_extra_text_field_by_dynamic_path(
    jab,
    label,
    value,
    dynamic_index,
    scope_hwnd=None,
    recover_after_failure=None,
):
    started_at = time.perf_counter()
    recovery_after_find = None

    def find_field():
        path_template = get_receipt_header_path_template(dynamic_index)
        path_found = (
            find_receipt_header_field_by_dynamic_path(
                jab,
                label,
                dynamic_index,
                scope_hwnd=scope_hwnd,
                require_showing=True,
                require_valid_bounds=False,
                path_template=path_template,
            )
            if path_template
            else {
                "ok": False,
                "label": label,
                "dynamic_index": dynamic_index,
                "reason": "header path template not learned",
            }
        )
        if path_found.get("ok"):
            path_found["source"] = "learned-header-template"
            return path_found
        semantic_found = find_receipt_header_field_by_live_semantic(
            jab,
            label,
            scope_hwnd=scope_hwnd,
            include_scoped=False,
        )
        semantic_found["dynamic_path_attempt"] = path_found
        if semantic_found.get("ok"):
            semantic_found["source"] = "semantic-live-after-template-miss"
            semantic_found["dynamic_index"] = dynamic_index
            semantic_found["dynamic_prefix"] = receipt_header_dynamic_prefix(dynamic_index)
            inferred_template = infer_header_path_template_from_field(
                semantic_found.get("path"),
                dynamic_index,
                label,
            )
            if inferred_template:
                set_receipt_header_path_template(dynamic_index, inferred_template)
                semantic_found["header_path_template_learned"] = inferred_template
        return semantic_found

    try:
        found = find_field()
    except Exception as exc:
        if recover_after_failure is None:
            raise
        recovery_after_find = recover_after_failure()
        if recovery_after_find.get("attempted") and recovery_after_find.get("ok"):
            try:
                found = find_field()
            except Exception as retry_exc:
                return {
                    "ok": False,
                    "stage": "resolve",
                    "label": label,
                    "value": value,
                    "dynamic_index": dynamic_index,
                    "exception": f"{type(retry_exc).__name__}: {retry_exc}",
                    "first_exception": f"{type(exc).__name__}: {exc}",
                    "modal_recovery": recovery_after_find,
                    "seconds": round(time.perf_counter() - started_at, 3),
                }
        else:
            return {
                "ok": False,
                "stage": "resolve",
                "label": label,
                "value": value,
                "dynamic_index": dynamic_index,
                "exception": f"{type(exc).__name__}: {exc}",
                "modal_recovery": recovery_after_find,
                "seconds": round(time.perf_counter() - started_at, 3),
            }
    if not found.get("ok") and recover_after_failure is not None:
        recovery_after_find = recover_after_failure()
        if recovery_after_find.get("attempted") and recovery_after_find.get("ok"):
            found = find_field()
    path_attempt = found.get("dynamic_path_attempt") or found
    if not found.get("ok"):
        return {
            "ok": False,
            "stage": "resolve",
            "label": label,
            "value": value,
            "dynamic_index": dynamic_index,
            "dynamic_path_attempt": path_attempt,
            "semantic_attempt": found,
            "modal_recovery": recovery_after_find,
            "reason": found.get("reason") or "extra text field not found",
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    def write_current_context(initial_recovery=None):
        modal_recovery = initial_recovery
        attempts = []
        max_attempts = 2
        for attempt_no in range(1, max_attempts + 1):
            write_started_at = time.perf_counter()
            info_before = jab.get_context_info(vm_id, context)
            before = jab.get_text_context_value(vm_id, context)
            paste_started_at = time.perf_counter()
            paste = guarded_paste_header_value(
                jab,
                vm_id,
                context,
                found.get("window") or {},
                value,
            )
            paste_seconds = round(time.perf_counter() - paste_started_at, 3)
            if not paste.get("ok") and recover_after_failure is not None:
                recovery_after_set = recover_after_failure()
                modal_recovery = recovery_after_set or modal_recovery
                if recovery_after_set.get("attempted") and recovery_after_set.get("ok"):
                    paste_started_at = time.perf_counter()
                    paste = guarded_paste_header_value(
                        jab,
                        vm_id,
                        context,
                        found.get("window") or {},
                        value,
                    )
                    paste_seconds = round(time.perf_counter() - paste_started_at, 3)
            info_after = jab.get_context_info(vm_id, context)
            after = jab.get_text_context_value(vm_id, context)
            backend_state = describe_backend_field_state(info_after, after, value=value)
            accepted = bool(
                backend_state.get("accepted") or backend_state.get("written")
            )
            ok = bool(paste.get("ok") and accepted)
            attempt = {
                "ok": ok,
                "input_ok": bool(paste.get("ok")),
                "attempt": attempt_no,
                "stage": "write",
                "label": label,
                "value": value,
                "path": found.get("path"),
                "source": found.get("source") or "dynamic-path",
                "dynamic_index": dynamic_index,
                "dynamic_prefix": found.get("dynamic_prefix"),
                "dynamic_path_attempt": found.get("dynamic_path_attempt")
                or path_attempt,
                "inferred_suffix": found.get("inferred_suffix"),
                "modal_recovery": modal_recovery,
                "text_before": before,
                "text_after": after,
                "description_before": info_before.description.strip()
                if info_before
                else "",
                "description_after": info_after.description.strip()
                if info_after
                else "",
                "post_write_snapshot": backend_state,
                "guarded_paste": paste,
                "enter_ok": bool(paste.get("enter_ok")),
                "reason": None
                if ok
                else (
                    paste.get("reason")
                    or f"写入后未读回目标文本字段值：{label}={value!r}"
                ),
                "timing": {
                    "paste_seconds": paste_seconds,
                    "write_seconds": round(time.perf_counter() - write_started_at, 3),
                    "total_seconds": round(time.perf_counter() - started_at, 3),
                },
                "seconds": round(time.perf_counter() - started_at, 3),
            }
            attempts.append(attempt)
            if ok or not paste.get("ok"):
                final = dict(attempt)
                final["attempts"] = attempts
                final["rewrites"] = attempts[1:]
                return final
            if attempt_no < max_attempts:
                time.sleep(0.1)

        final = dict(attempts[-1])
        final["attempts"] = attempts
        final["rewrites"] = attempts[1:]
        return final

    try:
        try:
            return write_current_context(recovery_after_find)
        except Exception as exc:
            if recover_after_failure is None:
                raise
            recovery_after_exception = recover_after_failure()
            if recovery_after_exception.get(
                "attempted"
            ) and recovery_after_exception.get("ok"):
                try:
                    retried = write_current_context(recovery_after_exception)
                except Exception as retry_exc:
                    return {
                        "ok": False,
                        "stage": "write",
                        "label": label,
                        "value": value,
                        "path": found.get("path"),
                        "dynamic_index": dynamic_index,
                        "dynamic_prefix": found.get("dynamic_prefix"),
                        "exception": f"{type(retry_exc).__name__}: {retry_exc}",
                        "first_exception": f"{type(exc).__name__}: {exc}",
                        "modal_recovery": recovery_after_exception,
                        "seconds": round(time.perf_counter() - started_at, 3),
                    }
                retried["retried_after_modal_recovery"] = True
                return retried
            return {
                "ok": False,
                "stage": "write",
                "label": label,
                "value": value,
                "path": found.get("path"),
                "dynamic_index": dynamic_index,
                "dynamic_prefix": found.get("dynamic_prefix"),
                "exception": f"{type(exc).__name__}: {exc}",
                "modal_recovery": recovery_after_exception,
                "seconds": round(time.perf_counter() - started_at, 3),
            }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def read_customer_name_after_header(
    jab,
    header_steps,
    dynamic_index,
    scope_hwnd,
    timeout=1.2,
    poll_interval=0.1,
):
    step = next(
        (item for item in header_steps or [] if item.get("label") == "客户"),
        None,
    )
    attempts = []
    if step and step.get("path"):
        deadline = time.perf_counter() + max(float(timeout or 0), 0.0)
        while True:
            found = find_receipt_header_field_by_dynamic_path(
                jab,
                "客户",
                step.get("dynamic_index") or dynamic_index,
                scope_hwnd=scope_hwnd,
                require_showing=False,
                require_valid_bounds=False,
                path_template=(step.get("path_attempt") or {}).get("path_template"),
            )
            attempt = read_customer_name_from_found_field(
                jab,
                found,
                source="path-readback",
            )
            attempts.append(attempt)
            value = str(attempt.get("value") or "").strip()
            if is_valid_customer_name_candidate(value):
                return {
                    "ok": True,
                    "value": value,
                    "source": attempt.get("source"),
                    "attempts": attempts,
                }
            remaining = deadline - time.perf_counter()
            if remaining <= 0:
                break
            time.sleep(min(max(float(poll_interval or 0), 0.02), remaining))
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
        "reason": format_customer_readback_failure(attempts),
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


def format_customer_readback_failure(attempts):
    details = []
    for attempt in attempts or []:
        source = str(attempt.get("source") or "unknown")
        if attempt.get("snapshot"):
            snapshot = attempt.get("snapshot") or {}
            text = str(snapshot.get("text") or "").strip()
            name = str(snapshot.get("name") or "").strip()
            description = str(snapshot.get("description") or "").strip()
        else:
            text = str(attempt.get("text") or "").strip()
            name = str(attempt.get("name") or "").strip()
            description = str(attempt.get("description") or "").strip()
        reason = str(attempt.get("reason") or "").strip()
        fields = []
        if text:
            fields.append(f"text={text!r}")
        if name:
            fields.append(f"name={name!r}")
        if description:
            fields.append(f"description={description!r}")
        if reason and not fields:
            fields.append(f"reason={reason}")
        if fields:
            details.append(f"{source}: " + ", ".join(fields))
    if not details:
        return "客户名称未确认：客户字段未回显有效 NC 客户名称"
    return "客户名称未确认：客户字段未回显有效 NC 客户名称；读回：" + "；".join(
        details[-3:]
    )


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
    if not post_query:
        return {"*": "后验查询失败"}
    issues = {}
    for group in post_query.get("groups") or []:
        match = group.get("match") or {}
        for row, reason in (match.get("issues") or {}).items():
            issues[str(row)] = reason or "后验未匹配"
        if not group.get("ok"):
            reason = group.get("reason") or "后验查询失败"
            for row in group.get("target_rows") or []:
                issues.setdefault(str(row), reason)
    if not post_query.get("ok") and not issues:
        return {"*": post_query.get("reason") or "后验查询失败"}
    return issues


def post_query_skip_reason(rows, exit_code):
    if not rows:
        return "没有完成任何收款单录入行，未执行后验查询"
    if exit_code != 0:
        return "录入/保存阶段未全部成功，未执行后验查询"
    return "后验查询条件未满足"


def should_retry_row_by_cancel_reopen(row_report):
    if not row_report or row_report.get("ok"):
        return False
    if (row_report.get("circuit_breaker") or {}).get("triggered"):
        return False
    if row_report.get("save_attempted"):
        return False
    failed_step = str(row_report.get("failed_step") or "")
    if failed_step not in CIRCUIT_BREAKER_RETRY_STEPS:
        return False
    save_report = row_report.get("save") or {}
    if save_report and not save_report.get("skipped"):
        return False
    return True


def summarize_retry_attempt(row_report):
    return {
        "ok": bool((row_report or {}).get("ok")),
        "excel_row": (row_report or {}).get("excel_row"),
        "failed_step": (row_report or {}).get("failed_step"),
        "reason": (row_report or {}).get("reason"),
        "slow_steps": (row_report or {}).get("slow_steps") or [],
        "detail_exchange_rate_guard": (row_report or {}).get(
            "detail_exchange_rate_guard"
        ),
        "extra_text_verify_failures": (row_report or {}).get(
            "extra_text_verify_failures"
        ),
        "detail_pipeline_verify": (row_report or {}).get("detail_pipeline_verify"),
        "detail_pipeline_verify_after_repair": (row_report or {}).get(
            "detail_pipeline_verify_after_repair"
        ),
    }


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


def save_receipt_by_ctrl_s(
    jab,
    scope_hwnd=None,
    timeout=3.0,
    min_samples=3,
    interval=0.1,
    success_samples=1,
    min_observe_seconds=0.0,
):
    page = None
    if not scope_hwnd:
        return {
            "ok": False,
            "triggered": False,
            "reason": "Ctrl+S 保存前未取得收款单窗口句柄",
            "page": page,
        }
    target_hwnd = root_hwnd(scope_hwnd) or scope_hwnd
    maximize = jab.maximize_window_by_handle(target_hwnd)
    guard = foreground_matches_window({"hwnd": target_hwnd})
    if not guard.get("ok"):
        return {
            "ok": False,
            "triggered": False,
            "reason": guard.get("reason") or "当前前台窗口不是目标 NC 窗口",
            "maximize": maximize,
            "guard": guard,
            "page": page,
            "scope_hwnd": scope_hwnd,
            "target_hwnd": target_hwnd,
        }
    try:
        send_hotkey_ctrl_s()
    except Exception as exc:
        return {
            "ok": False,
            "triggered": False,
            "reason": f"Ctrl+S SendInput 触发失败：{type(exc).__name__}: {exc}",
            "guard": guard,
            "page": page,
            "scope_hwnd": scope_hwnd,
            "target_hwnd": target_hwnd,
        }
    wait_parent = wait_receipt_parent_new_ready_after_entry_exit(
        jab,
        timeout=timeout,
        interval=interval,
        success_samples=success_samples,
    )
    wait_parent.setdefault("oracle", {})["name"] = "receipt_parent_new_ready_after_save"
    if wait_parent.get("ok"):
        return {
            "ok": True,
            "triggered": True,
            "hotkey": {"ok": True, "mode": "send_input", "key": "Ctrl+S"},
            "precondition": {
                "page": page,
                "maximize": maximize,
                "foreground_guard": guard,
                "scope_hwnd": scope_hwnd,
                "target_hwnd": target_hwnd,
            },
            "seconds": wait_parent.get("seconds"),
            "samples": wait_parent.get("samples") or [],
            "min_samples": int(min_samples),
            "success_samples": int(success_samples),
            "timeout": float(timeout),
            "min_observe_seconds": float(min_observe_seconds),
            "oracle": wait_parent.get("oracle"),
            "entry_state": wait_parent.get("entry_state"),
            "parent_new_state": wait_parent.get("parent_new_state"),
        }
    return {
        "ok": False,
        "triggered": True,
        "hotkey": {"ok": True, "mode": "send_input", "key": "Ctrl+S"},
        "precondition": {
            "page": page,
            "maximize": maximize,
            "foreground_guard": guard,
            "scope_hwnd": scope_hwnd,
            "target_hwnd": target_hwnd,
        },
        "seconds": wait_parent.get("seconds"),
        "samples": wait_parent.get("samples") or [],
        "min_samples": int(min_samples),
        "success_samples": int(success_samples),
        "timeout": float(timeout),
        "min_observe_seconds": float(min_observe_seconds),
        "reason": "保存后未确认收款单父页【新增】已恢复，不能证明保存成功",
        "oracle": wait_parent.get("oracle"),
        "entry_state": wait_parent.get("entry_state"),
        "parent_new_state": wait_parent.get("parent_new_state"),
    }


def cancel_current_receipt_entry(
    config,
    timeout=3.0,
    interval=0.1,
    confirm_wait=0.8,
):
    jab = JABOperator(config)
    jab.hide_blank_awt_windows_enabled = False
    report = {
        "ok": False,
        "method": "ctrl-q-confirm-alt-y",
        "triggered": False,
        "confirmed": False,
    }
    try:
        jab.ensure_started()
        before_windows = collect_receipt_new_windows(jab)
        before_entry_state = detect_self_made_entry_state(before_windows)
        report["entry_state_before"] = before_entry_state
        if not before_entry_state.get("ok"):
            report["reason"] = "取消前未检测到保存/暂存/取消录入态按钮"
            return report
        target_hwnd = current_receipt_root_from_entry_state(before_entry_state)
        if not target_hwnd:
            report["reason"] = "取消前未取得当前收款单窗口句柄"
            return report
        report["target_hwnd"] = target_hwnd
        maximize = jab.maximize_window_by_handle(target_hwnd)
        guard = foreground_matches_window({"hwnd": target_hwnd})
        report["precondition"] = {
            "maximize": maximize,
            "foreground_guard": guard,
        }
        if not guard.get("ok"):
            report["reason"] = guard.get("reason") or "当前前台窗口不是目标 NC 窗口"
            return report
        dialogs_before = collect_visible_java_dialogs(jab)
        report["dialogs_before_count"] = len(dialogs_before)
        try:
            send_hotkey_ctrl_q()
        except Exception as exc:
            report["reason"] = f"Ctrl+Q SendInput 触发失败：{type(exc).__name__}: {exc}"
            return report
        report["triggered"] = True
        dialog_wait = wait_confirm_cancel_dialog(
            jab,
            dialogs_before,
            timeout=float(confirm_wait or 0.8),
            interval=0.08,
        )
        report["confirm_dialog"] = dialog_wait
        dialog = dialog_wait.get("dialog")
        if not dialog_wait.get("ok") or not dialog:
            report["reason"] = dialog_wait.get("reason") or "未检测到确认取消弹窗"
            return report
        focus = focus_window(dialog.get("hwnd"))
        report["confirm_focus"] = focus
        try:
            send_hotkey_alt_y()
        except Exception as exc:
            report["reason"] = f"Alt+Y SendInput 确认失败：{type(exc).__name__}: {exc}"
            return report
        report["confirmed"] = True
        wait_parent = wait_receipt_parent_new_ready_after_entry_exit(
            jab,
            timeout=timeout,
            interval=interval,
        )
        report["parent_ready_after_cancel"] = wait_parent
        report["ok"] = bool(wait_parent.get("ok"))
        if not report["ok"]:
            report["reason"] = wait_parent.get("reason") or "取消后未确认父页新增可用"
        return report
    finally:
        jab.close()


def current_receipt_root_from_entry_state(entry_state):
    hits = (entry_state or {}).get("hits") or []
    for hit in hits:
        window = hit.get("window") or {}
        hwnd = window.get("hwnd")
        if hwnd:
            return root_hwnd(hwnd) or hwnd
    return None


def wait_confirm_cancel_dialog(jab, before_dialogs, timeout=0.8, interval=0.08):
    started = time.perf_counter()
    before_keys = {dialog_key(item) for item in before_dialogs or []}
    attempts = []
    while True:
        dialogs = collect_visible_java_dialogs(jab)
        matching = [
            item
            for item in dialogs
            if is_confirm_cancel_dialog(item)
            and (
                dialog_key(item) not in before_keys
                or not before_keys
            )
        ]
        attempts.append(
            {
                "t": round(time.perf_counter() - started, 3),
                "dialog_count": len(dialogs),
                "matching_count": len(matching),
            }
        )
        if matching:
            return {
                "ok": True,
                "seconds": round(time.perf_counter() - started, 3),
                "attempts": attempts,
                "dialog": summarize_dialog_for_report(matching[0]),
            }
        if time.perf_counter() - started >= float(timeout or 0):
            return {
                "ok": False,
                "seconds": round(time.perf_counter() - started, 3),
                "attempts": attempts,
                "reason": "未发现标题为【确认取消】且包含【是(Y)/否(N)】的确认弹窗",
                "last_dialogs": [summarize_dialog_for_report(item) for item in dialogs],
            }
        time.sleep(float(interval or 0.08))


def dialog_key(dialog):
    return (
        (dialog or {}).get("hwnd"),
        (dialog or {}).get("title"),
        (dialog or {}).get("class_name"),
    )


def is_confirm_cancel_dialog(dialog):
    if (dialog or {}).get("class_name") != "SunAwtDialog":
        return False
    if (dialog or {}).get("title") != "确认取消":
        return False
    names = {button.get("name") for button in (dialog or {}).get("buttons") or []}
    return {"是(Y)", "否(N)"} <= names


def summarize_dialog_for_report(dialog):
    if not dialog:
        return None
    return {
        "hwnd": dialog.get("hwnd"),
        "title": dialog.get("title"),
        "class_name": dialog.get("class_name"),
        "pid": dialog.get("pid"),
        "visible": dialog.get("visible"),
        "root_hwnd": dialog.get("root_hwnd"),
        "buttons": [
            {
                "path": button.get("path"),
                "name": button.get("name"),
                "description": button.get("description"),
                "bounds": button.get("bounds"),
            }
            for button in (dialog.get("buttons") or [])
        ],
    }


def wait_receipt_parent_new_ready_after_entry_exit(
    jab,
    timeout=3.0,
    interval=0.1,
    success_samples=1,
):
    started = time.perf_counter()
    samples = []
    strong_success_streak = 0
    last_state = None
    last_parent_new_state = None
    while True:
        sample_started = time.perf_counter()
        windows = collect_receipt_new_windows(jab)
        state = detect_self_made_entry_state(windows)
        parent_new_state = detect_receipt_parent_new_ready(windows)
        last_state = state
        last_parent_new_state = parent_new_state
        strong_success = bool(parent_new_state.get("ok")) and not state.get("ok")
        strong_success_streak = strong_success_streak + 1 if strong_success else 0
        samples.append(
            {
                "sample_index": len(samples) + 1,
                "t": round(time.perf_counter() - started, 3),
                "collect_seconds": round(time.perf_counter() - sample_started, 3),
                "new_candidate_count": parent_new_state.get("candidate_count"),
                "new_usable_count": parent_new_state.get("usable_new_button_count"),
                "entry_ok": bool(state.get("ok")),
                "entry_partial_ok": bool(state.get("partial_ok")),
                "strong_success": strong_success,
                "strong_success_streak": strong_success_streak,
            }
        )
        if strong_success_streak >= int(success_samples or 1):
            return {
                "ok": True,
                "seconds": round(time.perf_counter() - started, 3),
                "samples": samples,
                "oracle": {
                    "name": "receipt_parent_new_ready_after_entry_exit",
                    "ok": True,
                    "evidence": "检测到收款单父页【新增】按钮可用，且保存/暂存/取消三按钮不再同时存在",
                    "parent_new_state": parent_new_state,
                    "self_made_entry_state": state,
                },
                "entry_state": state,
                "parent_new_state": parent_new_state,
            }
        if time.perf_counter() - started >= float(timeout or 0):
            return {
                "ok": False,
                "seconds": round(time.perf_counter() - started, 3),
                "samples": samples,
                "reason": "未确认收款单父页【新增】已恢复",
                "oracle": {
                    "name": "receipt_parent_new_ready_after_entry_exit",
                    "ok": False,
                    "evidence": "需要同时满足：前台收款单父页【新增】按钮可用，且保存/暂存/取消三按钮不再同时存在",
                    "parent_new_state": last_parent_new_state,
                    "self_made_entry_state": last_state,
                },
                "entry_state": last_state,
                "parent_new_state": last_parent_new_state,
            }
        time.sleep(float(interval or 0.1))


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
        }
    )
    attach_slow_step_summary(row_report, timings)
    return row_report


def attach_slow_step_summary(
    row_report,
    timings,
    threshold_seconds=SLOW_STEP_THRESHOLD_SECONDS,
):
    timing_items = list(getattr(timings, "items", []) or [])
    row_report["timings"] = timing_items
    row_report["slow_step_threshold_seconds"] = float(threshold_seconds)
    slow_steps = []

    def add_step(name, seconds, source, details=None):
        try:
            value = float(seconds)
        except (TypeError, ValueError):
            return
        if value < float(threshold_seconds):
            return
        item = {
            "name": name,
            "seconds": round(value, 3),
            "source": source,
        }
        if details:
            item["details"] = details
        slow_steps.append(item)

    for item in timing_items:
        add_step(item.get("name"), item.get("seconds"), "row")

    open_step = find_report_step(row_report, "open-self-made")
    parsed = (open_step or {}).get("parsed") or {}
    entry_context = parsed.get("entry_context_snapshot")
    if entry_context:
        row_report["open_self_made_entry_context"] = entry_context
    for item in parsed.get("timings") or []:
        add_step(
            f"open.self-made/{item.get('name')}",
            item.get("seconds"),
            "open.self-made",
        )

    slow_steps.sort(key=lambda item: item["seconds"], reverse=True)
    row_report["slow_steps"] = slow_steps[:30]


def find_report_step(row_report, name):
    for step in (row_report or {}).get("steps") or []:
        if step.get("name") == name:
            return step
    return None


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
    elif report.get("post_query_failed_rows"):
        ok_rows = [row.get("excel_row") for row in rows if row.get("ok")]
        lines.append(f"录入保存通过行：{ok_rows}")
        lines.append("失败阶段：post-query")
        for row, reason in (report.get("post_query_failed_rows") or {}).items():
            lines.append(f"后验未匹配行 {row}：{reason}")
    elif report.get("ok"):
        ok_rows = [row.get("excel_row") for row in rows if row.get("ok")]
        lines.append(f"通过行：{ok_rows}")
        post_query = report.get("post_query") or {}
        if post_query:
            matched = 0
            issues = 0
            for group in post_query.get("groups") or []:
                match = group.get("match") or {}
                matched += len(match.get("matched") or {})
                issues += len(match.get("issues") or {})
            lines.append(f"后验查询：已执行，匹配 {matched} 行，未匹配 {issues} 行")
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
