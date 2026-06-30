# 职责：编排收款单完整流程入口，消费 ReceiptPlanRow 跑开单/表头/明细/手续费/保存闸/后验查询
# 不做什么：不做录入前 NC 查重，不复用历史 T0 保存脚本，不把保存设为默认行为
# 允许依赖层：core 收款计划/配置/JAB、tools 下已正式化的收款开单/明细/查询组件
# 谁不应该 import：core 层模块不应 import 本入口；凭证批量模块不应 import

import argparse
from dataclasses import asdict
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
from core.run_state import RunStateRecorder  # noqa: E402
from core.utils import load_config  # noqa: E402
from core.receipt_detail_async_verifier import DetailPipelineVerifier  # noqa: E402
from core.receipt_detail_fields import validate_main_row_exchange_rate  # noqa: E402
from core.receipt_detail_row_cleanup import delete_extra_row_if_present  # noqa: E402
from core.receipt_detail_rows import StepTimer, run_fee_only  # noqa: E402
from core.receipt_detail_writer import (  # noqa: E402
    write_detail_line_by_screen,
    write_field_once,
)
from core.receipt_modal_guard import (  # noqa: E402
    recover_cancelable_modal_now,
)
from core.receipt_post_save_query import run_post_save_batch_query  # noqa: E402
from core.receipt_self_made_fill_trial import (  # noqa: E402
    fill_header,
    find_receipt_header_field_by_dynamic_path,
    find_receipt_header_field_by_live_semantic,
    get_receipt_header_path_template,
    infer_header_path_template_from_field,
    is_valid_customer_name_candidate,
    read_body_table,
    receipt_header_dynamic_prefix,
    run_receipt_new_probe,
    run_receipt_new_probe_with_jab,
    set_receipt_header_path_template,
    wait_header_account_description,
    guarded_paste_header_value,
    describe_backend_field_state,
)

from core.receipt_locator_cache import (  # noqa: E402
    build_header_scope_for_followup,
    cache_receipt_header_scope,
    resolve_body_table_by_dynamic_prefix,
)
from core.receipt_counterparty import (  # noqa: E402
    ensure_header_counterparty_customer,
)
from core.receipt_save_cancel import (  # noqa: E402
    CIRCUIT_BREAKER_RETRY_STEPS,
    cancel_current_receipt_entry,
    save_receipt_by_ctrl_s,
    should_retry_row_by_cancel_reopen,
    summarize_retry_attempt,
)
from core.receipt_report import (  # noqa: E402
    attach_slow_step_summary,
    build_batch_results,
    fail,
    post_query_failure_reasons,
    post_query_skip_reason,
    print_report,
    serializable,
    summarize_header_failure,
    user_excel_locked_message,
    write_last_report,
)
from core.receipt_row_stages import resolve_entry_header_scope  # noqa: E402


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
    try:
        workbook = ReceiptEntryWorkbook(config, excel_path=args.excel_path)
        plan_rows, issues, summary = workbook.build_local_plan(write_sheet=False)
    except Exception as exc:  # 计划装载早退也要收尾 run_state，避免永停 running 态
        reason = f"计划装载失败：{exc}"
        report.update(
            {
                "ok": False,
                "reason": reason,
                "error_category": "plan_load_failed",
                "total_seconds": round(time.perf_counter() - started, 3),
            }
        )
        recorder.event("plan-load-failed", reason=reason)
        recorder.finish("failed", error=reason)
        write_last_report(report)
        print_report(report, args)
        return 1
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


def build_header_unified_targets(header_steps, extra_text_report):
    targets = []
    seen = set()
    for step in header_steps or []:
        label = str(step.get("label") or "").strip()
        path = str(step.get("path") or "").strip()
        value = str(step.get("value") or "").strip()
        if label == "财务组织":
            continue
        if not label or not path or not step.get("ok"):
            continue
        key = ("header", label, path)
        if key in seen:
            continue
        seen.add(key)
        targets.append(
            {
                "kind": "customer" if label == "客户" else "header",
                "label": label,
                "value": value,
                "path": path,
                "accepted_text": step.get("accepted_text"),
                "source": step.get("source") or step.get("method") or "header-fill",
            }
        )
    return targets


def read_header_target_by_exact_path(jab, target, scope_hwnd):
    path = str((target or {}).get("path") or "").strip()
    label = str((target or {}).get("label") or "").strip()
    if not path:
        return {
            "ok": False,
            "label": label,
            "path": path,
            "kind": (target or {}).get("kind"),
            "reason": "缺少表头字段 path",
        }
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="text",
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "label": label,
            "path": path,
            "kind": (target or {}).get("kind"),
            "reason": "表头字段 path 不可读",
        }
    try:
        info = jab.get_context_info(vm_id, context)
        text = jab.get_text_context_value(vm_id, context)
        description = info.description.strip() if info else ""
        name = info.name.strip() if info else ""
        kind = (target or {}).get("kind")
        snapshot = describe_backend_field_state(
            info,
            text,
            value=(target or {}).get("value"),
            accepted_text=(target or {}).get("accepted_text"),
        )
        if kind == "customer":
            actual_value = first_valid_text(description, text, name)
            ok = bool(actual_value)
        elif kind == "extra_text":
            actual_value = description or str(text or "").strip() or name
            ok = bool(snapshot.get("written"))
        elif label == "币种":
            actual_value = description or str(text or "").strip() or name
            ok = header_currency_matches(
                (target or {}).get("value"),
                (target or {}).get("accepted_text"),
                actual_value,
                text,
                name,
                description,
            )
        else:
            actual_value = description or str(text or "").strip() or name
            ok = bool(snapshot.get("accepted") or snapshot.get("written"))
        return {
            "ok": ok,
            "label": label,
            "path": path,
            "kind": kind,
            "value": (target or {}).get("value"),
            "actual_value": actual_value,
            "text": text,
            "description": description,
            "name": name,
            "window": window_info,
            "snapshot": snapshot,
            "reason": None if ok else f"{label} 未读回目标值",
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def normalize_header_currency_value(*values):
    for value in values:
        text = str(value or "").strip().upper()
        if not text:
            continue
        compact = (
            text.replace("/", "")
            .replace("\\", "")
            .replace(" ", "")
            .replace("-", "")
            .replace("_", "")
        )
        if "USD" in compact or "美元" in compact:
            return "USD"
        if "CNY" in compact or "RMB" in compact or "人民币" in compact:
            return "CNY"
    return ""


def header_currency_matches(expected_value, accepted_text, *actual_values):
    expected = normalize_header_currency_value(expected_value, accepted_text)
    actual = normalize_header_currency_value(*actual_values)
    return bool(expected and actual and expected == actual)


def rewrite_header_target_by_exact_path(
    jab,
    target,
    scope_hwnd,
    recover_after_failure=None,
):
    path = str((target or {}).get("path") or "").strip()
    label = str((target or {}).get("label") or "").strip()
    if not path:
        return {
            "ok": False,
            "label": label,
            "path": path,
            "kind": (target or {}).get("kind"),
            "reason": "缺少表头字段 path，不能补写",
        }
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="text",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context and recover_after_failure is not None:
        recovery = recover_after_failure()
        if recovery.get("attempted") and recovery.get("ok"):
            context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
                path,
                class_name="SunAwtCanvas",
                scope_hwnd=scope_hwnd,
                role="text",
                require_showing=True,
                require_valid_bounds=False,
            )
        else:
            recovery = None
    else:
        recovery = None
    if not context:
        return {
            "ok": False,
            "label": label,
            "path": path,
            "kind": (target or {}).get("kind"),
            "modal_recovery": recovery,
            "reason": "表头字段 path 不可写，不能补写",
        }
    try:
        paste = guarded_paste_header_value(
            jab,
            vm_id,
            context,
            window_info or {},
            (target or {}).get("value"),
        )
        return {
            "ok": bool(paste.get("ok")),
            "label": label,
            "path": path,
            "kind": (target or {}).get("kind"),
            "value": (target or {}).get("value"),
            "guarded_paste": paste,
            "modal_recovery": recovery,
            "reason": None if paste.get("ok") else paste.get("reason"),
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def verify_and_repair_header_targets(
    jab,
    header_steps,
    extra_text_report,
    dynamic_index,
    scope_hwnd,
    recover_after_failure=None,
):
    started_at = time.perf_counter()
    targets = build_header_unified_targets(header_steps, extra_text_report)
    if not targets:
        return {
            "ok": True,
            "skipped": True,
            "reason": "没有可用 path 的表头统一校验目标",
            "dynamic_index": dynamic_index,
            "scope_hwnd": scope_hwnd,
            "targets": [],
            "reads": [],
            "missing": [],
            "repairs": [],
            "rereads": [],
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    reads = [read_header_target_by_exact_path(jab, target, scope_hwnd) for target in targets]
    missing = [item for item in reads if not item.get("ok")]
    repairs = []
    rereads = []
    if missing:
        targets_by_path = {target.get("path"): target for target in targets}
        for item in missing:
            target = targets_by_path.get(item.get("path"))
            if not target:
                continue
            repair = rewrite_header_target_by_exact_path(
                jab,
                target,
                scope_hwnd,
                recover_after_failure=recover_after_failure,
            )
            repairs.append(repair)
        rereads = [
            read_header_target_by_exact_path(jab, targets_by_path[item.get("path")], scope_hwnd)
            for item in missing
            if item.get("path") in targets_by_path
        ]
    failed_after_repair = [item for item in rereads if not item.get("ok")]
    ok = not missing or (bool(repairs) and not failed_after_repair)
    reason = ""
    if failed_after_repair:
        reason = "表头统一校验补写后仍有缺失：" + "；".join(
            f"{item.get('label')}: {item.get('reason')}" for item in failed_after_repair
        )
    elif missing and not repairs:
        reason = "表头统一校验发现缺失，但没有可补写字段"
    return {
        "ok": bool(ok),
        "dynamic_index": dynamic_index,
        "scope_hwnd": scope_hwnd,
        "targets": targets,
        "reads": reads,
        "missing": missing,
        "repairs": repairs,
        "rereads": rereads,
        "reason": reason,
        "seconds": round(time.perf_counter() - started_at, 3),
    }


def customer_name_from_header_unified_check(report):
    candidates = []
    for section in ("rereads", "reads"):
        for item in (report or {}).get(section) or []:
            if item.get("kind") == "customer" and item.get("ok"):
                candidates.append(item.get("actual_value"))
    for value in candidates:
        text = str(value or "").strip()
        if is_valid_customer_name_candidate(text):
            return text
    return ""


class StageAbort(Exception):
    """阶段中止信号:任一阶段判失败时抛出,由 RowRun.execute() 统一转 fail()。"""

    def __init__(self, step, reason):
        super().__init__(reason)
        self.step = step
        self.reason = reason


class RowRun:
    """单行收款单全流程的一次性执行器:开单→表头→明细主行→手续费→保存。

    原 run_one_row 上帝函数(radon F63、647行)的内聚状态归位为对象字段、五阶段
    归位为方法;任一阶段判失败 raise StageAbort,由 execute() 统一收尾(关 verifier、
    关 jab)。行为与原函数严格等价:row_report 键、fail 的 step 名、event 名均不变。
    """

    def __init__(
        self,
        config,
        row,
        save_enabled=False,
        recorder=None,
        pause_after_header_field=None,
        diagnose_header_after_pause=False,
        diagnose_detail_repair=False,
        header_scope_cache=None,
    ):
        self.config = config
        self.row = row
        self.save_enabled = save_enabled
        self.recorder = recorder
        self.pause_after_header_field = pause_after_header_field
        self.diagnose_header_after_pause = diagnose_header_after_pause
        self.diagnose_detail_repair = diagnose_detail_repair
        self.header_scope_cache = header_scope_cache
        self.current_stage = ""
        self.timings = StepTimer()
        self.flow_started_at = time.perf_counter()
        self.row_report = {
            "excel_row": row.row,
            "plan_row": serializable(asdict(row)),
            "steps": [],
            "save_enabled": bool(save_enabled),
        }
        self.business = business_from_plan_row(row)
        self.jab = JABOperator(config)
        self.jab_lock = threading.RLock()
        self.pipeline_verifier = None
        self.modal_events = []
        # 表头阶段填、后续阶段读的跨阶段派生状态
        self.open_step = None
        self.entry_scope_hwnd = None
        self.entry_dynamic_index = None
        self.entry_anchor_path = None
        self.located = None
        self.extra_text_report = None
        self.header_pause_reports = []
        self.header_steps_so_far_labels = []
        # 明细阶段建、手续费/校验阶段读的 pipeline 任务台账
        self.pipeline_field_task_ids = []
        self.pipeline_field_tasks = {}
        self.pipeline_text_task_ids = []
        self.pipeline_snapshot_task_ids = []
        self.pipeline_row_count_task_id = None

    # ---- 共享小工具(原闭包) ----
    def stage(self, stage, **fields):
        self.current_stage = stage
        if self.recorder is not None:
            self.recorder.set_stage(stage, **fields)

    def event(self, name, **fields):
        if self.recorder is not None:
            self.recorder.event(name, **fields)

    def recover_modal(self):
        result = recover_cancelable_modal_now(
            self.jab,
            stage=self.current_stage or "",
        )
        if result.get("attempted"):
            self.modal_events.append(result)
        return result

    def _after_header_field(self, label, _value, step):
        if label and label not in self.header_steps_so_far_labels:
            self.header_steps_so_far_labels.append(label)
        if self.pause_after_header_field != label:
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
        if self.diagnose_header_after_pause:
            report["header_readback"] = diagnose_written_header_fields(
                self.jab,
                list(self.header_steps_so_far_labels),
                step.get("dynamic_index"),
                step.get("dynamic_prefix"),
                self.entry_scope_hwnd,
            )
            report["ok"] = all(
                item.get("present") for item in report["header_readback"]
            )
            if not report["ok"]:
                report["reason"] = "暂停恢复后检测到已写表头字段为空或不可读"
        self.header_pause_reports.append(report)
        return report

    def _submit_detail_verify(self, row_index, field, business_values, _step):
        task_id = self.pipeline_verifier.submit_field(
            row_index,
            field,
            business_values,
        )
        self.pipeline_field_task_ids.append(task_id)
        self.pipeline_field_tasks[task_id] = {
            "row_index": int(row_index),
            "field": dict(field),
            "business": business_values,
        }
        return task_id

    # ---- 五个阶段(任一失败 raise StageAbort) ----
    def open(self):
        self.timings.measure("jab.ensure-started", self.jab.ensure_started)
        self.stage("开单", excel_row=self.row.row)
        open_step = self.timings.measure(
            "open.self-made",
            run_with_jab_lock,
            self.jab_lock,
            open_self_made_entry,
            self.config,
            self.jab,
        )
        self.open_step = open_step
        self.row_report["steps"].append({"name": "open-self-made", **open_step})
        if not open_step.get("ok"):
            self.event(
                "open-failed", excel_row=self.row.row, error=open_step.get("reason")
            )
            raise StageAbort("open-self-made", open_step.get("reason"))
        self.row_report["modal_recovery"] = {"events": self.modal_events}

    def header(self):
        scope = resolve_entry_header_scope(
            self.jab,
            self.jab_lock,
            run_with_jab_lock,
            self.open_step,
            self.header_scope_cache,
            self.timings,
            self.row_report,
        )
        if not scope["ok"]:
            self.event(
                "header-anchor-failed", excel_row=self.row.row, error=scope["reason"]
            )
            raise StageAbort("header-anchor", scope["reason"])
        self.entry_scope_hwnd = scope["scope_hwnd"]
        self.entry_dynamic_index = scope["dynamic_index"]
        self.entry_anchor_path = scope["anchor_path"]
        self.stage("表头", excel_row=self.row.row)
        if self.pause_after_header_field:
            header_steps = self.timings.measure(
                "header.fill",
                run_with_jab_lock,
                self.jab_lock,
                fill_header,
                self.jab,
                self.business,
                scope_hwnd=self.entry_scope_hwnd,
                dynamic_index=self.entry_dynamic_index,
                anchor_path=self.entry_anchor_path,
                trust_provided_scope=True,
                recover_after_failure=self.recover_modal,
                after_field=self._after_header_field,
            )
        else:
            header_steps = self.timings.measure(
                "header.fill",
                run_with_jab_lock,
                self.jab_lock,
                fill_header,
                self.jab,
                self.business,
                after_field=self._after_header_field,
                scope_hwnd=self.entry_scope_hwnd,
                dynamic_index=self.entry_dynamic_index,
                anchor_path=self.entry_anchor_path,
                trust_provided_scope=True,
                recover_after_failure=self.recover_modal,
            )
        if self.header_pause_reports:
            self.row_report["header_pause_diagnostics"] = self.header_pause_reports
        self.row_report["header_steps"] = header_steps
        cached_after_header = (
            getattr(self.jab, "_receipt_header_scope_cache", None) or {}
        )
        if cached_after_header.get("ok"):
            cache_receipt_header_scope(
                self.jab, self.header_scope_cache, cached_after_header
            )
        if any(not step.get("ok") for step in header_steps):
            header_error = summarize_header_failure(header_steps)
            self.event(
                "header-fill-failed", excel_row=self.row.row, error=header_error
            )
            raise StageAbort("header-fill", header_error)
        self.located = self.timings.measure(
            "body.locate",
            run_with_jab_lock,
            self.jab_lock,
            resolve_body_table_by_dynamic_prefix,
            self.jab,
            self.entry_dynamic_index,
            self.entry_scope_hwnd,
        )
        located = self.located
        self.row_report["body_locate"] = {
            "source": located.get("source"),
            "cache_hit": located.get("cache_hit"),
            "fallback_used": located.get("fallback_used"),
            "status": located.get("status"),
            "seconds": located.get("seconds"),
            "started_offset_seconds": located.get("started_offset_seconds"),
        }
        self.row_report["table_candidates"] = located.get("candidates", [])[:5]
        if not located.get("best"):
            self.event(
                "locate-table-failed", excel_row=self.row.row, error="未定位到明细表"
            )
            raise StageAbort("locate-body-table", "未定位到明细表")
        counterparty_report = self.timings.measure(
            "header.counterparty-type",
            run_with_jab_lock,
            self.jab_lock,
            ensure_header_counterparty_customer,
            self.jab,
            self.entry_dynamic_index,
            self.entry_scope_hwnd,
            located=located,
            recover_after_failure=self.recover_modal,
        )
        self.row_report["header_counterparty"] = counterparty_report
        if not counterparty_report.get("ok"):
            self.event(
                "header-counterparty-failed",
                excel_row=self.row.row,
                error=counterparty_report.get("reason"),
            )
            raise StageAbort(
                "header-counterparty-type",
                counterparty_report.get("reason") or "往来对象未确认",
            )
        extra_text_report = self.timings.measure(
            "header.extra-text-fields",
            run_with_jab_lock,
            self.jab_lock,
            write_extra_text_fields,
            self.jab,
            self.row.extra_text_fields,
            self.entry_dynamic_index,
            self.entry_scope_hwnd,
            recover_after_failure=self.recover_modal,
            verify_after_write=False,
        )
        self.extra_text_report = extra_text_report
        self.row_report["extra_text_fields"] = extra_text_report
        if not extra_text_report.get("ok"):
            self.event(
                "header-extra-text-failed",
                excel_row=self.row.row,
                error=extra_text_report.get("reason"),
            )
            raise StageAbort(
                "header-extra-text-fields",
                extra_text_report.get("reason") or "扩展文本字段写入失败",
            )
        header_unified_check = self.timings.measure(
            "header.unified-check",
            run_with_jab_lock,
            self.jab_lock,
            verify_and_repair_header_targets,
            self.jab,
            header_steps,
            extra_text_report,
            self.entry_dynamic_index,
            self.entry_scope_hwnd,
            recover_after_failure=self.recover_modal,
        )
        self.row_report["header_unified_check"] = header_unified_check
        if not header_unified_check.get("ok"):
            reason = header_unified_check.get("reason") or "表头统一校验失败"
            self.event(
                "header-unified-check-failed",
                excel_row=self.row.row,
                error=reason,
            )
            raise StageAbort("header-unified-check", reason)
        unified_customer_name = customer_name_from_header_unified_check(
            header_unified_check
        )
        if unified_customer_name:
            customer_name = {
                "ok": True,
                "value": unified_customer_name,
                "source": "header-unified-check",
                "skipped_legacy_readback": True,
            }
        else:
            customer_name = self.timings.measure(
                "header.customer-name-readback",
                run_with_jab_lock,
                self.jab_lock,
                read_customer_name_after_header,
                self.jab,
                header_steps,
                self.entry_dynamic_index,
                self.entry_scope_hwnd,
            )
        self.row_report["customer_name_readback"] = customer_name
        self.row_report["nc_customer_name"] = str(
            customer_name.get("value") or ""
        ).strip()
        if not self.row_report["nc_customer_name"]:
            reason = customer_name.get("reason") or "客户名称未确认"
            self.event(
                "header-customer-readback-failed",
                excel_row=self.row.row,
                error=reason,
            )
            raise StageAbort("header-customer-name", reason)

    def detail_main(self):
        located = self.located
        self.stage("明细主行", excel_row=self.row.row)
        self.row_report["before_table"] = {
            "ok": True,
            "skipped": True,
            "reason": "明细表 path 已定位；后台 pipeline verifier 负责预热 path 和并发读回",
        }
        self.pipeline_verifier = DetailPipelineVerifier(
            self.config,
            located,
            flow_started_at=self.flow_started_at,
            jab=self.jab,
            jab_lock=self.jab_lock,
        )
        self.pipeline_verifier.start()
        for field_report in self.extra_text_report.get("fields") or []:
            if not field_report.get("ok") or not field_report.get("value"):
                continue
            if not field_report.get("path"):
                raise StageAbort(
                    "header-extra-text-fields",
                    f"扩展文本字段 {field_report.get('label')!r} 缺少后台验证 path",
                )
            self.pipeline_text_task_ids.append(
                self.pipeline_verifier.submit_path_text(
                    field_report.get("label"),
                    field_report.get("path"),
                    field_report.get("value"),
                    scope_hwnd=self.entry_scope_hwnd,
                )
            )
        self.row_report["extra_text_verify_tasks"] = self.pipeline_text_task_ids
        detail_steps = self.timings.measure(
            "detail.main-line",
            run_with_jab_lock,
            self.jab_lock,
            write_detail_line_by_screen,
            self.jab,
            self.business,
            located,
            after_field=self._submit_detail_verify,
            recover_after_failure=self.recover_modal,
        )
        self.row_report["detail_steps"] = detail_steps
        self.pipeline_snapshot_task_ids.append(
            self.pipeline_verifier.submit_snapshot(
                "after-main-line",
                max_rows=3,
                min_matches=len(detail_steps),
            )
        )
        if not all(step.get("ok") for step in detail_steps):
            self.event(
                "detail-main-failed", excel_row=self.row.row, error="明细主行写入失败"
            )
            raise StageAbort("detail-main-line", "明细主行写入失败")

    def fee_and_verify(self):
        located = self.located
        if self.row.fee > 0:
            self.stage("手续费", excel_row=self.row.row)
            self.row_report["extra_row_delete"] = {
                "ok": True,
                "skipped": True,
                "reason": "手续费非 0，保留主行后自动带出的第 2 行给手续费覆盖",
            }
            add_row, fee_steps, clear_account, delete_extra = self.timings.measure(
                "detail.fee-line",
                run_with_jab_lock,
                self.jab_lock,
                run_fee_only,
                self.jab,
                located,
                str(self.row.fee),
                after_field=self._submit_detail_verify,
                recover_after_failure=self.recover_modal,
            )
            self.row_report["fee_row_add"] = add_row
            self.row_report["fee_steps"] = fee_steps
            self.pipeline_snapshot_task_ids.append(
                self.pipeline_verifier.submit_snapshot(
                    "after-fee-line",
                    max_rows=4,
                )
            )
            self.row_report["fee_account_clear"] = clear_account
            self.row_report["fee_extra_row_delete"] = delete_extra
            if delete_extra.get("ok"):
                self.pipeline_row_count_task_id = (
                    self.pipeline_verifier.submit_row_count(2)
                )
            if (
                not add_row.get("ok")
                or not all(step.get("ok") for step in fee_steps)
                or not clear_account.get("ok")
                or not delete_extra.get("ok")
            ):
                self.event(
                    "fee-line-failed", excel_row=self.row.row, error="手续费行处理失败"
                )
                raise StageAbort("detail-fee-line", "手续费行处理失败")
        else:
            self.row_report["extra_row_delete"] = self.timings.measure(
                "detail.delete-extra-after-main",
                run_with_jab_lock,
                self.jab_lock,
                delete_extra_row_if_present,
                self.jab,
                located,
                1,
                scope_hwnd=self.entry_scope_hwnd,
                defer_wait=True,
            )
            if not self.row_report["extra_row_delete"].get("ok"):
                self.event(
                    "delete-extra-failed",
                    excel_row=self.row.row,
                    error="主行后多余行删除失败",
                )
                raise StageAbort(
                    "detail-delete-extra-after-main", "主行后多余行删除失败"
                )
            self.row_report["fee_skipped"] = {
                "ok": True,
                "reason": "手续费为 0，跳过手续费行",
            }
            self.pipeline_row_count_task_id = self.pipeline_verifier.submit_row_count(1)
        expected_detail_rows = 2 if self.row.fee > 0 else 1
        pipeline_wait_ids = []
        if self.pipeline_field_task_ids:
            pipeline_wait_ids.append(self.pipeline_field_task_ids[-1])
        if self.pipeline_row_count_task_id:
            pipeline_wait_ids.append(self.pipeline_row_count_task_id)
        pipeline_wait_ids.extend(self.pipeline_text_task_ids)
        pipeline_wait_started = time.perf_counter()
        self.row_report["detail_pipeline_verify"] = self.pipeline_verifier.wait(
            pipeline_wait_ids,
            timeout=2.0,
        )
        self.row_report["detail_pipeline_state"] = verifier_snapshot(
            self.pipeline_verifier
        )
        self.timings.add(
            "detail.pipeline-final-wait",
            time.perf_counter() - pipeline_wait_started,
        )
        extra_text_failures = []
        if self.pipeline_text_task_ids:
            pipeline_results = (
                self.row_report["detail_pipeline_verify"].get("results") or {}
            )
            for task_id in self.pipeline_text_task_ids:
                result = pipeline_results.get(task_id)
                if result is None or not result.get("ok"):
                    extra_text_failures.append(
                        {
                            "task_id": task_id,
                            "result": result,
                        }
                    )
        if extra_text_failures:
            self.row_report["extra_text_verify_failures"] = extra_text_failures
            self.event(
                "extra-text-verify-failed",
                excel_row=self.row.row,
                error="扩展文本字段后台验证未通过",
            )
            raise StageAbort(
                "detail-extra-text-verify", "扩展文本字段后台验证未通过"
            )
        if self.diagnose_detail_repair:
            self.row_report["detail_pipeline_verify_before_repair_drill"] = dict(
                self.row_report["detail_pipeline_verify"]
            )
            self.row_report["detail_pipeline_verify"] = force_one_detail_field_pending(
                self.row_report["detail_pipeline_verify"],
                self.pipeline_field_task_ids,
            )
        self.row_report["detail_pipeline_snapshots"] = self.pipeline_snapshot_task_ids
        detail_pipeline_ok = bool(self.row_report["detail_pipeline_verify"].get("ok"))
        if not detail_pipeline_ok:
            repair_report = self.timings.measure(
                "detail.pipeline-repair",
                repair_detail_pipeline_failures,
                self.jab,
                self.jab_lock,
                located,
                self.pipeline_verifier,
                self.row_report["detail_pipeline_verify"],
                self.pipeline_field_tasks,
                self.pipeline_row_count_task_id,
                expected_detail_rows,
                self.entry_scope_hwnd,
                self.recover_modal,
            )
            self.row_report["detail_pipeline_repair"] = repair_report
            if repair_report.get("snapshot_task_id"):
                self.pipeline_snapshot_task_ids.append(
                    repair_report["snapshot_task_id"]
                )
            repair_wait_ids = repair_report.get("wait_ids") or []
            if repair_wait_ids:
                repair_wait_started = time.perf_counter()
                self.row_report["detail_pipeline_verify_after_repair"] = (
                    self.pipeline_verifier.wait(
                        repair_wait_ids,
                        timeout=2.0,
                    )
                )
                self.row_report["detail_pipeline_state_after_repair"] = (
                    verifier_snapshot(self.pipeline_verifier)
                )
                self.timings.add(
                    "detail.pipeline-repair-wait",
                    time.perf_counter() - repair_wait_started,
                )
                detail_pipeline_ok = bool(
                    self.row_report["detail_pipeline_verify_after_repair"].get("ok")
                )
        if not detail_pipeline_ok:
            self.row_report["after_table"] = self.timings.measure(
                "body.read-after-fallback",
                run_with_jab_lock,
                self.jab_lock,
                read_body_table,
                self.jab,
                "after_detail_fill",
                self.entry_scope_hwnd,
            )
            self.event(
                "pipeline-verify-failed",
                excel_row=self.row.row,
                error="后台明细验证未通过，已执行整表读 fallback",
            )
            raise StageAbort(
                "detail-pipeline-verify",
                "后台明细验证未通过，已执行整表读 fallback",
            )
        self.row_report["after_table"] = {
            "ok": True,
            "skipped": True,
            "reason": "后台 pipeline verifier 已覆盖最后字段与最终行数，跳过同步整表读",
        }
        guard_source = self.row_report.get(
            "detail_pipeline_verify_after_repair"
        ) or self.row_report.get("detail_pipeline_verify")
        main_row_cells = latest_snapshot_row_cells(guard_source, row_index=0)
        exchange_rate_check = validate_main_row_exchange_rate(
            main_row_cells,
            self.row.currency,
            self.row.raw_amount,
            row_index=0,
        )
        self.row_report["detail_exchange_rate_check"] = exchange_rate_check
        self.row_report["detail_exchange_rate_guard"] = exchange_rate_check
        if not exchange_rate_check.get("ok"):
            self.event(
                "exchange-rate-check-failed",
                excel_row=self.row.row,
                error=exchange_rate_check.get("reason"),
            )
            raise StageAbort(
                "detail-exchange-rate-check",
                exchange_rate_check.get("reason") or "汇率保存前校验未通过",
            )
        account_check = self.timings.measure(
            "header.account-readback-after-detail",
            run_with_jab_lock,
            self.jab_lock,
            wait_header_account_description,
            self.jab,
            0.0,
            scope=build_header_scope_for_followup(
                self.entry_scope_hwnd,
                self.entry_dynamic_index,
            ),
        )
        self.row_report["header_account"] = account_check
        if not account_check.get("accepted"):
            self.row_report["header_account_readback_warning"] = {
                "ok": False,
                "reason": "表头收款银行账户未从 JAB 后端读回；明细账号已由后台 pipeline 校验，继续保存/后验查询闭包",
                "account_check": account_check,
            }
            self.event(
                "account-readback-warning",
                excel_row=self.row.row,
                warning="表头收款银行账户未从 JAB 后端读回，继续执行",
            )

    def save(self):
        if self.save_enabled:
            self.stage("保存", excel_row=self.row.row)
            self.row_report["save_attempted"] = True
            save_result = self.timings.measure(
                "save.ctrl-s",
                run_with_jab_lock,
                self.jab_lock,
                save_receipt_by_ctrl_s,
                self.jab,
                self.entry_scope_hwnd,
            )
            if not save_result.get("ok"):
                recovery = self.timings.measure(
                    "save.modal-recovery-after-failure",
                    self.recover_modal,
                )
                if recovery.get("attempted") and recovery.get("ok"):
                    save_result = self.timings.measure(
                        "save.ctrl-s-retry-after-modal",
                        run_with_jab_lock,
                        self.jab_lock,
                        save_receipt_by_ctrl_s,
                        self.jab,
                        self.entry_scope_hwnd,
                    )
                    save_result["retried_after_modal_recovery"] = True
                    save_result["modal_recovery"] = recovery
                else:
                    save_result["modal_recovery"] = recovery
            self.row_report["save"] = save_result
            if not save_result.get("ok"):
                self.event(
                    "save-failed",
                    excel_row=self.row.row,
                    error=save_result.get("reason"),
                )
                raise StageAbort("save", save_result.get("reason"))
        else:
            self.row_report["save"] = {
                "ok": True,
                "skipped": True,
                "reason": "no-save 模式：已停在保存前，未触发 Ctrl+S",
            }

    def execute(self):
        try:
            self.open()
            self.header()
            self.detail_main()
            self.fee_and_verify()
            self.save()
            self.row_report["ok"] = True
            attach_slow_step_summary(self.row_report, self.timings)
            return self.row_report
        except StageAbort as abort:
            return fail(self.row_report, abort.step, self.timings, abort.reason)
        except Exception as exc:
            self.row_report["exception"] = {
                "type": type(exc).__name__,
                "message": str(exc),
                "stage": self.current_stage or "",
            }
            self.event(
                "row-exception",
                excel_row=self.row.row,
                stage=self.current_stage or "",
                error=f"{type(exc).__name__}: {exc}",
            )
            return fail(
                self.row_report,
                "exception",
                self.timings,
                f"{type(exc).__name__}: {exc}",
            )
        finally:
            if self.pipeline_verifier is not None:
                self.pipeline_verifier.close(timeout=0.2)
            self.row_report["modal_recovery"] = {"events": self.modal_events}
            self.jab.close()


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
    return RowRun(
        config,
        row,
        save_enabled=save_enabled,
        recorder=recorder,
        pause_after_header_field=pause_after_header_field,
        diagnose_header_after_pause=diagnose_header_after_pause,
        diagnose_detail_repair=diagnose_detail_repair,
        header_scope_cache=header_scope_cache,
    ).execute()


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
                field,
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
    verify_after_write=True,
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
            verify_after_write=verify_after_write,
        )
        results.append(result)
        if not result.get("ok"):
            return {
                "ok": False,
                "fields": results,
                "reason": f"未能写入 NC 文本字段 {label!r}",
            }
    return {"ok": True, "fields": results}


def write_extra_text_field_by_dynamic_path(
    jab,
    label,
    value,
    dynamic_index,
    scope_hwnd=None,
    recover_after_failure=None,
    verify_after_write=True,
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
            info_before = jab.get_context_info(vm_id, context) if verify_after_write else None
            before = jab.get_text_context_value(vm_id, context) if verify_after_write else ""
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
            if verify_after_write:
                info_after = jab.get_context_info(vm_id, context)
                after = jab.get_text_context_value(vm_id, context)
                backend_state = describe_backend_field_state(info_after, after, value=value)
                accepted = bool(
                    backend_state.get("accepted") or backend_state.get("written")
                )
                ok = bool(paste.get("ok") and accepted)
            else:
                info_after = None
                after = ""
                backend_state = {}
                accepted = None
                ok = bool(paste.get("ok"))
            attempt = {
                "ok": ok,
                "input_ok": bool(paste.get("ok")),
                "verify_after_write": bool(verify_after_write),
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
            if ok or not paste.get("ok") or not verify_after_write:
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


if __name__ == "__main__":
    raise SystemExit(main())
