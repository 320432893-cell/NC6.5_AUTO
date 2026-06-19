# 职责：提供收款单明细主行/手续费行正式测试 CLI 入口
# 不做什么：不打开收款单、不写表头、不保存/暂存、不读取 Excel 批量计划
# 允许依赖层：core 配置/JAB、tools.receipt_detail_*、tools.receipt_self_made_flow 读表兼容函数
# 谁不应该 import：Sheet 写入、收款匹配、凭证批量模块不应 import

import argparse
import json
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_config import ReceiptEntryConfig  # noqa: E402
from core.run_state import RunStateRecorder  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready  # noqa: E402
from tools.receipt_keyboard_utils import (  # noqa: E402
    STOP_HOTKEY,
    is_stop_hotkey_pressed,
)
from tools.receipt_body_table_locator import locate_receipt_body_table_cached  # noqa: E402
from tools.receipt_detail_report import print_header, print_summary  # noqa: E402
from tools.receipt_detail_row_cleanup import (  # noqa: E402
    cleanup_rows_after_first,
    delete_extra_row_if_present,
)
from tools.receipt_detail_rows import StepTimer, run_fee_only  # noqa: E402
from tools.receipt_detail_writer import write_detail_line_by_screen  # noqa: E402
from tools.receipt_self_made_flow import (  # noqa: E402
    read_body_table,
    wait_header_account_description,
)

DEFAULT_TEST_BANK_LABEL = "招行"
DEFAULT_TEST_CURRENCY = "美元"
START_DELAY_SECONDS = 2


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="收款单明细主行/手续费行正式测试入口；不保存、不暂存。"
    )
    parser.add_argument(
        "--fee-only",
        action="store_true",
        help="只测试手续费：Ctrl+I 增行后写新增行，不写主行。",
    )
    parser.add_argument("--fee-amount", default="10", help="手续费测试金额，默认 10。")
    parser.add_argument(
        "--bank-label",
        default=DEFAULT_TEST_BANK_LABEL,
        help="测试银行标签，默认 招行；实际账号从 config.json 读取。",
    )
    parser.add_argument(
        "--cleanup-extra-rows-only",
        action="store_true",
        help="只删除明细第 1 行以外的多余行，不写入、不保存。",
    )
    parser.add_argument(
        "--config",
        default=str(ROOT / "config.json"),
        help="配置文件路径，默认项目根 config.json。",
    )
    parser.add_argument(
        "--start-delay",
        type=float,
        default=START_DELAY_SECONDS,
        help="启动后等待切到 NC 窗口的秒数，默认 2。",
    )
    parser.add_argument(
        "--no-wait",
        action="store_true",
        help="跳过结尾等待回车（GUI 调用时使用），默认不跳过。",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="在 stdout 最后一行输出结构化结果信封（GUI 解析用）。",
    )
    return parser.parse_args(argv)


def get_test_account(config, bank_label):
    receipt_config = ReceiptEntryConfig(config)
    account = receipt_config.account_for_bank(bank_label)
    if account:
        return account
    raise RuntimeError(f"config.json 中找不到银行账户映射：{bank_label}")


def detail_bank_account_no(account):
    return account.account_no


def build_business(account):
    return {
        "currency": DEFAULT_TEST_CURRENCY,
        "bank_account": detail_bank_account_no(account),
        "amount": "1090",
        "settlement": "网银",
        "main_subject": "1002",
        "main_business_type": "货款",
    }


def wait_exit():
    try:
        input("按回车退出...")
    except (KeyboardInterrupt, EOFError):
        print()
        print("已退出。")


def main(argv=None):
    args = parse_args(argv)
    config = load_config(args.config)
    account = get_test_account(config, args.bank_label)

    recorder = RunStateRecorder(command="receipt-detail", config=config)

    print_header(detail_bank_account_no(account), args, args.start_delay)
    print()
    print(f"请在 {args.start_delay:g} 秒内切到 NC 收款单窗口...")
    wait_started_at = time.perf_counter()
    time.sleep(max(float(args.start_delay), 0))
    print("开始测试。")
    run_started_at = time.perf_counter()
    timings = StepTimer()
    timings.add("startup.wait-before-run", time.perf_counter() - wait_started_at)

    report: dict[str, object] = {
        "launcher": "receipt_detail_entry.py",
        "bank_label": args.bank_label,
        "account": detail_bank_account_no(account),
        "mode": _mode_from_args(args),
        "fee_amount": args.fee_amount if args.fee_only else None,
        "stop_hotkey": STOP_HOTKEY,
        "start_delay_seconds": args.start_delay,
    }
    try:
        result = run_detail_trial(config, account, args, report, timings, recorder)
        report["total_seconds"] = round(time.perf_counter() - run_started_at, 3)
        report["timings"] = timings.items
        print()
        print_summary(report)
        print()
        if not args.no_wait:
            wait_exit()
        if args.json_output:
            _print_envelope(report, result, run_started_at)
        _finish_from_result(recorder, result, report)
        return result
    except Exception as exc:
        report.update(
            {
                "ok": False,
                # 给人看的主原因：业务可读，不直接抛 Python 异常类名/栈。
                "reason": (
                    "明细写入未完成：运行中发生未预期错误，已停止，未保存、未暂存；"
                    "请确认 NC 停在收款单自制录入界面、无参照窗口遮挡后重试。"
                ),
                # exception 仅作环境类错误的内部分类标记，不作为给用户的主消息文案。
                "exception": type(exc).__name__,
                # 技术细节仅作开发诊断，不作为给用户的主消息。
                "error_detail": f"{type(exc).__name__}: {exc}",
                "error_traceback": traceback.format_exc(),
            }
        )
        report["total_seconds"] = round(time.perf_counter() - run_started_at, 3)
        report["timings"] = timings.items
        print()
        print_summary(report)
        print()
        if not args.no_wait:
            wait_exit()
        if args.json_output:
            _print_envelope(report, 1, run_started_at)
        recorder.finish("failed", error=str(exc))
        return 1


def _mode_from_args(args):
    if args.cleanup_extra_rows_only:
        return "cleanup-only"
    return "fee-only" if args.fee_only else "main-line"


def _finish_from_result(recorder, result, report):
    """将退出码映射到 RunStateRecorder.finish 状态,不改变返回值。"""
    if report.get("stopped_by_hotkey"):
        recorder.finish("aborted")
    elif result == 0 and report.get("ok"):
        recorder.finish("success")
    else:
        recorder.finish(
            "failed", error=report.get("reason") or report.get("failed_step")
        )


def run_detail_trial(config, account, args, report, timings, recorder=None):
    if is_stop_hotkey_pressed():
        report.update(
            {
                "ok": False,
                "stopped_by_hotkey": True,
                "failed_step": "before-start",
            }
        )
        return 1

    jab = JABOperator(config)
    try:
        timings.measure("jab.ensure-started", jab.ensure_started)
        if not _check_health(jab, report, timings):
            return 1
        located = _locate_body(jab, report, timings)
        if not located.get("best"):
            return 1
        if is_stop_hotkey_pressed():
            report.update(
                {
                    "ok": False,
                    "stopped_by_hotkey": True,
                    "failed_step": "before-fill-detail",
                }
            )
            return 1

        steps = _run_selected_mode(
            jab, account, located, args, report, timings, recorder
        )
        if not args.cleanup_extra_rows_only:
            report["fill_steps"] = steps
        report["after_table"] = timings.measure(
            "body.read-after", read_body_table, jab, "after_detail_fill"
        )
        _set_final_status(args, steps, report)
        return 0 if report.get("ok") else 1
    finally:
        jab.close()


def _check_health(jab, report, timings):
    health = timings.measure("jab.health-check", check_jab_ready, jab)
    report["jab_health"] = health
    if not health.get("ok"):
        report.update(
            {
                "ok": False,
                "failed_step": "jab-health-check",
                "reason": health.get("reason"),
            }
        )
        return False
    report["header_account"] = timings.measure(
        "header.account-read",
        wait_header_account_description,
        jab,
        timeout=0.0,
    )
    return True


def _locate_body(jab, report, timings):
    located = timings.measure(
        "body.locate-initial",
        locate_receipt_body_table_cached,
        jab,
        max_rows=3,
    )
    report["table_candidates"] = located.get("candidates", [])[:5]
    if not located.get("best"):
        report.update({"ok": False, "failed_step": "locate-body-table"})
        return located
    report["before_table"] = timings.measure(
        "body.read-before", read_body_table, jab, "before_detail_fill"
    )
    return located


def _run_selected_mode(jab, account, located, args, report, timings, recorder=None):
    if args.cleanup_extra_rows_only:
        if recorder is not None:
            recorder.set_stage("删多余行", step_index=1, total_steps=1)
        delete_extra = timings.measure(
            "cleanup.rows-after-first",
            cleanup_rows_after_first,
            jab,
            located,
        )
        report["extra_row_delete"] = delete_extra
        report["fill_steps"] = []
        if not delete_extra.get("ok"):
            report["failed_step"] = "cleanup-extra-rows"
            if recorder is not None:
                recorder.event(
                    "cleanup-extra-rows-failed", error=delete_extra.get("reason")
                )
        return []
    if args.fee_only:
        return _run_fee_mode(jab, located, args, report, timings, recorder)
    return _run_main_line_mode(jab, account, located, report, timings, recorder)


def _run_fee_mode(jab, located, args, report, timings, recorder=None):
    if recorder is not None:
        recorder.set_stage("手续费行", step_index=1, total_steps=1)
    add_row, steps, clear_account, delete_extra = timings.measure(
        "fee.total",
        run_fee_only,
        jab,
        located,
        args.fee_amount,
    )
    for item in delete_extra.get("timings") or add_row.get("timings") or []:
        timings.add(item.get("name"), item.get("seconds") or 0)
    report["fee_row_add"] = add_row
    report["fee_account_clear"] = clear_account
    report["extra_row_delete"] = delete_extra
    if not add_row.get("ok"):
        report["failed_step"] = "add-fee-row"
        if recorder is not None:
            recorder.event("add-fee-row-failed", error=add_row.get("reason"))
    elif not all(bool(step.get("ok")) for step in steps):
        report["failed_step"] = "fill-fee-line"
    elif not clear_account.get("ok"):
        report["failed_step"] = "clear-fee-account"
    elif not delete_extra.get("ok"):
        report["failed_step"] = "delete-extra-row"
    if recorder is not None:
        recorder.update_counts(fee_steps=len(steps))
    return steps


def _run_main_line_mode(jab, account, located, report, timings, recorder=None):
    if recorder is not None:
        recorder.set_stage("明细主行写入", step_index=1, total_steps=1)
    steps = timings.measure(
        "main.write-line",
        write_detail_line_by_screen,
        jab,
        build_business(account),
        located,
    )
    before_table = report.get("before_table")
    before_rows = int(
        before_table.get("row_count") or 0 if isinstance(before_table, dict) else 0
    )
    refreshed_after_main = timings.measure(
        "main.locate-after-write",
        locate_receipt_body_table_cached,
        jab,
        cached=located,
        max_rows=5,
    )
    if recorder is not None:
        recorder.set_stage("删多余行", step_index=2, total_steps=2)
    delete_extra = timings.measure(
        "main.delete-extra-after-write",
        delete_extra_row_if_present,
        jab,
        refreshed_after_main,
        expected_rows=before_rows,
    )
    report["extra_row_delete"] = delete_extra
    if not delete_extra.get("ok"):
        report["failed_step"] = "delete-extra-row"
        if recorder is not None:
            recorder.event("delete-extra-row-failed", error=delete_extra.get("reason"))
    if recorder is not None:
        recorder.update_counts(main_steps=len(steps))
    return steps


def _set_final_status(args, steps, report):
    if args.cleanup_extra_rows_only:
        report["ok"] = not report.get("failed_step")
    else:
        report["ok"] = (
            all(bool(step.get("ok")) for step in steps)
            and not report.get("failed_step")
            and bool(steps)
        )
    if not report["ok"]:
        report["failed_step"] = report.get("failed_step") or (
            "fill-fee-line" if args.fee_only else "fill-detail-line"
        )


def _print_envelope(report: dict, exit_code: object, run_started_at: float) -> None:
    """在 stdout 最后一行打印结构化结果信封（§1.2 契约）。"""
    ok = bool(report.get("ok", False))
    ec = int(exit_code) if isinstance(exit_code, int) else (0 if ok else 1)
    steps = report.get("fill_steps") or []
    items = []
    for i, step in enumerate(steps):
        outcome = "success" if step.get("ok") else "failed"
        items.append(
            {
                "ref": step.get("name") or step.get("field") or str(i),
                "outcome": outcome,
                "reason": step.get("reason") or step.get("error") or "",
            }
        )
    if not items:
        # cleanup-only 或无步骤时生成单条汇总 item
        outcome = "success" if ok else "failed"
        items = [
            {
                "ref": report.get("mode") or "main",
                "outcome": outcome,
                "reason": report.get("reason") or "",
            }
        ]
    succeeded = sum(1 for it in items if it["outcome"] == "success")
    failed = sum(1 for it in items if it["outcome"] == "failed")
    elapsed = round(time.perf_counter() - run_started_at, 3)
    err_msg = report.get("reason") or report.get("exception") or ""
    err_category = "none"
    if not ok:
        if report.get("stopped_by_hotkey"):
            err_category = "aborted"
        elif report.get("exception"):
            err_category = "environment"
        else:
            err_category = "business"
    envelope = {
        "ok": ok,
        "command": "receipt-detail",
        "exit_code": ec,
        "summary": {
            "total": len(items),
            "succeeded": succeeded,
            "failed": failed,
            "skipped": 0,
        },
        "items": items,
        "error": {"category": err_category, "message": str(err_msg)},
        "elapsed_s": elapsed,
        "resumable": {"can_resume": False, "resume_command": None},
    }
    print(json.dumps(envelope, ensure_ascii=True))


if __name__ == "__main__":
    raise SystemExit(main())
