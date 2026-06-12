# 职责：编排收款单完整流程测试入口，消费 ReceiptPlanRow 跑开单/表头/明细/手续费/可选保存
# 不做什么：不做录入前 NC 查重，不复用历史 T0 保存脚本，不把保存设为默认行为
# 允许依赖层：core 收款计划/配置/JAB、tools 下已正式化的收款开单/明细/查询组件
# 谁不应该 import：core 层模块不应 import 本入口；凭证批量模块不应 import

import argparse
from collections import Counter
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
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table_cached  # noqa: E402
from tools.receipt_detail_row_cleanup import delete_extra_row_if_present  # noqa: E402
from tools.receipt_detail_rows import StepTimer, run_fee_only  # noqa: E402
from tools.receipt_detail_writer import write_detail_line_by_screen  # noqa: E402
from tools.receipt_new_probe import (  # noqa: E402
    collect_receipt_new_windows,
    detect_self_made_entry_state,
)
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    detect_existing_self_made_entry,
    fill_header,
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
    parser.add_argument("--limit", type=int, default=1, help="最多测试几行，默认 1")
    parser.add_argument(
        "--write-plan-sheet",
        action="store_true",
        help="运行前先把本地预检结果写入 Sheet2",
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

    report = {
        "launcher": "receipt_full_flow_entry.py",
        "mode": "save" if args.save else "no-save",
        "query_after_save": bool(args.query_after_save),
        "started_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "rows": [],
    }
    started = time.perf_counter()
    workbook = ReceiptEntryWorkbook(config, excel_path=args.excel_path)
    plan_rows, issues, summary = workbook.build_local_plan(
        write_sheet=args.write_plan_sheet
    )
    selected_rows = select_plan_rows(plan_rows, issues, args)
    report["local_plan"] = {
        "summary": summary,
        "issue_count": len(issues),
        "selected_rows": [row.row for row in selected_rows],
        "write_plan_sheet": bool(args.write_plan_sheet),
    }
    if not selected_rows:
        report.update(
            {
                "ok": False,
                "reason": "没有可测试的通过预检计划行",
                "total_seconds": round(time.perf_counter() - started, 3),
            }
        )
        print_report(report, args)
        return 2

    if not args.json:
        print("收款单完整流程测试入口")
        print(f"模式：{'真实保存' if args.save else '不保存，停在保存前'}")
        print(f"计划行：{[row.row for row in selected_rows]}")
        print(f"启动后等待 {args.start_delay:g} 秒，请切到 NC 收款单录入页面。")
    time.sleep(max(args.start_delay, 0))

    exit_code = 0
    for row in selected_rows:
        row_report = run_one_row(config, row, save_enabled=args.save)
        report["rows"].append(row_report)
        if not row_report.get("ok"):
            exit_code = 1
            break
    if args.query_after_save and report["rows"] and exit_code == 0:
        report["post_query"] = build_post_query_defer_report(selected_rows)

    report["ok"] = exit_code == 0
    report["total_seconds"] = round(time.perf_counter() - started, 3)
    print_report(report, args)
    return exit_code


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
    if args.excel_row is not None:
        runnable = [row for row in runnable if row.row == args.excel_row]
    limit = max(int(args.limit or 0), 0)
    if limit:
        runnable = runnable[:limit]
    return runnable


def run_one_row(config, row, save_enabled=False):
    timings = StepTimer()
    row_report = {
        "excel_row": row.row,
        "plan_row": serializable(asdict(row)),
        "steps": [],
        "save_enabled": bool(save_enabled),
    }
    business = business_from_plan_row(row)
    open_step = timings.measure("open.self-made", open_self_made_entry, config)
    row_report["steps"].append({"name": "open-self-made", **open_step})
    if not open_step.get("ok"):
        return fail(row_report, "open-self-made", timings, open_step.get("reason"))

    jab = JABOperator(config)
    try:
        timings.measure("jab.ensure-started", jab.ensure_started)
        health = timings.measure("jab.health-check", check_jab_ready, jab)
        row_report["jab_health"] = health
        if not health.get("ok"):
            return fail(row_report, "jab-health-check", timings, health.get("reason"))
        header_steps = timings.measure(
            "header.fill",
            fill_header,
            jab,
            business,
            False,
            True,
        )
        row_report["header_steps"] = header_steps
        if any(not step.get("ok") for step in header_steps):
            return fail(row_report, "header-fill", timings, "表头字段写入失败")
        account_check = timings.measure(
            "header.account-readback",
            wait_header_account_description,
            jab,
            5.0,
        )
        row_report["header_account"] = account_check
        if not account_check.get("accepted"):
            return fail(
                row_report, "header-account-readback", timings, "收款银行账户为空"
            )

        located = timings.measure(
            "body.locate",
            locate_receipt_body_table_cached,
            jab,
            max_rows=5,
        )
        row_report["table_candidates"] = located.get("candidates", [])[:5]
        if not located.get("best"):
            return fail(row_report, "locate-body-table", timings, "未定位到明细表")
        row_report["before_table"] = timings.measure(
            "body.read-before", read_body_table, jab, "before_detail_fill"
        )
        detail_steps = timings.measure(
            "detail.main-line",
            write_detail_line_by_screen,
            jab,
            business,
            located,
        )
        row_report["detail_steps"] = detail_steps
        if not all(step.get("ok") for step in detail_steps):
            return fail(row_report, "detail-main-line", timings, "明细主行写入失败")
        row_report["extra_row_delete"] = timings.measure(
            "detail.delete-extra-after-main",
            delete_extra_row_if_present,
            jab,
            located,
            1,
        )
        if row.fee > 0:
            add_row, fee_steps, clear_account, delete_extra = timings.measure(
                "detail.fee-line",
                run_fee_only,
                jab,
                located,
                str(row.fee),
            )
            row_report["fee_row_add"] = add_row
            row_report["fee_steps"] = fee_steps
            row_report["fee_account_clear"] = clear_account
            row_report["fee_extra_row_delete"] = delete_extra
            if (
                not add_row.get("ok")
                or not all(step.get("ok") for step in fee_steps)
                or not clear_account.get("ok")
                or not delete_extra.get("ok")
            ):
                return fail(row_report, "detail-fee-line", timings, "手续费行处理失败")
        else:
            row_report["fee_skipped"] = {
                "ok": True,
                "reason": "手续费为 0，跳过手续费行",
            }
        row_report["after_table"] = timings.measure(
            "body.read-after", read_body_table, jab, "after_detail_fill"
        )
        if save_enabled:
            save_result = timings.measure("save.ctrl-s", save_receipt_by_ctrl_s, jab)
            row_report["save"] = save_result
            if not save_result.get("ok"):
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


def build_post_query_defer_report(rows):
    by_org = Counter(row.organization_code for row in rows)
    date_from = min(row.receipt_date for row in rows).isoformat()
    date_to = max(row.receipt_date for row in rows).isoformat()
    return {
        "ok": False,
        "deferred": True,
        "reason": "查询后验编排下一步接入；当前入口先验证保存前/保存动作",
        "organizations": dict(by_org),
        "date_from": date_from,
        "date_to": date_to,
    }


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
