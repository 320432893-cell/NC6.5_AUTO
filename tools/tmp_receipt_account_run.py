import argparse
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_entry import ReceiptEntryConfig  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready, print_jab_health_failure  # noqa: E402
from tools.receipt_account_reference_try import (  # noqa: E402
    STOP_HOTKEY,
    get_clipboard_text,
    run_full_account_reference,
    set_clipboard_text,
)

TEST_EXCEL_ROW = "手工测试"
DEFAULT_TEST_BANK_LABEL = "招行"
START_DELAY_SECONDS = 2


def parse_args():
    parser = argparse.ArgumentParser(description="收款单表头账号参照临时搜索脚本。")
    parser.add_argument(
        "--bank-label",
        default=DEFAULT_TEST_BANK_LABEL,
        help="测试银行标签，默认 招行；实际账号从 config.json 读取。",
    )
    parser.add_argument(
        "--currency",
        default="人民币",
        help="候选账号币种，默认 人民币。",
    )
    return parser.parse_args()


def get_test_account(config, bank_label):
    receipt_config = ReceiptEntryConfig(config)
    account = receipt_config.account_for_bank(bank_label)
    if account:
        return account
    raise RuntimeError(f"config.json 中找不到银行账户映射：{bank_label}")


def print_header(bank_label, account_no=None):
    print("测试功能：收款单表头【收款银行账户】参照搜索并带回账号")
    print()
    print("测试数据来源：")
    print(f"1. 试验行：{TEST_EXCEL_ROW}（用户指定账号测试）")
    print(f"2. 银行标签：{bank_label}（来自 config.json 映射）")
    if account_no:
        print(
            f"3. 搜索账号：{account_no}（来自 config.json 的 receipt_entry.accounts 映射）"
        )
    else:
        print("3. 搜索账号：启动后从 config.json 的 receipt_entry.accounts 映射读取")
    print()
    print("前置条件：")
    print("1. NC 已停在收款单自制录入界面")
    print("2. 【收款银行账户】的【使用权参照】窗口已经打开")
    print("3. 【使用权参照】窗口必须是当前前台窗口")
    print()
    print("本脚本会做：")
    print("1. 按候选账号顺序在参照窗口搜索账号")
    print("2. 每个候选都读取搜索结果表")
    print("3. 第一个有结果的候选会选中第一行")
    print("4. 点【确定】把账号带回表头")
    print()
    print("不会做：保存、暂存、填写明细、关闭收款单")
    print(f"紧急停止：按 {STOP_HOTKEY}")
    print(f"启动后等待：{START_DELAY_SECONDS} 秒，用来切到【使用权参照】窗口")
    print("=" * 60)


def print_summary(report):
    print()
    print("测试结果：")
    attempts = report.get("attempts") or []
    if attempts:
        print("候选尝试：")
        for item in attempts:
            table = item.get("table") or {}
            print(
                "  - "
                f"{item.get('candidate')}: "
                f"{'成功' if item.get('ok') else '失败'}，"
                f"结果行数={table.get('row_count')}，"
                f"失败位置={item.get('failed_step')}"
            )
    if report.get("ok"):
        if report.get("matched_candidate"):
            print(f"命中候选：{report.get('matched_candidate')}")
        if report.get("confirmed") is False:
            print("成功：账号已在参照窗口搜索，并选中第一行。")
            print(
                "本轮按测试要求没有点击【确定】，请先确认搜索框里的账号是否完整正确。"
            )
        else:
            print("成功：账号已在参照窗口搜索、选中第一行，并点击【确定】。")
        return

    failed_step = report.get("failed_step")
    reason = report.get("reason")
    if report.get("stopped_by_hotkey"):
        print(f"已停止：检测到紧急停止键 {STOP_HOTKEY}。")
        print(f"停止位置：{failed_step}")
        return

    if failed_step == "find-dialog":
        print("失败：没有找到可用的【使用权参照】窗口。")
        print("本次没有输入账号，也没有点击【确定】。")
        print("请先在 NC 里打开【收款银行账户】参照窗口，再运行本脚本。")
        return

    if failed_step == "jab-health-check":
        print("失败：NC 当前没有可用的 Java Access Bridge 根窗口。")
        print("本次没有输入账号，也没有点击【确定】。")
        health = report.get("jab_health") or {}
        if isinstance(health, dict):
            print_jab_health_failure(health)
        return

    if failed_step == "check-foreground":
        print("失败：【使用权参照】不是当前前台窗口。")
        print("本次没有输入账号，也没有点击【确定】。")
        print("请先点一下【使用权参照】窗口，让它在最前面，再运行本脚本。")
        return

    if failed_step == "type-search":
        print("失败：账号搜索输入没有完成。")
        print(f"原因：{reason or (report.get('search') or {}).get('reason')}")
        return

    if failed_step == "clipboard-precheck":
        print("失败：Windows 剪贴板预检失败。")
        print("本次没有输入账号，也没有点击【确定】。")
        print(f"原因：{reason}")
        return

    if failed_step == "read-table":
        print("失败：搜索后没有读到结果表或结果表为空。")
        table = report.get("table") or {}
        if isinstance(table, dict):
            print(f"结果行数：{table.get('row_count')}")
            print(f"原因：{table.get('reason')}")
        return

    if failed_step == "select-first":
        print("失败：已读到结果表，但选中第一行失败。")
        return

    if failed_step == "confirm":
        print("失败：第一行已选中，但点击【确定】失败。")
        return

    if report.get("exception"):
        print(f"脚本异常：{report.get('exception')}")
        print(f"原因：{report.get('reason')}")
        return

    print("失败：未知位置。")
    print(f"失败位置：{failed_step}")
    print(f"原因：{reason}")


def wait_exit():
    try:
        input("按回车退出...")
    except (KeyboardInterrupt, EOFError):
        print()
        print("已退出。")


def check_clipboard_ready(account_no):
    before = get_clipboard_text()
    set_clipboard_text(account_no)
    after = get_clipboard_text()
    if before is not None:
        set_clipboard_text(before)
    return {
        "ok": after == account_no,
        "before_had_text": before is not None,
        "after": after,
    }


def build_account_candidates(account_no):
    raw = str(account_no).strip()
    candidates = [raw]
    for suffix in ("RMB", "USD", "CNY"):
        if raw.upper().endswith(suffix):
            continue
        candidates.append(f"{raw}{suffix}")
    deduped = []
    seen = set()
    for value in candidates:
        if value and value not in seen:
            deduped.append(value)
            seen.add(value)
    return deduped


def run_candidate_account_reference(jab, candidates):
    attempts = []
    for candidate in candidates:
        result = run_full_account_reference(
            jab,
            account=candidate,
            press_enter=True,
            check_timeout=1.0,
            poll_timeout=3.0,
            confirm_selection=True,
        )
        attempts.append({"candidate": candidate, **result})
        if result.get("ok"):
            return {
                "ok": True,
                "matched_candidate": candidate,
                "attempts": attempts,
                **result,
            }
        if result.get("failed_step") not in ("read-table",):
            return {
                "ok": False,
                "failed_step": result.get("failed_step"),
                "reason": result.get("reason"),
                "attempts": attempts,
            }
    return {
        "ok": False,
        "failed_step": "read-table",
        "reason": "所有候选账号都没有搜索结果",
        "attempts": attempts,
    }


def main():
    args = parse_args()
    config = load_config(str(ROOT / "config.json"))
    account = get_test_account(config, args.bank_label)
    candidates = account.nc_candidates(args.currency)
    if not candidates:
        candidates = build_account_candidates(account.account_no)
    print_header(args.bank_label, account.account_no)
    print("候选搜索顺序：")
    for candidate in candidates:
        print(f"  - {candidate}")
    print()
    print(f"请在 {START_DELAY_SECONDS} 秒内切到【使用权参照】窗口...")
    time.sleep(START_DELAY_SECONDS)
    print("开始测试。")

    report: dict[str, object] = {
        "launcher": "tmp_receipt_account_run.py",
        "excel_row": TEST_EXCEL_ROW,
        "bank_label": args.bank_label,
        "currency": args.currency,
        "account": account.account_no,
        "account_candidates": candidates,
        "stop_hotkey": STOP_HOTKEY,
        "start_delay_seconds": START_DELAY_SECONDS,
    }
    try:
        clipboard_check = check_clipboard_ready(max(candidates, key=len))
        report["clipboard_precheck"] = clipboard_check
        if not clipboard_check.get("ok"):
            report.update(
                {
                    "ok": False,
                    "failed_step": "clipboard-precheck",
                    "reason": "Windows 剪贴板预检失败，未开始操作 NC。",
                }
            )
            print()
            print_summary(report)
            print()
            wait_exit()
            return 1

        jab = JABOperator(config)
        try:
            jab.ensure_started()
            health = check_jab_ready(jab)
            report["jab_health"] = health
            if not health.get("ok"):
                report.update(
                    {
                        "ok": False,
                        "failed_step": "jab-health-check",
                        "reason": health.get("reason"),
                    }
                )
                print()
                print_summary(report)
                print()
                wait_exit()
                return 1
            report.update(run_candidate_account_reference(jab, candidates))
        finally:
            jab.close()
    except Exception as exc:
        report.update(
            {
                "ok": False,
                "exception": type(exc).__name__,
                "reason": str(exc),
                "traceback": traceback.format_exc(),
            }
        )

    print()
    print_summary(report)
    print()
    wait_exit()
    return 0 if report.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
