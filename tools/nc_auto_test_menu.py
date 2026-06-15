# -*- coding: utf-8 -*-
"""Windows interactive test menu for NC automation."""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
TOOLS_DIR = PROJECT_DIR / "tools"
PYTHON = sys.executable


def clear_screen() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def ask(prompt: str, default: str = "") -> str:
    value = input(prompt).strip()
    return value or default


def pause(prompt: str = "按回车继续...") -> None:
    input(prompt)


def run_logged(title: str, risk: str, command: list[str]) -> int:
    print()
    print(f"功能：{title}")
    print(f"风险：{risk}")
    print(f"开始时间：{datetime.now():%Y-%m-%d %H:%M:%S}")
    print("命令：" + " ".join(f'"{part}"' if " " in part else part for part in command))
    print()
    pause("请确认当前 NC/Excel 状态满足该功能前置条件；按回车开始，Ctrl+C 取消...")
    try:
        result = subprocess.run(command, cwd=PROJECT_DIR, check=False)
        return_code = result.returncode
    except KeyboardInterrupt:
        print()
        print("已取消。")
        return_code = 130
    print()
    print(f"结束时间：{datetime.now():%Y-%m-%d %H:%M:%S}")
    print(f"退出码：{return_code}")
    pause()
    return return_code


def prompt_range() -> list[str]:
    args: list[str] = []
    limit = ask("仅处理前 N 条，直接回车不限制：")
    if limit:
        args += ["--limit", limit]
    start_row = ask("Excel 起始行，直接回车不限制：")
    if start_row:
        args += ["--start-row", start_row]
    end_row = ask("Excel 结束行，直接回车不限制：")
    if end_row:
        args += ["--end-row", end_row]
    return args


def prompt_receipt_full_flow() -> list[str]:
    args: list[str] = []
    row = ask("指定 Sheet1 行号，直接回车自动取第一条通过预检行：")
    if row:
        args += ["--excel-row", row]
    limit = ask("最多测试几行，直接回车默认 1：", "1")
    args += ["--limit", limit]
    start_delay = ask("启动前等待秒数，直接回车默认 2：", "2")
    args += ["--start-delay", start_delay]
    return args


def prompt_receipt_query() -> list[str]:
    org_code = ask("主体编码，直接回车默认 A001：", "A001")
    date_from = ask("开始日期 YYYY-MM-DD，直接回车默认 2026-05-01：", "2026-05-01")
    date_to = ask("结束日期 YYYY-MM-DD，直接回车默认 2026-06-02：", "2026-06-02")
    max_rows = ask("最多读取行数，直接回车默认 600：", "600")
    max_cols = ask("最多读取列数，直接回车默认 140：", "140")
    return [
        "--org-code",
        org_code,
        "--date-from",
        date_from,
        "--date-to",
        date_to,
        "--confirm",
        "--max-rows",
        max_rows,
        "--max-cols",
        max_cols,
    ]


def project_check() -> int:
    return run_logged(
        "工程检查 changed",
        "A 只读/本地检查",
        [PYTHON, str(TOOLS_DIR / "check.py"), "changed"],
    )


def voucher_plan() -> int:
    args = prompt_range()
    return run_logged(
        "凭证计划预览",
        "A 只读/不保存",
        [PYTHON, str(TOOLS_DIR / "jab_batch.py"), "plan", *args],
    )


def voucher_generate() -> int:
    print()
    print("[高风险] 该功能会真实点击 NC 生成并保存凭证，也会写 Excel 状态。")
    print("Python 入口仍可能再次要求输入 yes。")
    if ask("确认继续请输入 SAVE：").upper() != "SAVE":
        print("已取消。")
        pause()
        return 0
    args = prompt_range()
    max_batches = ask("最多执行几个批次，直接回车不限制：")
    if max_batches:
        args += ["--max-batches", max_batches]
    if ask("是否记录 perf JSONL？输入 y 开启：").lower() == "y":
        args += ["--perf", "--perf-label", "voucher-generate-menu"]
    return run_logged(
        "凭证真实生成并保存",
        "D 真实保存",
        [PYTHON, str(TOOLS_DIR / "jab_batch.py"), "generate", *args],
    )


def voucher_backfill() -> int:
    args = prompt_range()
    if ask("不自动从待生成切已生成？输入 y 开启：").lower() == "y":
        args += ["--no-backfill-auto-switch"]
    return run_logged(
        "凭证回填凭证号",
        "B 写 Excel + 读取 NC",
        [PYTHON, str(TOOLS_DIR / "jab_batch.py"), "backfill", *args],
    )


def voucher_switch_generated() -> int:
    args: list[str] = []
    generated_date = ask("目的业务日期 YYYY-MM-DD，直接回车用配置或当天：")
    if generated_date:
        args += ["--generated-date", generated_date]
    if ask("是否记录 perf JSONL？输入 y 开启：").lower() == "y":
        args += ["--perf", "--perf-label", "switch-generated-menu"]
    return run_logged(
        "凭证切到已生成/正式单据列表",
        "C 真实 NC 不保存",
        [PYTHON, str(TOOLS_DIR / "jab_batch.py"), "switch-generated", *args],
    )


def receipt_local_plan() -> int:
    return run_logged(
        "收款单本地预检",
        "A 只读/不写 Excel/不碰 NC",
        [PYTHON, str(TOOLS_DIR / "receipt_entry_check.py")],
    )


def receipt_write_sheet2() -> int:
    print()
    print("[写 Excel] 该功能会维护 Sheet2：收款单自动化结果。")
    if ask("确认写 Sheet2 请输入 WRITE：").upper() != "WRITE":
        print("已取消。")
        pause()
        return 0
    args: list[str] = []
    if ask("是否跳过异常行继续生成可运行计划？输入 y 开启：").lower() == "y":
        args += ["--validation-mode", "skip_invalid_rows"]
    return run_logged(
        "收款单本地预检并写 Sheet2",
        "B 写 Excel",
        [PYTHON, str(TOOLS_DIR / "receipt_entry_check.py"), *args, "--write"],
    )


def receipt_full_no_save() -> int:
    args = prompt_receipt_full_flow()
    if ask("运行前是否写 Sheet2 本地预检结果？输入 y 开启：").lower() == "y":
        args += ["--write-plan-sheet"]
    return run_logged(
        "收款单完整流程测试（不保存）",
        "C 真实 NC 不保存",
        [PYTHON, str(TOOLS_DIR / "receipt_full_flow_entry.py"), *args],
    )


def receipt_full_save() -> int:
    print()
    print("[高风险] 该功能会开自制收款单，填写表头/明细/手续费，并真实保存。")
    print("保存前请确认测试单据可清理、NC 已在收款单录入页、Excel/WPS 未占用。")
    if ask("确认真实保存请输入 SAVE：").upper() != "SAVE":
        print("已取消。")
        pause()
        return 0
    args = prompt_receipt_full_flow()
    return run_logged(
        "收款单完整流程真实保存",
        "D 真实保存",
        [
            PYTHON,
            str(TOOLS_DIR / "receipt_full_flow_entry.py"),
            *args,
            "--save",
            "--yes-i-understand",
        ],
    )


def receipt_no_save_trial() -> int:
    print()
    print("[真实 NC 不保存] 要求 NC 在收款单录入页，脚本会试填字段，但不会保存/暂存。")
    row = ask("请输入 Excel 行号：")
    if not row:
        print("已取消。")
        pause()
        return 0
    args = [row, "--open-self-made"]
    if ask("是否继续填明细？输入 y 开启：").lower() == "y":
        args += ["--fill-detail"]
    return run_logged(
        "收款单单行无保存试填",
        "C 真实 NC 不保存",
        [PYTHON, str(TOOLS_DIR / "receipt_self_made_fill_trial.py"), *args],
    )


def receipt_query_read() -> int:
    args = prompt_receipt_query()
    return run_logged(
        "收款单查询读取结果",
        "C 读取 NC 不写 Excel",
        [PYTHON, str(TOOLS_DIR / "receipt_query_fill.py"), *args, "--read-results"],
    )


def receipt_query_write_back() -> int:
    print()
    print("[写 Excel] 这是历史查重/诊断入口，会写 Sheet1 状态列，不是新批量录入主线。")
    if ask("确认写回 Sheet1 请输入 WRITEBACK：").upper() != "WRITEBACK":
        print("已取消。")
        pause()
        return 0
    args = prompt_receipt_query()
    args += ["--dry-run-match", "--write-back"]
    if ask("是否覆盖已有状态列？输入 y 开启：").lower() == "y":
        args += ["--include-filled-status"]
    return run_logged(
        "收款单历史查重写回 Sheet1",
        "B 写 Excel + 读取 NC",
        [PYTHON, str(TOOLS_DIR / "receipt_query_fill.py"), *args],
    )


def receipt_detail_menu() -> int:
    while True:
        clear_screen()
        print("收款单明细测试入口")
        print()
        print("前置条件：")
        print(" 1. NC 已停在【收款单自制录入】界面")
        print(" 2. 当前没有参照窗口或提示框")
        print(" 3. 本入口不会保存、不会暂存、不会关闭收款单")
        print()
        print("请选择测试功能：")
        print(" 1. 写明细主行")
        print(" 2. 写手续费行（默认手续费 10）")
        print(" 3. 只清理第 1 行以外的多余行")
        print(" 4. 显示命令帮助")
        print(" 0. 返回主菜单")
        print()
        choice = ask("请输入编号后回车：")
        if choice == "1":
            run_logged(
                "收款单明细主行测试",
                "C 真实 NC 不保存",
                [PYTHON, str(TOOLS_DIR / "receipt_detail_entry.py")],
            )
        elif choice == "2":
            fee_amount = ask("请输入手续费金额，直接回车默认 10：", "10")
            run_logged(
                "收款单手续费行测试",
                "C 真实 NC 不保存",
                [
                    PYTHON,
                    str(TOOLS_DIR / "receipt_detail_entry.py"),
                    "--fee-only",
                    "--fee-amount",
                    fee_amount,
                ],
            )
        elif choice == "3":
            run_logged(
                "收款单清理多余明细行",
                "C 真实 NC 不保存",
                [
                    PYTHON,
                    str(TOOLS_DIR / "receipt_detail_entry.py"),
                    "--cleanup-extra-rows-only",
                ],
            )
        elif choice == "4":
            run_logged(
                "收款单明细命令帮助",
                "A 只读/本地帮助",
                [PYTHON, str(TOOLS_DIR / "receipt_detail_entry.py"), "--help"],
            )
        elif choice == "0":
            return 0
        else:
            print("输入无效，请重新选择。")
            pause()


def main_menu() -> int:
    actions = {
        "1": project_check,
        "2": voucher_plan,
        "3": voucher_generate,
        "4": voucher_backfill,
        "5": voucher_switch_generated,
        "6": receipt_local_plan,
        "7": receipt_write_sheet2,
        "8": receipt_full_no_save,
        "9": receipt_full_save,
        "10": receipt_no_save_trial,
        "11": receipt_detail_menu,
        "12": receipt_query_read,
        "13": receipt_query_write_back,
    }
    while True:
        clear_screen()
        print("NC 自动化项目级测试入口")
        print()
        print(f"Python: {PYTHON}")
        print(f"项目目录: {PROJECT_DIR}")
        print()
        print("风险分组：")
        print(" A. 只读/本地检查：不写 Excel，不点击保存")
        print(" B. 写 Excel：会写 Sheet2 或 Sheet1 状态列")
        print(" C. 真实 NC 不保存：会点击/填写当前 NC 页面，但不保存业务单据")
        print(" D. 真实保存：会保存凭证或收款单，必须二次确认")
        print()
        print("请选择功能：")
        print("  1. 工程检查 tools/check.py changed")
        print("  2. 凭证计划预览（不保存）")
        print("  3. 凭证真实生成并保存")
        print("  4. 凭证回填凭证号")
        print("  5. 凭证切到已生成/正式单据列表（可记录 perf）")
        print("  6. 收款单本地预检（不写 Excel）")
        print("  7. 收款单本地预检并写 Sheet2")
        print("  8. 收款单完整流程测试（默认不保存，消费 ReceiptPlanRow）")
        print("  9. 收款单完整流程真实保存（高风险）")
        print(" 10. 收款单单行无保存试填（旧分阶段入口）")
        print(" 11. 收款单明细测试子菜单（主行/手续费/删多余行，不保存）")
        print(" 12. 收款单查询读取结果（不写 Excel）")
        print(" 13. 收款单历史查重写回 Sheet1 状态列")
        print("  0. 退出")
        print()
        choice = ask("请输入编号后回车：")
        if choice == "0":
            return 0
        action = actions.get(choice)
        if action is None:
            print("输入无效，请重新选择。")
            pause()
            continue
        action()


def main() -> int:
    parser = argparse.ArgumentParser(description="NC 自动化中文测试菜单")
    parser.add_argument(
        "--detail-menu", action="store_true", help="打开收款单明细测试子菜单"
    )
    args = parser.parse_args()
    if args.detail_menu:
        return receipt_detail_menu()
    return main_menu()


if __name__ == "__main__":
    raise SystemExit(main())
