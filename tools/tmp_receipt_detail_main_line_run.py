# 生命周期：T0 一次性（删除条件：明细主行整行屏幕写入路径确认后删除）
# 覆盖的业务阶段：收款单自制录入-明细主行整行试写
# 依赖的服务/环境：Windows Python、NC 收款单自制录入界面、Java Access Bridge
# 运行方式：python tools/tmp_receipt_detail_main_line_run.py

import argparse
import ctypes
from decimal import Decimal, InvalidOperation
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
from tools.receipt_account_reference_try import STOP_HOTKEY, is_stop_hotkey_pressed  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    read_body_table,
    wait_header_account_description,
)
from tools.tmp_receipt_cell_probe_run import (  # noqa: E402
    amount_matches,
    cell_center,
    guarded_send_ctrl_d,
    guarded_send_ctrl_i,
    guarded_send_delete,
    mouse_click,
    move_mouse,
    read_window_info,
    screen_write_amount_cell,
)

DEFAULT_TEST_BANK_LABEL = "招行"
START_DELAY_SECONDS = 2
ADD_FEE_ROW_HOTKEY = "Ctrl+I"
DETAIL_FIELDS = [
    {"col": 1, "name": "收款业务类型", "value_key": "main_business_type"},
    {"col": 3, "name": "币种", "value_key": "currency"},
    {"col": 4, "name": "收款银行账户", "value_key": "bank_account"},
    {"col": 5, "name": "科目", "value_key": "main_subject", "kind": "code_prefix"},
    {"col": 7, "name": "贷方原币金额", "value_key": "amount", "kind": "amount"},
    # 结算方式放最后：它是后置字段，不再用它提交其他字段。
    {"col": 11, "name": "结算方式", "value_key": "settlement"},
]
ACCOUNT_COL = 4
FEE_FIELDS = [
    {"col": 1, "name": "收款业务类型", "value_key": "fee_business_type"},
    {"col": 5, "name": "科目", "value_key": "fee_subject", "kind": "code_prefix"},
    {"col": 7, "name": "贷方原币金额", "value_key": "fee_amount", "kind": "amount"},
    # 手续费行也需要结算方式，且结算方式放最后。
    {"col": 11, "name": "结算方式", "value_key": "settlement"},
]


def get_test_account(config, bank_label):
    receipt_config = ReceiptEntryConfig(config)
    account = receipt_config.account_for_bank(bank_label)
    if account:
        return account
    raise RuntimeError(f"config.json 中找不到银行账户映射：{bank_label}")


def build_business(account):
    return {
        "currency": "美元",
        "bank_account": account.account_no,
        "amount": "1090",
        "settlement": "网银",
        "main_subject": "1002",
        "main_business_type": "货款",
    }


def build_fee_business(fee_amount):
    return {
        "fee_business_type": "手续费",
        "fee_subject": "660305",
        "fee_amount": str(fee_amount),
        "settlement": "网银",
    }


def parse_args():
    parser = argparse.ArgumentParser(
        description="收款单明细主行/手续费行临时试写脚本。"
    )
    parser.add_argument(
        "--fee-only",
        action="store_true",
        help="只测试手续费：Ctrl+I 增行后写新增行，不写主行。",
    )
    parser.add_argument(
        "--fee-amount",
        default="10",
        help="手续费测试金额，默认 10。",
    )
    parser.add_argument(
        "--bank-label",
        default=DEFAULT_TEST_BANK_LABEL,
        help="测试银行标签，默认 招行；实际账号从 config.json 读取。",
    )
    return parser.parse_args()


def print_header(account, args):
    if args.fee_only:
        print("测试功能：收款单手续费行从增行开始完整试写")
    else:
        print("测试功能：收款单明细主行后台填入")
    print()
    print("测试数据来源：")
    print(f"1. 银行标签：{args.bank_label}（来自 config.json 映射）")
    print(f"2. 收款银行账户：{account.account_no}")
    if args.fee_only:
        print(f"3. 手续费行：手续费 / 660305 / {args.fee_amount} / 网银")
    else:
        print("3. 明细主行：货款 / 美元 / 1002 / 1090 / 网银")
    print()
    print("前置条件：")
    print("1. NC 已停在收款单自制录入界面")
    print("2. 当前明细主行已存在；手续费模式要求主行已写好")
    print("3. 当前没有打开参照窗口或提示框")
    print()
    print("本脚本会做：")
    if args.fee_only:
        print("1. 定位 25 列明细表，并读取当前行数")
        print(f"2. 前台守卫通过后发送 {ADD_FEE_ROW_HOTKEY} 新增手续费行")
        print("3. 在新增后的最后一行写入：手续费、660305、金额、网银")
        print("4. 每写一个字段立刻读回明细表关键列")
    else:
        print("1. 读取表头【收款银行账户】状态")
        print("2. 定位 25 列明细表")
        print("3. 用受保护屏幕输入逐字段写入明细主行")
        print("4. 每写一个字段立刻读回明细表关键列")
    print()
    print("不会做：保存、暂存、关闭收款单")
    if not args.fee_only:
        print(f"增行规则：只允许手续费非零分支使用 {ADD_FEE_ROW_HOTKEY}；主行不增行")
    print("说明：写下一个字段时，会通过点击下一个目标格自然提交上一个字段。")
    print("说明：结算方式放最后；最后会点回第一个字段完成提交。")
    print("兼容性：坐标按当前 JAB 表格 bounds/行数/列数动态计算，不使用固定像素。")
    print("限制：如果用户拖动列宽、隐藏列或横向滚动，当前均分列宽算法需要升级。")
    print(f"紧急停止：按 {STOP_HOTKEY}")
    print(f"启动后等待：{START_DELAY_SECONDS} 秒，用来切到 NC 窗口")
    print("=" * 60)


def print_table_snapshot(title, snapshot):
    print(title)
    if not snapshot.get("ok"):
        print(f"  失败：{snapshot.get('reason')}")
        return
    print(f"  明细表：{snapshot.get('row_count')} 行 x {snapshot.get('col_count')} 列")
    rows = snapshot.get("rows") or []
    for row in rows[:3]:
        row_no = int(row.get("row_index", 0)) + 1
        cells = row.get("cells") or {}
        print(
            f"  第 {row_no} 行关键列："
            f"业务类型={cells.get('1')!r}, "
            f"币种={cells.get('3')!r}, "
            f"账户={cells.get('4')!r}, "
            f"科目={cells.get('5')!r}, "
            f"金额={cells.get('7')!r}, "
            f"结算={cells.get('11')!r}"
        )


def print_table_candidates(candidates):
    if not candidates:
        print("  未发现任何 JAB 表格候选。")
        return
    print("  候选表摘要：")
    for item in candidates[:5]:
        print(
            "  - "
            f"{item.get('row_count')} 行 x {item.get('col_count')} 列，"
            f"score={item.get('score')}，"
            f"窗口={((item.get('window') or {}).get('title') or '<无标题>')}"
        )
        print(f"    原因：{item.get('reasons')}")
        rows = item.get("rows") or []
        if rows:
            cells = rows[0].get("cells") or {}
            print(f"    第 1 行关键列：{cells}")


def print_fill_summary(steps):
    print("明细写入结果：")
    for step in steps:
        name = step.get("name")
        value = step.get("value")
        ok = "成功" if step.get("ok") else "失败"
        actual = step.get("actual")
        print(f"  {name}: {ok} | 期望={value!r} | 实际={actual!r}")
        if step.get("target"):
            print(f"    点击坐标：{step.get('target')}")
        if step.get("geometry"):
            geometry = step.get("geometry") or {}
            print(
                "    坐标依据："
                f"bounds={geometry.get('table_bounds')} "
                f"行数={geometry.get('row_count')} 列数={geometry.get('col_count')} "
                f"单元格={geometry.get('cell_width')}x{geometry.get('cell_height')}"
            )
        if not step.get("ok"):
            print(f"    原因：{step.get('reason')}")


def print_summary(report):
    print()
    print("测试结果：")
    if report.get("stopped_by_hotkey"):
        print(f"已停止：检测到紧急停止键 {STOP_HOTKEY}。")
        print(f"停止位置：{report.get('failed_step')}")
        return

    if report.get("exception"):
        print(f"脚本异常：{report.get('exception')}")
        print(f"原因：{report.get('reason')}")
        return

    if report.get("failed_step") == "jab-health-check":
        print("JAB 启动状态：")
        health = report.get("jab_health") or {}
        if isinstance(health, dict):
            print_jab_health_failure(health)
        print("失败：当前不能读取 NC JAB 控件树，未执行表头检查和明细写入。")
        return

    header = report.get("header_account") or {}
    print("表头账户检查：")
    if header.get("accepted"):
        print(f"  已读到账户字段：{header.get('text') or header.get('description')}")
    else:
        print("  未确认读到账户字段；本轮仍继续尝试明细写入。")

    failed_step = report.get("failed_step")
    if failed_step == "locate-body-table":
        print("明细表定位：")
        print("  失败：没有找到符合收款单明细特征的 25 列表。")
        print("  本轮已阻塞停止，没有尝试写入明细。")
        print_table_candidates(report.get("table_candidates") or [])
        print("失败：请确认当前停在【收款单自制录入界面】，且没有参照窗口/提示框遮挡。")
        return

    print_table_snapshot("写入前明细表：", report.get("before_table") or {})
    if report.get("fee_row_add") is not None:
        add_row = report.get("fee_row_add") or {}
        print("手续费增行：")
        print(
            f"  {add_row.get('hotkey')}: "
            f"{'成功' if add_row.get('ok') else '失败'} | "
            f"行数 {add_row.get('before_rows')} -> {add_row.get('after_rows')}"
        )
        pressed = add_row.get("pressed") or {}
        if pressed.get("mode"):
            print(f"  发送方式：{pressed.get('mode')}")
        if not add_row.get("ok"):
            print(f"  原因：{add_row.get('reason')}")
    if report.get("fill_steps") is not None:
        print_fill_summary(report.get("fill_steps") or [])
    if report.get("fee_account_clear") is not None:
        clear = report.get("fee_account_clear") or {}
        print("手续费账户清空：")
        if clear.get("skipped"):
            print("  跳过：手续费行账户本来就是空。")
        else:
            print(
                f"  {'成功' if clear.get('ok') else '失败'} | "
                f"清空前={clear.get('before')!r} | 清空后={clear.get('after')!r}"
            )
            if not clear.get("ok"):
                print(f"  原因：{clear.get('reason')}")
    if report.get("extra_row_delete") is not None:
        delete = report.get("extra_row_delete") or {}
        print("多余空行删除：")
        if delete.get("skipped"):
            print("  跳过：未发现需要删除的多余行。")
        else:
            print(
                f"  {'成功' if delete.get('ok') else '失败'} | "
                f"行数 {delete.get('before_rows')} -> {delete.get('after_rows')}"
            )
            if not delete.get("ok"):
                print(f"  原因：{delete.get('reason')}")
    if report.get("after_table") is not None:
        print_table_snapshot("写入后明细表：", report.get("after_table") or {})

    if report.get("ok"):
        if report.get("mode") == "fee-only":
            print("成功：手续费行已增行并写入，通过表格读回校验。")
        else:
            print("成功：明细主行字段已按屏幕输入写入，并通过表格读回校验。")
    else:
        print("失败：至少一个明细字段没有写入成功；本次没有保存、没有暂存。")


def wait_exit():
    try:
        input("按回车退出...")
    except (KeyboardInterrupt, EOFError):
        print()
        print("已退出。")


def normalize_text(value):
    return str(value or "").strip()


def normalize_amount_text(value):
    text = normalize_text(value).replace(",", "")
    if not text:
        return ""
    try:
        return str(Decimal(text).quantize(Decimal("0.01")))
    except (InvalidOperation, ValueError):
        return normalize_text(value)


def field_matches(actual, expected, kind=None):
    if kind == "amount":
        return amount_matches(actual, expected)
    if kind == "code_prefix":
        actual_text = normalize_text(actual)
        expected_text = normalize_text(expected)
        return actual_text == expected_text or actual_text.startswith(
            f"{expected_text}\\"
        )
    return normalize_text(actual) == normalize_text(expected)


def read_first_row_cells(jab):
    snapshot = read_body_table(jab, "field_readback")
    if not snapshot.get("ok"):
        return snapshot, {}
    rows = snapshot.get("rows") or []
    cells = (rows[0].get("cells") if rows else {}) or {}
    return snapshot, cells


def read_row_cells(jab, row_index):
    snapshot = read_body_table(jab, f"row_{row_index}_readback")
    if not snapshot.get("ok"):
        return snapshot, {}
    rows = snapshot.get("rows") or []
    for row in rows:
        if int(row.get("row_index", -1)) == int(row_index):
            return snapshot, (row.get("cells") or {})
    return snapshot, {}


def guarded_add_fee_row_by_ctrl_i(jab, located):
    before = read_body_table(jab, "before_fee_row_add")
    if not before.get("ok"):
        return {
            "ok": False,
            "reason": f"增行前无法读取明细表：{before.get('reason')}",
            "before": before,
        }

    best = located.get("best") or {}
    table_window = best.get("window") or {}
    pressed = guarded_send_ctrl_i(table_window)
    after = read_body_table(jab, "after_fee_row_add")
    before_rows = int(before.get("row_count") or 0)
    after_rows = int(after.get("row_count") or 0)
    ok = (
        bool(pressed.get("ok"))
        and bool(after.get("ok"))
        and after_rows == before_rows + 1
    )
    return {
        "ok": ok,
        "hotkey": ADD_FEE_ROW_HOTKEY,
        "before_rows": before_rows,
        "after_rows": after_rows,
        "before": before,
        "after": after,
        "pressed": pressed,
        "reason": None
        if ok
        else (
            pressed.get("reason")
            or after.get("reason")
            or f"行数未按预期从 {before_rows} 变为 {before_rows + 1}，实际 {after_rows}"
        ),
    }


def click_detail_cell(table_window, table_bounds, row, col, row_count, col_count):
    if sys.platform != "win32":
        return {"ok": False, "reason": "必须在 Windows Python 下运行"}
    if not table_bounds or len(table_bounds) != 4:
        return {"ok": False, "reason": "缺少有效表格 bounds"}
    x, y, width, height = table_bounds
    if x < 0 or y < 0 or width <= 0 or height <= 0:
        return {"ok": False, "reason": f"表格 bounds 不可见：{table_bounds}"}
    if col_count <= 0:
        return {"ok": False, "reason": f"列数无效：{col_count}"}
    table_info = read_window_info((table_window or {}).get("hwnd"))
    foreground_info = read_window_info(ctypes.windll.user32.GetForegroundWindow())
    if not table_info or not foreground_info:
        return {
            "ok": False,
            "reason": "无法读取当前前台窗口或明细表窗口",
            "table_window": table_info,
            "foreground": foreground_info,
        }
    same_root = (
        foreground_info.get("hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("root_hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("hwnd") == table_info.get("hwnd")
    )
    if not same_root:
        return {
            "ok": False,
            "reason": "当前前台窗口不是本次定位到的 NC 收款单窗口，未执行点击",
            "table_window": table_info,
            "foreground": foreground_info,
        }
    target_x, target_y, cell_width, cell_height = cell_center(
        table_bounds, row, col, row_count, col_count
    )
    try:
        move_mouse(target_x, target_y)
        mouse_click()
    except Exception as exc:
        return {
            "ok": False,
            "reason": f"点击明细单元格失败：{type(exc).__name__}: {exc}",
            "target": [target_x, target_y],
            "table_window": table_info,
            "foreground": foreground_info,
        }
    time.sleep(0.45)
    return {
        "ok": True,
        "target": [target_x, target_y],
        "table_bounds": table_bounds,
        "cell_width": cell_width,
        "cell_height": cell_height,
    }


def validate_step_from_cells(step, cells, screen_ok=True, reason=None):
    actual = cells.get(str(step["col"]))
    ok = bool(screen_ok) and field_matches(
        actual, step.get("raw_value") or step["value"], step.get("kind")
    )
    step["ok"] = ok
    step["blocked"] = not ok
    step["actual"] = actual
    if not ok:
        step["reason"] = reason or "表格读回值未匹配目标值；可能该列需要参照/下拉确认"
    else:
        step["reason"] = None


def write_detail_line_by_screen(jab, business, located, fields=None, row_index=0):
    fields = fields or DETAIL_FIELDS
    best = located.get("best") or {}
    table_window = best.get("window") or {}
    table_bounds = best.get("bounds")
    col_count = int(best.get("col_count") or 0)
    row_count = int(best.get("row_count") or 0)
    if not table_bounds:
        return [
            {
                "ok": False,
                "name": "明细表",
                "reason": "缺少明细表 bounds，无法安全计算点击坐标",
            }
        ]
    if row_count <= row_index or col_count < 25:
        return [
            {
                "ok": False,
                "name": "明细表",
                "reason": f"明细表尺寸异常：{row_count} 行 x {col_count} 列，目标第 {row_index + 1} 行",
            }
        ]

    steps = []
    previous_step = None
    for field in fields:
        if is_stop_hotkey_pressed():
            steps.append(
                {
                    "ok": False,
                    "name": field["name"],
                    "value": business[field["value_key"]],
                    "reason": f"检测到紧急停止键 {STOP_HOTKEY}",
                }
            )
            break

        value = str(business[field["value_key"]])
        before_snapshot, before_cells = read_row_cells(jab, row_index)
        try:
            screen = screen_write_amount_cell(
                table_window,
                table_bounds,
                row_index,
                int(field["col"]),
                col_count,
                value,
                "none",
                row_count,
            )
        except Exception as exc:
            screen = {"ok": False, "reason": f"{type(exc).__name__}: {exc}"}
        time.sleep(0.45)
        after_snapshot, after_cells = read_row_cells(jab, row_index)
        if previous_step is not None:
            validate_step_from_cells(previous_step, after_cells)

        step = {
            "step": "detail_cell_screen",
            "ok": None,
            "blocked": None,
            "row": row_index,
            "col": field["col"],
            "name": field["name"],
            "value": normalize_amount_text(value)
            if field.get("kind") == "amount"
            else value,
            "raw_value": value,
            "kind": field.get("kind"),
            "actual": after_cells.get(str(field["col"])),
            "before": before_cells.get(str(field["col"])),
            "target": screen.get("target"),
            "geometry": {
                "table_bounds": screen.get("table_bounds"),
                "row_count": row_count,
                "col_count": col_count,
                "cell_width": screen.get("cell_width"),
                "cell_height": screen.get("cell_height"),
            },
            "reason": None if screen.get("ok") else screen.get("reason"),
            "input_ok": bool(screen.get("ok")),
            "before_table_ok": before_snapshot.get("ok"),
            "after_table_ok": after_snapshot.get("ok"),
        }
        if not screen.get("ok"):
            step["ok"] = False
            step["blocked"] = True
        steps.append(step)
        previous_step = step

    if previous_step is not None and previous_step.get("ok") is None:
        commit_click = click_detail_cell(
            table_window,
            table_bounds,
            row_index,
            int(fields[0]["col"]),
            row_count,
            col_count,
        )
        _snapshot, cells = read_row_cells(jab, row_index)
        validate_step_from_cells(
            previous_step,
            cells,
            screen_ok=bool(commit_click.get("ok")),
            reason=commit_click.get("reason"),
        )
        previous_step["commit_click"] = commit_click
    return steps


def run_fee_only(jab, located, fee_amount):
    add_row = guarded_add_fee_row_by_ctrl_i(jab, located)
    if not add_row.get("ok"):
        return (
            add_row,
            [],
            {"ok": False, "skipped": True, "reason": "增行失败，未清空手续费账户"},
            {"ok": False, "skipped": True, "reason": "增行失败，未删除多余行"},
        )

    # Ctrl+I 后表格行数变化，重新定位一次，拿到最新 bounds/row_count。
    refreshed = locate_receipt_body_table(jab, max_rows=5)
    if not refreshed.get("best"):
        add_row["ok"] = False
        add_row["reason"] = "增行后无法重新定位明细表"
        return (
            add_row,
            [],
            {
                "ok": False,
                "skipped": True,
                "reason": "增行后定位失败，未清空手续费账户",
            },
            {"ok": False, "skipped": True, "reason": "增行后定位失败，未删除多余行"},
        )

    target_row = int(add_row.get("after_rows") or 0) - 1
    steps = write_detail_line_by_screen(
        jab,
        build_fee_business(fee_amount),
        refreshed,
        fields=FEE_FIELDS,
        row_index=target_row,
    )
    refreshed_after_write = locate_receipt_body_table(jab, max_rows=5)
    clear_account = clear_fee_account_if_filled(jab, refreshed_after_write, target_row)
    refreshed_after_clear = locate_receipt_body_table(jab, max_rows=5)
    delete_extra = delete_extra_row_if_present(
        jab, refreshed_after_clear, expected_rows=int(add_row.get("after_rows") or 0)
    )
    return add_row, steps, clear_account, delete_extra


def clear_fee_account_if_filled(jab, located, row_index):
    snapshot, cells = read_row_cells(jab, row_index)
    before = normalize_text(cells.get(str(ACCOUNT_COL)))
    if not snapshot.get("ok"):
        return {
            "ok": False,
            "reason": f"清空前无法读取手续费行：{snapshot.get('reason')}",
        }
    if not before:
        return {"ok": True, "skipped": True, "before": before, "after": before}

    best = located.get("best") or {}
    table_window = best.get("window") or {}
    table_bounds = best.get("bounds")
    row_count = int(best.get("row_count") or 0)
    col_count = int(best.get("col_count") or 0)
    clicked = click_detail_cell(
        table_window, table_bounds, row_index, ACCOUNT_COL, row_count, col_count
    )
    if not clicked.get("ok"):
        return {
            "ok": False,
            "before": before,
            "reason": clicked.get("reason"),
            "clicked": clicked,
        }

    sent = guarded_send_delete(table_window)
    time.sleep(0.5)
    _after_snapshot, after_cells = read_row_cells(jab, row_index)
    after = normalize_text(after_cells.get(str(ACCOUNT_COL)))
    return {
        "ok": bool(sent.get("ok")) and not after,
        "before": before,
        "after": after,
        "clicked": clicked,
        "sent": sent,
        "reason": None
        if bool(sent.get("ok")) and not after
        else sent.get("reason") or "Delete 后账户列仍非空",
    }


def delete_extra_row_if_present(jab, located, expected_rows):
    before = read_body_table(jab, "before_extra_row_delete")
    if not before.get("ok"):
        return {
            "ok": False,
            "reason": f"删行前无法读取明细表：{before.get('reason')}",
        }
    before_rows = int(before.get("row_count") or 0)
    if before_rows <= expected_rows:
        return {
            "ok": True,
            "skipped": True,
            "before_rows": before_rows,
            "after_rows": before_rows,
        }

    best = located.get("best") or {}
    table_window = best.get("window") or {}
    table_bounds = best.get("bounds")
    row_count = int(best.get("row_count") or 0)
    col_count = int(best.get("col_count") or 0)
    target_row = before_rows - 1
    clicked = click_detail_cell(
        table_window, table_bounds, target_row, 1, row_count, col_count
    )
    if not clicked.get("ok"):
        return {
            "ok": False,
            "before_rows": before_rows,
            "reason": clicked.get("reason"),
            "clicked": clicked,
        }

    sent = guarded_send_ctrl_d(table_window)
    time.sleep(0.8)
    after = read_body_table(jab, "after_extra_row_delete")
    after_rows = int(after.get("row_count") or 0)
    return {
        "ok": bool(sent.get("ok")) and after.get("ok") and after_rows == expected_rows,
        "before_rows": before_rows,
        "after_rows": after_rows,
        "clicked": clicked,
        "sent": sent,
        "reason": None
        if bool(sent.get("ok")) and after.get("ok") and after_rows == expected_rows
        else sent.get("reason")
        or after.get("reason")
        or f"Ctrl+D 后行数未回到 {expected_rows}，实际 {after_rows}",
    }


def main():
    args = parse_args()
    config = load_config(str(ROOT / "config.json"))
    account = get_test_account(config, args.bank_label)

    print_header(account, args)
    print()
    print(f"请在 {START_DELAY_SECONDS} 秒内切到 NC 收款单窗口...")
    time.sleep(START_DELAY_SECONDS)
    print("开始测试。")

    report: dict[str, object] = {
        "launcher": "tmp_receipt_detail_main_line_run.py",
        "bank_label": args.bank_label,
        "account": account.account_no,
        "mode": "fee-only" if args.fee_only else "main-line",
        "fee_amount": args.fee_amount if args.fee_only else None,
        "stop_hotkey": STOP_HOTKEY,
        "start_delay_seconds": START_DELAY_SECONDS,
    }
    try:
        if is_stop_hotkey_pressed():
            report.update(
                {
                    "ok": False,
                    "stopped_by_hotkey": True,
                    "failed_step": "before-start",
                }
            )
            print_summary(report)
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
                print_summary(report)
                wait_exit()
                return 1
            report["header_account"] = wait_header_account_description(jab, timeout=2.0)
            located = locate_receipt_body_table(jab, max_rows=3)
            report["table_candidates"] = located.get("candidates", [])[:5]
            if not located.get("best"):
                report.update({"ok": False, "failed_step": "locate-body-table"})
            elif is_stop_hotkey_pressed():
                report.update(
                    {
                        "ok": False,
                        "stopped_by_hotkey": True,
                        "failed_step": "before-fill-detail",
                    }
                )
            else:
                report["before_table"] = read_body_table(jab, "before_detail_fill")
                if args.fee_only:
                    add_row, steps, clear_account, delete_extra = run_fee_only(
                        jab, located, args.fee_amount
                    )
                    report["fee_row_add"] = add_row
                    report["fee_account_clear"] = clear_account
                    report["extra_row_delete"] = delete_extra
                    if not add_row.get("ok"):
                        report["failed_step"] = "add-fee-row"
                    elif not all(bool(step.get("ok")) for step in steps):
                        report["failed_step"] = "fill-fee-line"
                    elif not clear_account.get("ok"):
                        report["failed_step"] = "clear-fee-account"
                    elif not delete_extra.get("ok"):
                        report["failed_step"] = "delete-extra-row"
                else:
                    steps = write_detail_line_by_screen(
                        jab, build_business(account), located
                    )
                report["fill_steps"] = steps
                report["after_table"] = read_body_table(jab, "after_detail_fill")
                report["ok"] = (
                    all(bool(step.get("ok")) for step in steps)
                    and not report.get("failed_step")
                    and bool(steps)
                )
                if not report["ok"]:
                    report["failed_step"] = report.get("failed_step") or (
                        "fill-fee-line" if args.fee_only else "fill-detail-line"
                    )
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
