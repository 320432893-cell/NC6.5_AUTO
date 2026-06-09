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
from tools.jab_probe import AccessibleTableCellInfo  # noqa: E402
from tools.tmp_receipt_cell_probe_run import (  # noqa: E402
    amount_matches,
    guarded_send_ctrl_d,
    guarded_send_ctrl_i,
    read_window_info,
    send_hotkey_ctrl_a,
    send_text,
    send_virtual_key,
)

DEFAULT_TEST_BANK_LABEL = "招行"
DEFAULT_TEST_CURRENCY = "人民币"
START_DELAY_SECONDS = 2
ADD_FEE_ROW_HOTKEY = "Ctrl+I"
DETAIL_FIELDS = [
    {"col": 1, "name": "收款业务类型", "value_key": "main_business_type"},
    {"col": 3, "name": "币种", "value_key": "currency"},
    {"col": 4, "name": "收款银行账户", "value_key": "bank_account"},
    {"col": 5, "name": "科目", "value_key": "main_subject", "kind": "code_prefix"},
    {"col": 7, "name": "贷方原币金额", "value_key": "amount", "kind": "amount"},
    # 结算方式放最后：它是后置字段，不再用它提交其他字段。
    {
        "col": 11,
        "name": "结算方式",
        "value_key": "settlement",
        "commit_key": "Enter",
    },
]
ACCOUNT_COL = 4
MAX_FIELD_RETRIES = 3
KEYBOARD_INPUT_COMMIT_KEY = "Right"
VK_KEYS = {
    "F2": 0x71,
    "Right": 0x27,
    "Left": 0x25,
    "Enter": 0x0D,
    "Delete": 0x2E,
}
FEE_FIELDS = [
    {"col": 1, "name": "收款业务类型", "value_key": "fee_business_type"},
    {"col": 4, "name": "收款银行账户", "value_key": "fee_account", "kind": "blank"},
    {"col": 5, "name": "科目", "value_key": "fee_subject", "kind": "code_prefix"},
    {"col": 7, "name": "贷方原币金额", "value_key": "fee_amount", "kind": "amount"},
    # 手续费行也需要结算方式，且结算方式放最后。
    {
        "col": 11,
        "name": "结算方式",
        "value_key": "settlement",
        "commit_key": "Enter",
    },
]


class StepTimer:
    def __init__(self):
        self.items = []

    def measure(self, name, func, *args, **kwargs):
        started_at = time.perf_counter()
        result = func(*args, **kwargs)
        self.add(name, time.perf_counter() - started_at)
        return result

    def add(self, name, seconds):
        self.items.append({"name": name, "seconds": round(float(seconds), 3)})


def get_test_account(config, bank_label):
    receipt_config = ReceiptEntryConfig(config)
    account = receipt_config.account_for_bank(bank_label)
    if account:
        return account
    raise RuntimeError(f"config.json 中找不到银行账户映射：{bank_label}")


def detail_bank_account_no(account, currency=DEFAULT_TEST_CURRENCY):
    candidates = account.nc_candidates(currency)
    return candidates[0] if candidates else account.account_no


def build_business(account):
    return {
        "currency": DEFAULT_TEST_CURRENCY,
        "bank_account": detail_bank_account_no(account),
        "amount": "1090",
        "settlement": "网银",
        "main_subject": "1002",
        "main_business_type": "货款",
    }


def build_fee_business(fee_amount):
    return {
        "fee_business_type": "手续费",
        "fee_account": "",
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
    parser.add_argument(
        "--cleanup-extra-rows-only",
        action="store_true",
        help="只删除明细第 1 行以外的多余行，不写入、不保存。",
    )
    return parser.parse_args()


def print_header(account, args):
    if args.cleanup_extra_rows_only:
        print("测试功能：收款单明细多余行清理")
    elif args.fee_only:
        print("测试功能：收款单手续费行从增行开始完整试写")
    else:
        print("测试功能：收款单明细主行后台填入")
    print()
    print("测试数据来源：")
    print(f"1. 银行标签：{args.bank_label}（来自 config.json 映射）")
    print(f"2. 收款银行账户：{detail_bank_account_no(account)}")
    if args.fee_only:
        print(f"3. 手续费行：手续费 / 660305 / {args.fee_amount} / 网银")
    else:
        print(f"3. 明细主行：货款 / {DEFAULT_TEST_CURRENCY} / 1002 / 1090 / 网银")
    print()
    print("前置条件：")
    print("1. NC 已停在收款单自制录入界面")
    print("2. 当前明细主行已存在；手续费模式要求主行已写好")
    print("3. 当前没有打开参照窗口或提示框")
    print()
    print("本脚本会做：")
    if args.cleanup_extra_rows_only:
        print("1. 定位 25 列明细表，并读取当前行数")
        print("2. 删除第 1 行以外的多余行")
        print("3. 删除后再次读取明细表")
    elif args.fee_only:
        print("1. 定位 25 列明细表，并读取当前行数")
        print(f"2. 前台守卫通过后发送 {ADD_FEE_ROW_HOTKEY} 新增手续费行")
        print("3. 在新增后的最后一行写入：手续费、660305、金额、网银")
        print("4. 写完后统一读回明细表关键列；失败字段最多修复 3 次")
    else:
        print("1. 读取表头【收款银行账户】状态")
        print("2. 定位 25 列明细表")
        print("3. 用 JAB 选中目标单元格后键盘写入明细主行")
        print("4. 写完后统一读回明细表关键列；失败字段最多修复 3 次")
    print()
    print("不会做：保存、暂存、关闭收款单")
    if not args.fee_only:
        print(f"增行规则：只允许手续费非零分支使用 {ADD_FEE_ROW_HOTKEY}；主行不增行")
    print("说明：每个字段都会先用 JAB selection API 选中 row/col，再 F2/Ctrl+A/输入。")
    print(
        f"说明：提交使用方向键 {KEYBOARD_INPUT_COMMIT_KEY}，不使用 Tab，避免误触发增行。"
    )
    print("兼容性：明细字段定位不再依赖列宽、窗口大小或横向滚动坐标。")
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
        attempts = step.get("attempts") or []
        elapsed = sum(float(item.get("seconds") or 0) for item in attempts)
        print(
            f"  {name}: {ok} | 期望={value!r} | 实际={actual!r} | "
            f"尝试={len(attempts)} | 用时={elapsed:.3f}s"
        )
        if step.get("target"):
            print(f"    目标单元格：{step.get('target')}")
        if step.get("geometry"):
            geometry = step.get("geometry") or {}
            print(
                "    表格依据："
                f"bounds={geometry.get('table_bounds')} "
                f"行数={geometry.get('row_count')} 列数={geometry.get('col_count')}"
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
        if report.get("mode") == "cleanup-only":
            print("成功：明细第 1 行以外的多余行已清理；本次没有保存、没有暂存。")
        elif report.get("mode") == "fee-only":
            print("成功：手续费行已增行并写入，通过表格读回校验。")
        else:
            print(
                "成功：明细主行字段已按 JAB 单元格选中+键盘输入写入，并通过表格读回校验。"
            )
    else:
        print("失败：至少一个明细字段没有写入成功；本次没有保存、没有暂存。")
    if report.get("total_seconds") is not None:
        print(f"总用时：{float(report.get('total_seconds') or 0):.3f}s")
    timings = report.get("timings") or []
    if timings:
        print("阶段计时：")
        for item in timings:
            print(f"  {item.get('name')}: {float(item.get('seconds') or 0):.3f}s")


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
    if kind == "blank":
        return normalize_text(actual) == ""
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
    started_at = time.perf_counter()
    before = read_body_table(jab, "before_fee_row_add")
    if not before.get("ok"):
        return {
            "ok": False,
            "reason": f"增行前无法读取明细表：{before.get('reason')}",
            "before": before,
            "seconds": round(time.perf_counter() - started_at, 3),
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
        "seconds": round(time.perf_counter() - started_at, 3),
        "reason": None
        if ok
        else (
            pressed.get("reason")
            or after.get("reason")
            or f"行数未按预期从 {before_rows} 变为 {before_rows + 1}，实际 {after_rows}"
        ),
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


def field_expected_value(field, business):
    value = str(business[field["value_key"]])
    return normalize_amount_text(value) if field.get("kind") == "amount" else value


def make_detail_step(field, business, row_index, row_count, col_count):
    value = str(business[field["value_key"]])
    return {
        "step": "detail_cell_screen",
        "ok": False,
        "blocked": True,
        "row": row_index,
        "col": field["col"],
        "name": field["name"],
        "value": field_expected_value(field, business),
        "raw_value": value,
        "kind": field.get("kind"),
        "actual": None,
        "before": None,
        "attempts": [],
        "input_ok": False,
        "geometry": {
            "table_bounds": None,
            "row_count": row_count,
            "col_count": col_count,
            "cell_width": None,
            "cell_height": None,
        },
    }


def activate_window(hwnd):
    if sys.platform != "win32" or not hwnd:
        return {"ok": False, "reason": "必须在 Windows Python 下运行且需要 hwnd"}
    user32 = ctypes.windll.user32
    hwnd = int(hwnd)
    root = user32.GetAncestor(hwnd, 2) or hwnd
    user32.ShowWindow(root, 9)
    user32.BringWindowToTop(root)
    ok = bool(user32.SetForegroundWindow(root))
    time.sleep(0.05)
    return {
        "ok": ok,
        "hwnd": hwnd,
        "root_hwnd": int(root),
        "foreground": read_window_info(user32.GetForegroundWindow()),
    }


def get_cell_context(jab, vm_id, table_context, row, col):
    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    if not ok or not cell_info.accessibleContext:
        return None
    return cell_info.accessibleContext


def get_cell_info(jab, vm_id, table_context, row, col):
    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    return cell_info if ok else None


def focus_detail_cell(jab, located, row_index, col_index):
    best = located.get("best") or {}
    context, vm_id, owned, window_info = jab.find_context_by_path_once(
        best.get("path"),
        class_name=(best.get("window") or {}).get("class_name"),
        scope_hwnd=(best.get("window") or {}).get("hwnd"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {"ok": False, "reason": "按 path 重新取得明细表 context 失败"}

    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {"ok": False, "reason": "getAccessibleTableInfo 失败"}
        if row_index < 0 or row_index >= table_info.rowCount:
            return {
                "ok": False,
                "reason": f"目标行越界：{row_index} / {table_info.rowCount}",
            }
        if col_index < 0 or col_index >= table_info.columnCount:
            return {
                "ok": False,
                "reason": f"目标列越界：{col_index} / {table_info.columnCount}",
            }
        foreground = foreground_matches_table(window_info)
        activate = (
            {
                "ok": True,
                "skipped": True,
                "reason": "NC 已在前台，未重复激活窗口",
                "foreground": foreground,
            }
            if foreground.get("ok")
            else activate_window(window_info.get("hwnd"))
        )
        child_index = row_index * table_info.columnCount + col_index
        if not jab.has_selection_api():
            return {"ok": False, "reason": "JAB selection API 不可用"}
        jab.dll.clearAccessibleSelectionFromContext(vm_id, context)
        jab.dll.addAccessibleSelectionFromContext(vm_id, context, child_index)
        time.sleep(0.05)
        cell_info = get_cell_info(jab, vm_id, context, row_index, col_index)
        selected = bool(cell_info and cell_info.isSelected)
        cell_context = cell_info.accessibleContext if cell_info else None
        focus_results = []
        if hasattr(jab.dll, "requestFocus"):
            focus_results.append(
                {"target": "table", "ok": bool(jab.dll.requestFocus(vm_id, context))}
            )
            time.sleep(0.02)
            if cell_context:
                focus_results.append(
                    {
                        "target": "cell",
                        "ok": bool(jab.dll.requestFocus(vm_id, cell_context)),
                    }
                )
                time.sleep(0.02)
        return {
            "ok": bool(activate.get("ok")) and selected,
            "activate": activate,
            "child_index": child_index,
            "selected": selected,
            "request_focus": focus_results,
            "window": window_info,
            "target": {"row": row_index, "col": col_index},
            "reason": None
            if bool(activate.get("ok")) and selected
            else activate.get("reason") or "JAB 选中目标单元格后读回不匹配",
        }
    finally:
        jab.release_contexts(vm_id, owned)


def foreground_matches_table(table_window):
    if sys.platform != "win32":
        return {"ok": False, "reason": "必须在 Windows Python 下运行"}
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
    return {
        "ok": bool(same_root),
        "reason": None if same_root else "当前前台窗口不是本次定位到的 NC 收款单窗口",
        "table_window": table_info,
        "foreground": foreground_info,
    }


def guarded_press_virtual_key(table_window, key_name):
    if key_name not in VK_KEYS:
        return {"ok": False, "reason": f"未知按键：{key_name}"}
    guard = foreground_matches_table(table_window)
    if not guard.get("ok"):
        return {**guard, "key": key_name, "ok": False}
    try:
        send_virtual_key(VK_KEYS[key_name], key_up=False)
        time.sleep(0.015)
        send_virtual_key(VK_KEYS[key_name], key_up=True)
    except Exception as exc:
        return {
            **guard,
            "ok": False,
            "key": key_name,
            "mode": f"SendInput({key_name})",
            "reason": f"{type(exc).__name__}: {exc}",
        }
    time.sleep(0.06)
    return {
        **guard,
        "ok": True,
        "key": key_name,
        "mode": f"SendInput({key_name})",
    }


def keyboard_write_selected_cell(
    table_window,
    value,
    commit_key=KEYBOARD_INPUT_COMMIT_KEY,
    clear_only=False,
    accept_key=None,
):
    guard = foreground_matches_table(table_window)
    if not guard.get("ok"):
        return guard
    try:
        send_virtual_key(VK_KEYS["F2"])
        time.sleep(0.025)
        send_hotkey_ctrl_a()
        time.sleep(0.02)
        if clear_only:
            guarded_press_virtual_key(table_window, "Delete")
        else:
            send_text(value)
        time.sleep(0.025)
        accept = None
        if accept_key:
            accept = guarded_press_virtual_key(table_window, accept_key)
            if not accept.get("ok"):
                return {
                    **guard,
                    "ok": False,
                    "mode": "keyboard",
                    "clear_only": clear_only,
                    "accept_key": accept_key,
                    "accept": accept,
                    "commit_key": commit_key,
                    "reason": accept.get("reason"),
                }
        commit = guarded_press_virtual_key(table_window, commit_key)
    except Exception as exc:
        return {
            **guard,
            "ok": False,
            "reason": f"键盘输入失败：{type(exc).__name__}: {exc}",
            "accept_key": accept_key,
            "commit_key": commit_key,
        }
    return {
        **guard,
        "ok": bool(commit.get("ok")),
        "mode": "keyboard",
        "clear_only": clear_only,
        "accept_key": accept_key,
        "accept": accept,
        "commit_key": commit_key,
        "commit": commit,
        "reason": None if commit.get("ok") else commit.get("reason"),
    }


def read_selected_cell(jab, located):
    best = located.get("best") or {}
    context, vm_id, owned, _window_info = jab.find_context_by_path_once(
        best.get("path"),
        class_name=(best.get("window") or {}).get("class_name"),
        scope_hwnd=(best.get("window") or {}).get("hwnd"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {"ok": False, "reason": "按 path 重新取得明细表 context 失败"}
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {"ok": False, "reason": "getAccessibleTableInfo 失败"}
        selected = []
        for row in range(table_info.rowCount):
            for col in range(table_info.columnCount):
                cell_info = AccessibleTableCellInfo()
                ok = jab.dll.getAccessibleTableCellInfo(
                    vm_id,
                    context,
                    row,
                    col,
                    ctypes.byref(cell_info),
                )
                if not ok or not cell_info.isSelected:
                    continue
                text = ""
                if cell_info.accessibleContext:
                    info = jab.get_context_info(vm_id, cell_info.accessibleContext)
                    if info:
                        text = info.name.strip() or info.description.strip()
                selected.append({"row": row, "col": col, "text": text})
        return {
            "ok": True,
            "selected": selected,
            "single": selected[0] if len(selected) == 1 else None,
        }
    finally:
        jab.release_contexts(vm_id, owned)


def write_field_once(
    jab,
    located,
    table_window,
    row_index,
    row_count,
    field,
    next_col,
    business,
    attempt_no,
):
    value = str(business[field["value_key"]])
    attempt_start = time.perf_counter()
    focus = focus_detail_cell(
        jab,
        located,
        row_index,
        int(field["col"]),
    )
    if focus.get("ok"):
        commit_key = field.get("commit_key") or KEYBOARD_INPUT_COMMIT_KEY
        screen = keyboard_write_selected_cell(
            table_window,
            value,
            commit_key=commit_key,
            clear_only=field.get("kind") == "blank",
            accept_key=field.get("accept_key"),
        )
    else:
        screen = {"ok": False, "reason": focus.get("reason"), "focus": focus}
    selected_after = None
    if screen.get("ok"):
        selected_after = {"ok": True, "skipped": True, "reason": "最终统一读回校验"}
    return {
        "attempt": attempt_no,
        "seconds": round(time.perf_counter() - attempt_start, 3),
        "mode": "keyboard",
        "input_ok": bool(screen.get("ok")),
        "input_reason": screen.get("reason"),
        "target": {"row": row_index, "col": int(field["col"])},
        "table_bounds": None,
        "cell_width": None,
        "cell_height": None,
        "focus": focus,
        "commit_ok": bool(screen.get("ok")),
        "commit_key": field.get("commit_key") or KEYBOARD_INPUT_COMMIT_KEY,
        "accept_key": field.get("accept_key"),
        "commit_col": int(next_col),
        "commit_target": selected_after,
        "commit_reason": screen.get("reason"),
        "ok": bool(screen.get("ok")),
    }


def apply_readback_to_steps(steps, cells):
    for step in steps:
        actual = cells.get(str(step["col"]))
        step["actual"] = actual
        ok = bool(step.get("input_ok")) and field_matches(
            actual, step.get("raw_value") or step["value"], step.get("kind")
        )
        step["ok"] = ok
        step["blocked"] = not ok
        step["reason"] = None if ok else "整行校验读回值未匹配目标值"


def refresh_unmatched_settlement_steps(
    jab, steps, row_index, timeout=0.6, interval=0.15
):
    if not any(step.get("name") == "结算方式" and not step.get("ok") for step in steps):
        return None
    deadline = time.perf_counter() + timeout
    last_snapshot = None
    last_cells = None
    while True:
        snapshot, cells = read_row_cells(jab, row_index)
        last_snapshot = snapshot
        last_cells = cells
        if snapshot.get("ok"):
            apply_readback_to_steps(steps, cells)
            if all(step.get("ok") for step in steps if step.get("name") == "结算方式"):
                return {
                    "ok": True,
                    "seconds": round(
                        timeout - max(deadline - time.perf_counter(), 0), 3
                    ),
                    "snapshot": snapshot,
                }
        if time.perf_counter() >= deadline:
            return {
                "ok": False,
                "seconds": timeout,
                "snapshot": last_snapshot,
                "cells": last_cells,
            }
        time.sleep(interval)


def cells_from_steps(steps):
    cells = {}
    for step in steps or []:
        if "actual" not in step:
            continue
        cells[str(step.get("col"))] = step.get("actual")
    return cells


def write_detail_line_by_screen(jab, business, located, fields=None, row_index=0):
    fields = fields or DETAIL_FIELDS
    best = located.get("best") or {}
    table_window = best.get("window") or {}
    table_bounds = best.get("bounds")
    col_count = int(best.get("col_count") or 0)
    row_count = int(best.get("row_count") or 0)
    if row_count <= row_index or col_count < 25:
        return [
            {
                "ok": False,
                "name": "明细表",
                "reason": f"明细表尺寸异常：{row_count} 行 x {col_count} 列，目标第 {row_index + 1} 行",
            }
        ]

    steps = []
    for index, field in enumerate(fields):
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

        step = make_detail_step(field, business, row_index, row_count, col_count)
        next_field = fields[index + 1] if index + 1 < len(fields) else fields[0]
        attempt = write_field_once(
            jab,
            located,
            table_window,
            row_index,
            row_count,
            field,
            next_field["col"],
            business,
            attempt_no=1,
        )
        step["attempts"].append(attempt)
        step["input_ok"] = bool(attempt.get("input_ok"))
        step["target"] = attempt.get("target")
        step["commit_click"] = {
            "ok": attempt.get("commit_ok"),
            "target": attempt.get("commit_target"),
            "reason": attempt.get("commit_reason"),
        }
        step["geometry"].update(
            {
                "table_bounds": attempt.get("table_bounds") or table_bounds,
                "cell_width": attempt.get("cell_width"),
                "cell_height": attempt.get("cell_height"),
            }
        )
        if not attempt.get("ok"):
            step["reason"] = attempt.get("input_reason") or attempt.get("commit_reason")
        steps.append(step)
        if not attempt.get("ok"):
            break
    else:
        _snapshot, cells = read_row_cells(jab, row_index)
        apply_readback_to_steps(steps, cells)
        settle_refresh = refresh_unmatched_settlement_steps(jab, steps, row_index)
        if settle_refresh:
            for step in steps:
                if step.get("name") == "结算方式":
                    step["settlement_stability_check"] = settle_refresh

        for step in steps:
            while (
                not step.get("ok")
                and len(step.get("attempts") or []) < MAX_FIELD_RETRIES
            ):
                field = next(item for item in fields if item["col"] == step["col"])
                field_index = fields.index(field)
                next_field = (
                    fields[field_index + 1]
                    if field_index + 1 < len(fields)
                    else fields[0]
                )
                # 只在修复重试时重新定位，避免正常路径每个字段都扫控件树。
                refreshed = locate_receipt_body_table(jab, max_rows=max(5, row_count))
                attempt = write_field_once(
                    jab,
                    refreshed,
                    table_window,
                    row_index,
                    row_count,
                    field,
                    next_field["col"],
                    business,
                    attempt_no=len(step["attempts"]) + 1,
                )
                step["attempts"].append(attempt)
                step["input_ok"] = bool(attempt.get("input_ok"))
                step["target"] = attempt.get("target")
                step["commit_click"] = {
                    "ok": attempt.get("commit_ok"),
                    "target": attempt.get("commit_target"),
                    "reason": attempt.get("commit_reason"),
                }
                _snapshot, cells = read_row_cells(jab, row_index)
                actual = cells.get(str(step["col"]))
                step["actual"] = actual
                ok = bool(attempt.get("ok")) and field_matches(
                    actual, step.get("raw_value") or step["value"], step.get("kind")
                )
                step["ok"] = ok
                step["blocked"] = not ok
                step["reason"] = (
                    None
                    if ok
                    else (
                        attempt.get("input_reason")
                        or attempt.get("commit_reason")
                        or "修复后读回值仍未匹配目标值"
                    )
                )
            if not step.get("ok"):
                break
    return steps


def run_fee_only(jab, located, fee_amount):
    timings = StepTimer()
    before = timings.measure(
        "fee.read-before-prepare", read_fee_prepare_row_count, jab, located
    )
    before_rows = int(before.get("row_count") or 0)
    if not before.get("ok"):
        add_row = {
            "ok": False,
            "reason": f"手续费准备前无法读取明细表：{before.get('reason')}",
            "before": before,
        }
    elif before_rows > 2:
        cleanup_extra = timings.measure(
            "fee.cleanup-to-second-row",
            delete_extra_row_if_present,
            jab,
            located,
            expected_rows=2,
        )
        if not cleanup_extra.get("ok"):
            cleanup_extra["timings"] = timings.items
            return (
                cleanup_extra,
                [],
                {
                    "ok": False,
                    "skipped": True,
                    "reason": "清理到第 2 行失败，未清空手续费账户",
                },
                cleanup_extra,
            )
        located = timings.measure(
            "fee.locate-after-cleanup", locate_receipt_body_table, jab, max_rows=5
        )
        before_rows = 2
        add_row = {
            "ok": True,
            "skipped": True,
            "reason": "当前超过 2 行，已删到 2 行并覆盖第 2 行为手续费行",
            "hotkey": ADD_FEE_ROW_HOTKEY,
            "before_rows": before.get("row_count"),
            "after_rows": before_rows,
            "before": before,
        }
    elif before_rows == 1:
        add_row = guarded_add_fee_row_by_ctrl_i(jab, located)
        timings.add("fee.add-row", add_row.get("seconds") or 0)
    elif before_rows == 2:
        add_row = {
            "ok": True,
            "skipped": True,
            "reason": "当前已有 2 行，直接覆盖第 2 行为手续费行",
            "hotkey": ADD_FEE_ROW_HOTKEY,
            "before_rows": before_rows,
            "after_rows": before_rows,
            "before": before,
            "after": before,
        }
    else:
        add_row = {
            "ok": False,
            "reason": f"手续费行固定第 2 行，但清理后当前仍有 {before_rows} 行",
            "before_rows": before_rows,
            "after_rows": before_rows,
            "before": before,
        }

    if not add_row.get("ok"):
        add_row["timings"] = timings.items
        return (
            add_row,
            [],
            {"ok": False, "skipped": True, "reason": "增行失败，未清空手续费账户"},
            {"ok": False, "skipped": True, "reason": "增行失败，未删除多余行"},
        )

    refreshed = located_with_row_count(
        located, int(add_row.get("after_rows") or before_rows)
    )

    target_row = 1
    steps = timings.measure(
        "fee.write-line",
        write_detail_line_by_screen,
        jab,
        build_fee_business(fee_amount),
        refreshed,
        fields=FEE_FIELDS,
        row_index=target_row,
    )
    clear_account = timings.measure(
        "fee.clear-account-if-filled",
        clear_fee_account_if_filled,
        jab,
        refreshed,
        target_row,
        known_cells=cells_from_steps(steps),
    )
    delete_extra = timings.measure(
        "fee.delete-extra-after-write",
        delete_extra_row_if_present,
        jab,
        refreshed,
        expected_rows=2,
    )
    delete_extra["timings"] = timings.items
    return add_row, steps, clear_account, delete_extra


def read_fee_prepare_row_count(jab, located):
    fast = read_table_row_count_by_path(jab, located)
    if fast.get("ok"):
        return {
            "ok": True,
            "fast_path": True,
            "row_count": fast.get("row_count"),
            "source": "read_table_row_count_by_path",
        }
    fallback = read_body_table(jab, "before_fee_row_prepare")
    fallback["fast_path"] = False
    fallback["fast_reason"] = fast.get("reason")
    return fallback


def located_with_row_count(located, row_count):
    best = dict((located.get("best") or {}))
    if row_count:
        best["row_count"] = int(row_count)
    return {**located, "best": best}


def cleanup_rows_after_first(jab, located):
    started_at = time.perf_counter()
    before = read_body_table(jab, "before_cleanup_rows_after_first")
    if not before.get("ok"):
        return {
            "ok": False,
            "reason": f"清理多余行前无法读取明细表：{before.get('reason')}",
            "before": before,
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    before_rows = int(before.get("row_count") or 0)
    if before_rows <= 1:
        return {
            "ok": True,
            "skipped": True,
            "reason": "当前只有主行，无需清理多余行",
            "before_rows": before_rows,
            "after_rows": before_rows,
            "steps": [],
            "seconds": round(time.perf_counter() - started_at, 3),
        }

    steps = []
    current_rows = before_rows
    while current_rows > 1:
        step_started_at = time.perf_counter()
        refreshed = locate_receipt_body_table(jab, max_rows=5)
        best = refreshed.get("best") or {}
        table_window = best.get("window") or {}
        target_row = current_rows - 1
        focused = focus_detail_cell(jab, refreshed, target_row, 1)
        if not focused.get("ok"):
            return {
                "ok": False,
                "reason": focused.get("reason"),
                "before_rows": before_rows,
                "after_rows": current_rows,
                "steps": steps,
                "focused": focused,
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        sent = guarded_send_ctrl_d(table_window)
        waited = wait_body_row_count(
            jab,
            expected_rows=current_rows - 1,
            label="after_cleanup_one_extra_row",
        )
        after = waited.get("snapshot") or {}
        after_rows = int(after.get("row_count") or 0)
        step = {
            "target_row": target_row,
            "before_rows": current_rows,
            "after_rows": after_rows,
            "seconds": round(time.perf_counter() - step_started_at, 3),
            "focused": focused,
            "sent": sent,
            "waited": waited,
            "ok": bool(sent.get("ok"))
            and after.get("ok")
            and after_rows == current_rows - 1,
            "reason": None
            if bool(sent.get("ok"))
            and after.get("ok")
            and after_rows == current_rows - 1
            else sent.get("reason")
            or after.get("reason")
            or f"Ctrl+D 后行数未从 {current_rows} 变为 {current_rows - 1}，实际 {after_rows}",
        }
        steps.append(step)
        if not step["ok"]:
            return {
                "ok": False,
                "reason": step["reason"],
                "before_rows": before_rows,
                "after_rows": after_rows,
                "steps": steps,
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        current_rows = after_rows

    return {
        "ok": True,
        "skipped": False,
        "reason": f"已删除第 1 行以外的多余行：{before_rows} -> 1",
        "before_rows": before_rows,
        "after_rows": current_rows,
        "steps": steps,
        "seconds": round(time.perf_counter() - started_at, 3),
    }


def clear_fee_account_if_filled(jab, located, row_index, known_cells=None):
    snapshot = None
    cells = known_cells or {}
    before = normalize_text(cells.get(str(ACCOUNT_COL)))
    if known_cells is None:
        snapshot, cells = read_row_cells(jab, row_index)
        before = normalize_text(cells.get(str(ACCOUNT_COL)))
        if not snapshot.get("ok"):
            return {
                "ok": False,
                "reason": f"清空前无法读取手续费行：{snapshot.get('reason')}",
            }
    if not before:
        return {
            "ok": True,
            "skipped": True,
            "before": before,
            "after": before,
            "source": "known_cells" if known_cells is not None else "read_row_cells",
        }

    best = located.get("best") or {}
    table_window = best.get("window") or {}
    focused = focus_detail_cell(jab, located, row_index, ACCOUNT_COL)
    if not focused.get("ok"):
        return {
            "ok": False,
            "before": before,
            "reason": focused.get("reason"),
            "focused": focused,
        }

    sent = keyboard_write_selected_cell(table_window, "", clear_only=True)
    _after_snapshot, after_cells = read_row_cells(jab, row_index)
    after = normalize_text(after_cells.get(str(ACCOUNT_COL)))
    return {
        "ok": bool(sent.get("ok")) and not after,
        "before": before,
        "after": after,
        "focused": focused,
        "sent": sent,
        "reason": None
        if bool(sent.get("ok")) and not after
        else sent.get("reason") or "Delete 后账户列仍非空",
    }


def delete_extra_row_if_present(jab, located, expected_rows):
    fast = fast_delete_extra_rows_by_row_count(jab, located, expected_rows)
    if fast.get("ok") or fast.get("skipped"):
        return fast

    started_at = time.perf_counter()
    before = read_body_table(jab, "before_extra_row_delete")
    if not before.get("ok"):
        return {
            "ok": False,
            "reason": f"删行前无法读取明细表：{before.get('reason')}",
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    before_rows = int(before.get("row_count") or 0)
    if before_rows <= expected_rows:
        return {
            "ok": True,
            "skipped": True,
            "before_rows": before_rows,
            "after_rows": before_rows,
            "seconds": round(time.perf_counter() - started_at, 3),
        }

    steps = []
    current_rows = before_rows
    while current_rows > expected_rows:
        step_started_at = time.perf_counter()
        refreshed = locate_receipt_body_table(jab, max_rows=max(5, current_rows))
        best = refreshed.get("best") or {}
        table_window = best.get("window") or {}
        target_row = current_rows - 1
        focused = focus_detail_cell(jab, refreshed, target_row, 1)
        if not focused.get("ok"):
            return {
                "ok": False,
                "before_rows": before_rows,
                "after_rows": current_rows,
                "reason": focused.get("reason"),
                "steps": steps,
                "focused": focused,
                "seconds": round(time.perf_counter() - started_at, 3),
            }

        sent = guarded_send_ctrl_d(table_window)
        waited = wait_body_row_count(
            jab,
            expected_rows=current_rows - 1,
            label="after_extra_row_delete",
        )
        after = waited.get("snapshot") or {}
        after_rows = int(after.get("row_count") or 0)
        step = {
            "target_row": target_row,
            "before_rows": current_rows,
            "after_rows": after_rows,
            "seconds": round(time.perf_counter() - step_started_at, 3),
            "focused": focused,
            "sent": sent,
            "waited": waited,
            "ok": bool(sent.get("ok"))
            and after.get("ok")
            and after_rows == current_rows - 1,
            "reason": None
            if bool(sent.get("ok"))
            and after.get("ok")
            and after_rows == current_rows - 1
            else sent.get("reason")
            or after.get("reason")
            or f"Ctrl+D 后行数未从 {current_rows} 变为 {current_rows - 1}，实际 {after_rows}",
        }
        steps.append(step)
        if not step["ok"]:
            return {
                "ok": False,
                "before_rows": before_rows,
                "after_rows": after_rows,
                "reason": step["reason"],
                "steps": steps,
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        current_rows = after_rows

    return {
        "ok": current_rows == expected_rows,
        "before_rows": before_rows,
        "after_rows": current_rows,
        "steps": steps,
        "seconds": round(time.perf_counter() - started_at, 3),
        "reason": None
        if current_rows == expected_rows
        else f"删行后行数未回到 {expected_rows}，实际 {current_rows}",
    }


def fast_delete_extra_rows_by_row_count(jab, located, expected_rows):
    started_at = time.perf_counter()
    before = read_table_row_count_by_path(jab, located)
    if not before.get("ok"):
        return {
            "ok": False,
            "fast_path": True,
            "fallback_required": True,
            "reason": before.get("reason"),
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    before_rows = int(before.get("row_count") or 0)
    if before_rows <= expected_rows:
        return {
            "ok": True,
            "skipped": True,
            "fast_path": True,
            "before_rows": before_rows,
            "after_rows": before_rows,
            "steps": [],
            "seconds": round(time.perf_counter() - started_at, 3),
        }

    best = located.get("best") or {}
    table_window = best.get("window") or {}
    steps = []
    current_rows = before_rows
    while current_rows > expected_rows:
        step_started_at = time.perf_counter()
        target_row = current_rows - 1
        focused = focus_detail_cell(jab, located, target_row, 1)
        if not focused.get("ok"):
            return {
                "ok": False,
                "fast_path": True,
                "fallback_required": True,
                "before_rows": before_rows,
                "after_rows": current_rows,
                "steps": steps,
                "focused": focused,
                "reason": focused.get("reason"),
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        sent = guarded_send_ctrl_d(table_window)
        waited = wait_table_row_count_by_path(
            jab,
            located,
            expected_rows=current_rows - 1,
            label="after_fast_extra_row_delete",
        )
        after_rows = int(waited.get("actual_rows") or 0)
        step = {
            "target_row": target_row,
            "before_rows": current_rows,
            "after_rows": after_rows,
            "seconds": round(time.perf_counter() - step_started_at, 3),
            "focused": focused,
            "sent": sent,
            "waited": waited,
            "ok": bool(sent.get("ok")) and waited.get("ok"),
            "reason": None
            if bool(sent.get("ok")) and waited.get("ok")
            else sent.get("reason") or waited.get("reason"),
        }
        steps.append(step)
        if not step["ok"]:
            return {
                "ok": False,
                "fast_path": True,
                "fallback_required": True,
                "before_rows": before_rows,
                "after_rows": after_rows,
                "steps": steps,
                "reason": step.get("reason"),
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        current_rows = after_rows

    return {
        "ok": True,
        "fast_path": True,
        "before_rows": before_rows,
        "after_rows": current_rows,
        "steps": steps,
        "seconds": round(time.perf_counter() - started_at, 3),
        "reason": None,
    }


def read_table_row_count_by_path(jab, located):
    best = located.get("best") or {}
    context, vm_id, owned, _window_info = jab.find_context_by_path_once(
        best.get("path"),
        class_name=(best.get("window") or {}).get("class_name"),
        scope_hwnd=(best.get("window") or {}).get("hwnd"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {"ok": False, "reason": "按 path 重新取得明细表 context 失败"}
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {"ok": False, "reason": "getAccessibleTableInfo 失败"}
        return {
            "ok": True,
            "row_count": int(table_info.rowCount),
            "col_count": int(table_info.columnCount),
        }
    finally:
        jab.release_contexts(vm_id, owned)


def wait_table_row_count_by_path(
    jab, located, expected_rows, label, timeout=0.8, interval=0.05
):
    started_at = time.perf_counter()
    deadline = time.perf_counter() + timeout
    last = {}
    while True:
        last = read_table_row_count_by_path(jab, located)
        rows = int(last.get("row_count") or 0) if last.get("ok") else 0
        if last.get("ok") and rows == expected_rows:
            return {
                "ok": True,
                "label": label,
                "seconds": round(time.perf_counter() - started_at, 3),
                "expected_rows": expected_rows,
                "actual_rows": rows,
                "snapshot": last,
            }
        if time.perf_counter() >= deadline:
            return {
                "ok": False,
                "label": label,
                "seconds": round(time.perf_counter() - started_at, 3),
                "expected_rows": expected_rows,
                "actual_rows": rows,
                "snapshot": last,
                "reason": f"等待行数变为 {expected_rows} 超时，实际 {rows}",
            }
        time.sleep(interval)


def wait_body_row_count(jab, expected_rows, label, timeout=1.2, interval=0.1):
    started_at = time.perf_counter()
    deadline = time.perf_counter() + timeout
    last = {}
    while True:
        last = read_body_table(jab, label)
        rows = int(last.get("row_count") or 0) if last.get("ok") else 0
        if last.get("ok") and rows == expected_rows:
            return {
                "ok": True,
                "seconds": round(time.perf_counter() - started_at, 3),
                "expected_rows": expected_rows,
                "actual_rows": rows,
                "snapshot": last,
            }
        if time.perf_counter() >= deadline:
            return {
                "ok": False,
                "seconds": round(time.perf_counter() - started_at, 3),
                "expected_rows": expected_rows,
                "actual_rows": rows,
                "snapshot": last,
                "reason": f"等待行数变为 {expected_rows} 超时，实际 {rows}",
            }
        time.sleep(interval)


def main():
    args = parse_args()
    config = load_config(str(ROOT / "config.json"))
    account = get_test_account(config, args.bank_label)

    print_header(account, args)
    print()
    print(f"请在 {START_DELAY_SECONDS} 秒内切到 NC 收款单窗口...")
    wait_started_at = time.perf_counter()
    time.sleep(START_DELAY_SECONDS)
    print("开始测试。")
    run_started_at = time.perf_counter()
    timings = StepTimer()
    timings.add("startup.wait-before-run", time.perf_counter() - wait_started_at)

    report: dict[str, object] = {
        "launcher": "tmp_receipt_detail_main_line_run.py",
        "bank_label": args.bank_label,
        "account": detail_bank_account_no(account),
        "mode": "cleanup-only"
        if args.cleanup_extra_rows_only
        else "fee-only"
        if args.fee_only
        else "main-line",
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
            timings.measure("jab.ensure-started", jab.ensure_started)
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
                print_summary(report)
                wait_exit()
                return 1
            report["header_account"] = timings.measure(
                "header.account-read",
                wait_header_account_description,
                jab,
                timeout=2.0,
            )
            located = timings.measure(
                "body.locate-initial", locate_receipt_body_table, jab, max_rows=3
            )
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
                report["before_table"] = timings.measure(
                    "body.read-before", read_body_table, jab, "before_detail_fill"
                )
                if args.cleanup_extra_rows_only:
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
                elif args.fee_only:
                    add_row, steps, clear_account, delete_extra = timings.measure(
                        "fee.total",
                        run_fee_only,
                        jab,
                        located,
                        args.fee_amount,
                    )
                    for item in (
                        delete_extra.get("timings") or add_row.get("timings") or []
                    ):
                        timings.add(item.get("name"), item.get("seconds") or 0)
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
                    steps = timings.measure(
                        "main.write-line",
                        write_detail_line_by_screen,
                        jab,
                        build_business(account),
                        located,
                    )
                    before_table = report.get("before_table")
                    before_rows = int(
                        before_table.get("row_count") or 0
                        if isinstance(before_table, dict)
                        else 0
                    )
                    refreshed_after_main = timings.measure(
                        "main.locate-after-write",
                        locate_receipt_body_table,
                        jab,
                        max_rows=5,
                    )
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
                if not args.cleanup_extra_rows_only:
                    report["fill_steps"] = steps
                report["after_table"] = timings.measure(
                    "body.read-after", read_body_table, jab, "after_detail_fill"
                )
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

    report["total_seconds"] = round(time.perf_counter() - run_started_at, 3)
    report["timings"] = timings.items
    print()
    print_summary(report)
    print()
    wait_exit()
    return 0 if report.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
