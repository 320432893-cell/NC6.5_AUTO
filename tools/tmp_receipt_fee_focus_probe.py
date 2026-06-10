# 生命周期：T0 一次性（删除条件：手续费结算方式后焦点漂移根因确认后删除）
# 覆盖的业务阶段：收款单自制录入-手续费结算方式后焦点观察
# 依赖的服务/环境：Windows Python、NC 收款单录入页、Java Access Bridge、收款单Excel
# 运行方式：python tools/tmp_receipt_fee_focus_probe.py

import ctypes
import os
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ["RECEIPT_SKIP_FEE_EXTRA_ROW_DELETE"] = "1"
os.environ["RECEIPT_SKIP_FEE_ACCOUNT_CLEAR"] = "1"

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_entry import ReceiptEntryWorkbook  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools import tmp_receipt_two_case_save_run as base_run  # noqa: E402
from tools.receipt_self_made_fill_trial import locate_receipt_header_scope  # noqa: E402
from tools.tmp_receipt_detail_main_line_run import read_row_cells  # noqa: E402


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_uint),
        ("flags", ctypes.c_uint),
        ("hwndActive", ctypes.c_void_p),
        ("hwndFocus", ctypes.c_void_p),
        ("hwndCapture", ctypes.c_void_p),
        ("hwndMenuOwner", ctypes.c_void_p),
        ("hwndMoveSize", ctypes.c_void_p),
        ("hwndCaret", ctypes.c_void_p),
        ("rcCaret", ctypes.c_long * 4),
    ]


def choose_a006_fee_case(config):
    workbook = ReceiptEntryWorkbook(config)
    rows, issues, _summary = workbook.build_local_plan(write_sheet=False)
    issue_rows = {issue.excel_row for issue in issues if issue.excel_row is not None}
    candidates = [
        row
        for row in rows
        if row.row not in issue_rows
        and row.organization_code == "A006"
        and row.header_currency_code == "USD"
    ]
    candidates.sort(key=lambda row: (row.receipt_date, row.row), reverse=True)
    if not candidates:
        raise RuntimeError("找不到 A006/USD 有效测试行")
    row = candidates[0]
    return base_run.TestCase(
        name="A006手续费焦点观察",
        excel_row=row.row,
        document_date=row.receipt_date.isoformat(),
        customer_code=row.customer_code,
        payer_name=row.payer_name,
        source_bank=row.bank,
        bank_label=row.account_label,
        bank_account_no=row.account_no,
        currency=row.currency,
        amount=str(row.raw_amount),
        fee="20.00",
    )


def window_class_name(hwnd):
    if not hwnd:
        return ""
    buffer = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


def window_text(hwnd):
    if not hwnd:
        return ""
    length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(length + 1)
    ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def read_gui_focus():
    user32 = ctypes.windll.user32
    foreground = user32.GetForegroundWindow()
    thread_id = user32.GetWindowThreadProcessId(foreground, None)
    info = GUITHREADINFO()
    info.cbSize = ctypes.sizeof(info)
    ok = bool(user32.GetGUIThreadInfo(thread_id, ctypes.byref(info)))
    hwnd_focus = int(info.hwndFocus or 0)
    return {
        "ok": ok,
        "foreground": int(foreground or 0),
        "foreground_class": window_class_name(foreground),
        "foreground_title": window_text(foreground),
        "hwnd_focus": hwnd_focus,
        "focus_class": window_class_name(hwnd_focus),
        "focus_title": window_text(hwnd_focus),
        "hwnd_caret": int(info.hwndCaret or 0),
    }


def describe_hwnd_accessible(jab, hwnd):
    if not hwnd:
        return {"ok": False, "reason": "没有 hwndFocus"}
    vm_id = ctypes.c_long()
    context = ctypes.c_longlong()
    if not jab.dll.getAccessibleContextFromHWND(
        hwnd, ctypes.byref(vm_id), ctypes.byref(context)
    ):
        return {"ok": False, "reason": "getAccessibleContextFromHWND 失败", "hwnd": hwnd}
    try:
        info = jab.get_context_info(vm_id.value, context.value)
        text = jab.get_text_context_value(vm_id.value, context.value)
        if not info:
            return {"ok": False, "reason": "getAccessibleContextInfo 失败", "hwnd": hwnd}
        return {
            "ok": True,
            "hwnd": hwnd,
            "role": info.role_en_US.strip() or info.role.strip(),
            "name": info.name.strip(),
            "description": info.description.strip(),
            "states": info.states_en_US.strip() or info.states.strip(),
            "bounds": [info.x, info.y, info.width, info.height],
            "text": str(text or ""),
        }
    finally:
        jab.release_contexts(vm_id.value, [context.value])


def read_header_settlement(jab):
    scope = locate_receipt_header_scope(jab)
    if not scope.get("ok"):
        return {"ok": False, "reason": scope.get("reason"), "scope": scope}
    return base_run.read_header_field_non_empty(
        jab,
        "结算方式",
        header_cache={},
        scope_hwnd=scope.get("scope_hwnd"),
    )


def print_focus_report(title, jab):
    focus = read_gui_focus()
    accessible = describe_hwnd_accessible(jab, focus.get("hwnd_focus"))
    settlement = read_header_settlement(jab)
    print(title)
    print(f"  GUI焦点: {focus}")
    print(f"  JAB焦点hwnd: {accessible}")
    print(f"  上方结算方式: {settlement}")


def main():
    base_run.SAVE_ENABLED = False
    base_run.TEST_BANK_ACCOUNT_NO = ""
    base_run.ALLOW_EXISTING_ENTRY_FOR_FIRST_CASE = False
    config = load_config(str(ROOT / "config.json"))
    case = choose_a006_fee_case(config)
    print("A006 手续费结算方式后焦点观察：")
    print(
        f"  Sheet1行={case.excel_row} | 日期={case.document_date} | "
        f"客户={case.customer_code} | 银行={case.bank_label} | "
        f"账号={case.bank_account_no} | 金额={case.amount} | 手续费={case.fee}"
    )
    print("  本轮不保存；不清手续费账户；不删空白行。执行后请人工取消当前单据。")
    print(
        f"请在 {base_run.START_DELAY_SECONDS} 秒内切到 NC【收款单录入】且能看到【新增】的页面..."
    )
    time.sleep(base_run.START_DELAY_SECONDS)

    report = base_run.run_one_case(config, case, allow_existing_entry=False)
    base_run.print_case_summary(report)

    jab = JABOperator(config)
    try:
        jab.ensure_started()
        print_focus_report("填写结束后焦点观察：", jab)
        snapshot, row0 = read_row_cells(jab, 0)
        _snapshot2, row1 = read_row_cells(jab, 1)
        print(f"  主行读回: ok={snapshot.get('ok')} cells={row0}")
        print(f"  手续费行读回: cells={row1}")
    finally:
        jab.close()
    return 0 if report.get("ok") or report.get("save_skipped") else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:
        traceback.print_exc()
        raise
