# 生命周期：持久维护
# 覆盖的业务场景：收款单完整流程入口测试的共享 plan_row/报告工厂、Fake 替身与 import
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_*.py

# 生命周期：持久维护
# 覆盖的业务场景：收款单完整流程入口的计划行选择、业务值映射和保存安全确认
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_entry.py

from datetime import date
from decimal import Decimal

import pytest

from core.receipt_models import ReceiptPlanIssue, ReceiptPlanRow
from tools.receipt_full_flow_entry import (
    build_console_report_lines,
    business_from_plan_row,
    confirm_save,
    extract_entry_anchor_path,
    extract_entry_dynamic_index,
    extract_entry_scope_hwnd,
    extract_header_accepted_text,
    open_self_made_entry,
    parse_args,
    post_query_failure_reasons,
    read_customer_name_after_header,
    run_one_row,
    save_receipt_by_ctrl_s,
    select_plan_rows,
    wait_receipt_header_anchor_in_current_canvas,
)
from tools.receipt_post_save_query import target_to_match_row


def plan_row(row, fee=Decimal("0.00")):
    return ReceiptPlanRow(
        row=row,
        receipt_date=date(2026, 5, 22),
        payer_name="ACME LTD",
        raw_amount=Decimal("1090.00"),
        bank="招行",
        currency="人民币",
        customer_code="YW03574",
        fee=fee,
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
        account_id="cmb_a001",
        account_label="大陆招行",
        account_no="FTE1219165931831",
        header_currency_code="CNY",
        duplicate_key=("A001", "2026-05-22", "招行"),
    )


def open_report_with_header_anchor(hwnd=2002, dynamic_index=5):
    return {
        "ok": True,
        "entry_state": {
            "hits": [
                {
                    "window": {
                        "hwnd": hwnd,
                        "class_name": "SunAwtCanvas",
                        "visible": True,
                    },
                    "control": {
                        "path": f"0.0.1.0.0.0.0.{dynamic_index}.0.0",
                        "dynamic_index": dynamic_index,
                    },
                }
            ]
        },
    }


class FakeInfo:
    role = "text"
    role_en_US = "text"
    states = "enabled,visible,showing,editable"
    states_en_US = "enabled,visible,showing,editable"

    def __init__(self, name="", description=""):
        self.name = name
        self.description = description


class Args:
    excel_row: int | None = None
    excel_rows: str | None = None
    limit: int = 1


__all__ = [
    'Args',
    'Decimal',
    'FakeInfo',
    'ReceiptPlanIssue',
    'ReceiptPlanRow',
    'build_console_report_lines',
    'business_from_plan_row',
    'confirm_save',
    'date',
    'extract_entry_anchor_path',
    'extract_entry_dynamic_index',
    'extract_entry_scope_hwnd',
    'extract_header_accepted_text',
    'open_report_with_header_anchor',
    'open_self_made_entry',
    'parse_args',
    'plan_row',
    'post_query_failure_reasons',
    'pytest',
    'read_customer_name_after_header',
    'run_one_row',
    'save_receipt_by_ctrl_s',
    'select_plan_rows',
    'target_to_match_row',
    'wait_receipt_header_anchor_in_current_canvas',
]
