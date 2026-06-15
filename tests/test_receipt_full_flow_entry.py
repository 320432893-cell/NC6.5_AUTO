# 生命周期：持久维护
# 覆盖的业务场景：收款单完整流程入口的计划行选择、业务值映射和保存安全确认
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_entry.py

from datetime import date
from decimal import Decimal

import pytest

from core.receipt_models import ReceiptPlanIssue, ReceiptPlanRow
from tools.receipt_full_flow_entry import (
    business_from_plan_row,
    confirm_save,
    run_one_row,
    select_plan_rows,
)


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


class Args:
    excel_row: int | None = None
    excel_rows: str | None = None
    limit: int = 1


def test_select_plan_rows_skips_issue_rows_and_defaults_limit_one():
    rows = [plan_row(2), plan_row(3), plan_row(4)]
    issues = [
        ReceiptPlanIssue(
            excel_row=2,
            stage="本地数据校验",
            issue_type="CUSTOMER_CODE_EMPTY",
            field="客户编码",
            raw_value="",
            config_node="receipt_entry.excel.customer_code_column",
            message="客户编码为空",
            action="跳过",
        )
    ]

    selected = select_plan_rows(rows, issues, Args())

    assert [row.row for row in selected] == [3]


def test_select_plan_rows_can_target_specific_excel_row():
    args = Args()
    args.excel_row = 4
    args.limit = 10

    selected = select_plan_rows([plan_row(3), plan_row(4)], [], args)

    assert [row.row for row in selected] == [4]


def test_select_plan_rows_can_target_multiple_excel_rows_in_order():
    args = Args()
    args.excel_rows = "4,2,4,3"
    args.limit = 10

    selected = select_plan_rows([plan_row(2), plan_row(3), plan_row(4)], [], args)

    assert [row.row for row in selected] == [4, 2, 3]


def test_business_from_plan_row_maps_receipt_plan_to_entry_values():
    business = business_from_plan_row(plan_row(8, fee=Decimal("20.00")))

    assert business["finance_org_code"] == "A001"
    assert business["document_date"] == "2026-05-22"
    assert business["customer_code"] == "YW03574"
    assert business["header_currency_code"] == "CNY"
    assert business["bank_account"] == "FTE1219165931831"
    assert business["amount"] == "1090.00"
    assert business["fee"] == "20.00"
    assert business["has_fee"] is True
    assert business["settlement"] == "网银"


def test_confirm_save_requires_uppercase_save_without_bypass(monkeypatch):
    class SaveArgs:
        yes_i_understand = False

    monkeypatch.setattr("builtins.input", lambda _prompt: "no")

    with pytest.raises(SystemExit, match="用户取消保存"):
        confirm_save(SaveArgs())


def test_confirm_save_bypass_is_explicit():
    class SaveArgs:
        yes_i_understand = True

    assert confirm_save(SaveArgs()) is None


def test_run_one_row_uses_detail_pipeline_verifier(monkeypatch):
    calls = {"start": 0, "field": [], "snapshot": [], "rows": [], "wait": []}

    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class FakeVerifier:
        def __init__(self, config, located, flow_started_at=None):
            self.config = config
            self.located = located
            self.flow_started_at = flow_started_at

        def start(self):
            calls["start"] += 1

        def submit_field(self, row_index, field, business):
            task_id = f"field-{len(calls['field'])}"
            calls["field"].append(
                (row_index, field["name"], business[field["value_key"]])
            )
            return task_id

        def submit_snapshot(
            self, label, max_rows=5, timeout=1.2, interval=0.08, min_matches=0
        ):
            task_id = f"snapshot-{len(calls['snapshot'])}"
            calls["snapshot"].append((label, max_rows, min_matches))
            return task_id

        def submit_row_count(self, expected_rows, timeout=1.1, interval=0.06):
            task_id = f"rows-{len(calls['rows'])}"
            calls["rows"].append(expected_rows)
            return task_id

        def wait(self, task_ids, timeout=2.0):
            calls["wait"].append((list(task_ids), timeout))
            return {"ok": True, "submitted": list(task_ids), "done": len(task_ids)}

        def close(self, timeout=1.0):
            calls["closed"] = timeout

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config: {"ok": True},
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", FakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.check_jab_ready", lambda _jab: {"ok": True}
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda _jab, _business: [
            {
                "ok": True,
                "label": "客户",
                "value": "YW03574",
                "accepted_text": "ACME LTD",
            }
        ],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda _jab, max_rows=5: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.read_body_table",
        lambda _jab, step: {"ok": True, "step": step, "rows": []},
    )

    def fake_write_detail(_jab, business, _located, after_field=None, **_kwargs):
        field = {"col": 4, "name": "收款银行账户", "value_key": "bank_account"}
        step = {"ok": True, "input_ok": True, "name": field["name"]}
        assert after_field is not None
        step["async_verify_task"] = after_field(0, field, business, step)
        return [step]

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        fake_write_detail,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", FakeVerifier
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_before_save",
        lambda *_args, **_kwargs: {"ok": True, "skipped": True},
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True
    assert calls["start"] == 1
    assert calls["field"] == [(0, "收款银行账户", "FTE1219165931831")]
    assert calls["snapshot"] == [("after-main-line", 3, 1)]
    assert calls["rows"] == [1]
    assert calls["wait"] == [(["field-0", "rows-0"], 2.0)]
    assert report["after_table"]["skipped"] is True
    assert calls["closed"] == 0.2
