from decimal import Decimal

import pytest

from core.errors import ContractViolation
from core.models import ExcelVoucherItem, MatchIssue, VoucherSaveMatch


def make_item(**overrides):
    values = {
        "row": 2,
        "raw_key": "",
        "raw_amount": "",
        "raw_partner": "",
        "amount": Decimal("1.00"),
        "partner": "深圳公司",
        "voucher": "",
        "source": "split_ab",
        "parse_error": "",
    }
    values.update(overrides)
    return ExcelVoucherItem(**values)


def test_excel_voucher_item_contract_rejects_missing_amount_and_partner():
    item = make_item(amount=None, partner="")

    with pytest.raises(
        ContractViolation, match="excel_row=2.*金额不能为空.*对手方不能为空"
    ):
        item.validate_for_processing(context="test")


def test_excel_voucher_item_contract_allows_parse_error_rows():
    item = make_item(amount=None, partner="", parse_error="格式错误")

    item.validate_for_processing(context="test")


def test_voucher_save_match_contract_rejects_invalid_table_fields():
    match = VoucherSaveMatch(
        item=make_item(row=5, amount=Decimal("9.99"), partner="上海公司"),
        nc_row=7,
        row_data={},
        table_index=-1,
        table_rows=0,
        voucher_row=-1,
        voucher_cells=[],
    )

    with pytest.raises(
        ContractViolation,
        match="excel_row=5.*amount=9.99.*partner='上海公司'.*nc_row=7",
    ):
        match.validate_for_save(context="test")


def test_match_issue_identifies_duplicate_match():
    item = make_item()

    assert MatchIssue(item=item, reason="重复2条", rows=[1, 17]).is_duplicate_match()
    assert not MatchIssue(item=item, reason="未找到", rows=[]).is_duplicate_match()
