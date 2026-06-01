from decimal import Decimal

import pytest

from core.data_handler import DataHandler


@pytest.fixture
def handler():
    return DataHandler({"excel_path": "unused.xlsx", "jab_batch": {}})


def test_parse_jab_concat_key(handler):
    amount, partner = handler.parse_jab_concat_key(
        "141,688.50 深圳市科贸电子科技有限公司"
    )

    assert amount == Decimal("141688.50")
    assert partner == "深圳市科贸电子科技有限公司"


def test_parse_jab_concat_key_requires_amount_prefix(handler):
    with pytest.raises(ValueError, match="需要以金额开头"):
        handler.parse_jab_concat_key("深圳市科贸电子科技有限公司141688.50")


def test_parse_split_row(handler):
    amount, partner, source = handler.parse_jab_row(
        raw_key="not a concat key",
        raw_amount="1,000",
        raw_partner=" 深圳 市 ",
    )

    assert amount == Decimal("1000.00")
    assert partner == "深圳市"
    assert source == "split_ab"


def test_select_jab_concat_candidate(handler):
    selected = handler.select_jab_concat_candidate("100A / 200B", 2)

    assert selected == "200B"


def test_looks_like_voucher(handler):
    assert handler._looks_like_voucher(123)
    assert handler._looks_like_voucher(123.0)
    assert handler._looks_like_voucher("00123")
    assert not handler._looks_like_voucher("已生成待回填")
    assert not handler._looks_like_voucher(0)
