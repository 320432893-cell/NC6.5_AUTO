from decimal import Decimal

from core.nc_state import (
    choose_main_signature_table,
    is_pending_signature,
    looks_loading,
    normalize_generated_voucher,
)


def test_normalize_generated_voucher():
    assert normalize_generated_voucher("000123", 9999) == 123
    assert normalize_generated_voucher("凭证号 88", 9999) == 88
    assert normalize_generated_voucher("0", 9999) is None
    assert normalize_generated_voucher("10000", 9999) is None
    assert normalize_generated_voucher("未生成", 9999) is None


def test_choose_main_signature_table_prefers_largest_area():
    tables = [
        {"row_count": 2, "col_count": 5},
        {"row_count": 10, "col_count": 3},
        {"row_count": 1, "col_count": 1},
    ]

    assert choose_main_signature_table(tables) == {"row_count": 10, "col_count": 3}


def test_is_pending_signature():
    assert is_pending_signature(
        {"col_count": 25, "voucher_values": []},
        {"查询", "生成", "前台生成"},
    )
    assert not is_pending_signature(
        {"col_count": 23, "voucher_values": ["0001"]},
        {"查询", "生成", "前台生成"},
    )


def test_looks_loading():
    assert looks_loading([], [])
    assert looks_loading([{"name": "单据生成"}], [])
    assert looks_loading([], [{"row_count": 0, "col_count": 25}])
    assert not looks_loading(
        [{"name": "单据生成"}],
        [{"row_count": 3, "col_count": 25, "rows": [{"amount": Decimal("1.00")}]}],
    )
