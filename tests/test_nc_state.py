from decimal import Decimal

from core.nc_state import (
    NCStateDetector,
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


class FakeJAB:
    amount_col = 4
    partner_col = 3

    def read_window_table_cells(self, *args, **kwargs):
        return []

    def window_exists(self, *args, **kwargs):
        return False


class EmptyProbe:
    def collect_named_controls(self, *args, **kwargs):
        return []

    def collect_visible_buttons_by_desc_tokens(self, *args, **kwargs):
        return []

    def read_page_table_signatures(self, *args, **kwargs):
        return []


def test_detect_page_state_fast_fails_when_parent_and_tables_missing():
    detector = NCStateDetector(
        FakeJAB(),
        {},
        "2026-06-17",
        18,
        22,
        9999,
        lambda *args, **kwargs: None,
        lambda *args, **kwargs: None,
    )
    detector.probe = EmptyProbe()

    state = detector.detect_page_state([])

    assert state.name == "error"
    assert "父页面/主表均未检测到" in state.reason
