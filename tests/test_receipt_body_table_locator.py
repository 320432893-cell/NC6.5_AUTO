from core.receipt_body_table_locator import KEY_COLUMNS


def test_key_columns_include_exchange_rate_for_guard_snapshot():
    assert KEY_COLUMNS[6] == "汇率"
