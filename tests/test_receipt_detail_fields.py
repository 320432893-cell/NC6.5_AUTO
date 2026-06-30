# 生命周期：持久维护
# 覆盖的业务场景：收款单明细字段映射、金额/科目前缀匹配、读回校验失败信息
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_detail_fields.py

from core.receipt_detail_fields import (
    DETAIL_FIELDS,
    apply_readback_to_steps,
    cells_from_steps,
    field_matches,
    make_detail_step,
    validate_main_row_exchange_rate,
)


def test_detail_field_matching_accepts_amount_and_subject_prefix():
    assert field_matches("1,090.00", "1090", "amount")
    assert field_matches("1002\\银行存款", "1002", "code_prefix")
    assert not field_matches("1003\\其他货币资金", "1002", "code_prefix")


def test_apply_readback_to_steps_and_cells_from_steps_keep_column_values():
    business = {
        "main_business_type": "货款",
        "bank_account": "123456",
        "main_subject": "1002",
        "amount": "1090",
        "settlement": "网银",
    }
    steps = [
        make_detail_step(field, business, row_index=0, row_count=1, col_count=25)
        for field in DETAIL_FIELDS
    ]
    for step in steps:
        step["input_ok"] = True

    apply_readback_to_steps(
        steps,
        {
            "1": "货款",
            "4": "123456",
            "5": "1002\\银行存款",
            "7": "1,090.00",
            "11": "网银",
        },
    )

    assert all(step["ok"] for step in steps)
    assert cells_from_steps(steps) == {
        "1": "货款",
        "4": "123456",
        "5": "1002\\银行存款",
        "7": "1,090.00",
        "11": "网银",
    }


def test_exchange_rate_check_accepts_currency_expected_values():
    assert validate_main_row_exchange_rate({"6": "7.12"}, "美元", "1090")["ok"] is True
    assert validate_main_row_exchange_rate({"6": "1"}, "CNY", "1090")["ok"] is True
    assert validate_main_row_exchange_rate({"6": "1.0"}, "RMB", "1090")["ok"] is True
    assert (
        validate_main_row_exchange_rate({"6": "1.0000"}, "人民币", "1090")["ok"] is True
    )


def test_exchange_rate_check_blocks_missing_rate():
    result = validate_main_row_exchange_rate({}, "美元", "1090")

    assert result["ok"] is False
    assert "汇率为空" in result["reason"]


def test_exchange_rate_check_blocks_non_numeric_rate():
    result = validate_main_row_exchange_rate({"6": "abc"}, "美元", "1090")

    assert result["ok"] is False
    assert "无法解析" in result["reason"]


def test_exchange_rate_check_blocks_amount_pollution():
    result = validate_main_row_exchange_rate({"6": "1,090.00"}, "美元", "1090")

    assert result["ok"] is False
    assert "金额误写入汇率列" in result["reason"]


def test_exchange_rate_check_blocks_invalid_usd_rate():
    result = validate_main_row_exchange_rate({"6": "5.9"}, "USD", "1090")

    assert result["ok"] is False
    assert "美元汇率不在有效区间" in result["reason"]


def test_exchange_rate_check_blocks_invalid_cny_rate():
    result = validate_main_row_exchange_rate({"6": "7.12"}, "CNY", "1090")

    assert result["ok"] is False
    assert "人民币汇率必须等于 1" in result["reason"]


def test_exchange_rate_check_blocks_large_value_pollution():
    result = validate_main_row_exchange_rate({"6": "1089.99"}, "USD", "1090")

    assert result["ok"] is False
    assert "异常偏大" in result["reason"]


def test_exchange_rate_check_blocks_unknown_currency():
    result = validate_main_row_exchange_rate({"6": "7.12"}, "EUR", "1090")

    assert result["ok"] is False
    assert "不支持的币种" in result["reason"]
