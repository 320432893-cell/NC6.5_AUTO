# 生命周期：持久维护
# 覆盖的业务场景：收款单明细字段映射、金额/科目前缀匹配、读回校验失败信息
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_detail_fields.py

from tools.receipt_detail_fields import (
    DETAIL_FIELDS,
    cells_from_steps,
    field_matches,
    make_detail_step,
)


def test_detail_field_matching_accepts_amount_and_subject_prefix():
    assert field_matches("1,090.00", "1090", "amount")
    assert field_matches("1002\\银行存款", "1002", "code_prefix")
    assert not field_matches("1003\\其他货币资金", "1002", "code_prefix")


def test_cells_from_steps_keeps_column_values():
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
    readback = {
        "1": "货款",
        "4": "123456",
        "5": "1002\\银行存款",
        "7": "1,090.00",
        "11": "网银",
    }
    for step in steps:
        step["actual"] = readback.get(str(step["col"]))

    assert cells_from_steps(steps) == readback
