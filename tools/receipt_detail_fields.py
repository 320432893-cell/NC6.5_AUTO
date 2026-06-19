# 职责：定义收款单明细行字段映射和读回校验口径
# 不做什么：不执行 JAB/GUI 动作，不定位 NC 表格，不读取 Excel
# 允许依赖层：标准库 decimal、tools.receipt_keyboard_utils 的金额比较函数
# 谁不应该 import：底层 JAB operator 和 NC 窗口探测模块不应 import

from decimal import Decimal, InvalidOperation

from tools.receipt_keyboard_utils import amount_matches


DETAIL_FIELDS = [
    {
        "col": 1,
        "name": "收款业务类型",
        "value_key": "main_business_type",
        "input_mode": "paste",
    },
    {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "bank_account",
        "commit_key": "Enter",
        "edit_mode": "selected",
        "input_mode": "paste",
        "pre_commit_wait": 0.1,
    },
    {
        "col": 5,
        "name": "科目",
        "value_key": "main_subject",
        "kind": "code_prefix",
        "input_mode": "paste",
    },
    {
        "col": 7,
        "name": "贷方原币金额",
        "value_key": "amount",
        "kind": "amount",
        "input_mode": "paste",
    },
    {
        "col": 11,
        "name": "结算方式",
        "value_key": "settlement",
        "commit_key": "Enter",
        "input_mode": "paste",
    },
]
FEE_FIELDS = [
    {
        "col": 1,
        "name": "收款业务类型",
        "value_key": "fee_business_type",
        "input_mode": "paste",
    },
    {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "fee_account",
        "kind": "blank",
        "edit_mode": "selected",
    },
    {
        "col": 5,
        "name": "科目",
        "value_key": "fee_subject",
        "kind": "code_prefix",
        "input_mode": "paste",
    },
    {
        "col": 7,
        "name": "贷方原币金额",
        "value_key": "fee_amount",
        "kind": "amount",
        "input_mode": "paste",
    },
    {
        "col": 11,
        "name": "结算方式",
        "value_key": "settlement",
        "commit_key": "Enter",
        "input_mode": "paste",
    },
]
ACCOUNT_COL = 4
SUBJECT_COL = 5
AMOUNT_COL = 7


def normalize_text(value):
    return str(value or "").strip()


def normalize_amount_text(value):
    text = normalize_text(value).replace(",", "")
    if not text:
        return ""
    try:
        return str(Decimal(text).quantize(Decimal("0.01")))
    except (InvalidOperation, ValueError):
        return normalize_text(value)


def field_matches(actual, expected, kind=None):
    if kind == "blank":
        return normalize_text(actual) == ""
    if kind == "amount":
        return amount_matches(actual, expected)
    if kind == "code_prefix":
        actual_text = normalize_text(actual)
        expected_text = normalize_text(expected)
        return actual_text == expected_text or actual_text.startswith(
            f"{expected_text}\\"
        )
    return normalize_text(actual) == normalize_text(expected)


def field_expected_value(field, business):
    value = str(business[field["value_key"]])
    return normalize_amount_text(value) if field.get("kind") == "amount" else value


def make_detail_step(field, business, row_index, row_count, col_count):
    value = str(business[field["value_key"]])
    return {
        "step": "detail_cell_screen",
        "ok": False,
        "blocked": True,
        "row": row_index,
        "col": field["col"],
        "name": field["name"],
        "value": field_expected_value(field, business),
        "raw_value": value,
        "kind": field.get("kind"),
        "actual": None,
        "before": None,
        "attempts": [],
        "input_ok": False,
        "geometry": {
            "table_bounds": None,
            "row_count": row_count,
            "col_count": col_count,
            "cell_width": None,
            "cell_height": None,
        },
    }


def cells_from_steps(steps):
    cells = {}
    for step in steps or []:
        if "actual" not in step:
            continue
        cells[str(step.get("col"))] = step.get("actual")
    return cells


def build_fee_business(fee_amount):
    return {
        "fee_business_type": "手续费",
        "fee_account": "",
        "fee_subject": "660305",
        "fee_amount": str(fee_amount),
        "settlement": "网银",
    }
