# 职责：定义收款单明细行字段映射和读回校验口径
# 不做什么：不执行 JAB/GUI 动作，不定位 NC 表格，不读取 Excel
# 允许依赖层：标准库 decimal、core.receipt_keyboard_utils 的金额比较函数
# 谁不应该 import：底层 JAB operator 和 NC 窗口探测模块不应 import

from decimal import Decimal, InvalidOperation

from core.receipt_keyboard_utils import amount_matches


DETAIL_FIELDS = [
    {
        "col": 1,
        "name": "收款业务类型",
        "value_key": "main_business_type",
        "input_mode": "paste",
        "immediate_verify": True,
        "immediate_verify_attempts": 2,
        "immediate_verify_wait": 0.05,
    },
    {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "bank_account",
        "commit_key": "Enter",
        "edit_mode": "selected",
        "input_mode": "paste",
        "pre_commit_wait": 0.1,
        "focus_via_col": 5,
        "pre_write_stabilize": True,
        "pre_write_stabilize_wait": 0.08,
        "immediate_verify": True,
        "immediate_verify_attempts": 3,
        "immediate_verify_wait": 0.2,
    },
    {
        "col": 5,
        "name": "科目",
        "value_key": "main_subject",
        "kind": "code_prefix",
        "input_mode": "paste",
        "sensitive_neighbor_cols": [4, 6, 7, 8],
        "immediate_verify": True,
        "immediate_verify_attempts": 3,
        "immediate_verify_wait": 0.2,
    },
    {
        "col": 7,
        "name": "贷方原币金额",
        "value_key": "amount",
        "kind": "amount",
        "input_mode": "paste",
        "sensitive_neighbor_cols": [6],
        "immediate_verify": True,
        "immediate_verify_attempts": 2,
        "immediate_verify_wait": 0.15,
    },
    {
        "col": 11,
        "name": "结算方式",
        "value_key": "settlement",
        "commit_key": "Enter",
        "input_mode": "paste",
        "pre_write_stabilize": True,
        "pre_write_stabilize_wait": 0.08,
        "immediate_verify": True,
        "immediate_verify_attempts": 2,
        "immediate_verify_wait": 0.05,
    },
]
FEE_FIELDS = [
    {
        "col": 1,
        "name": "收款业务类型",
        "value_key": "fee_business_type",
        "input_mode": "paste",
        "immediate_verify": True,
        "immediate_verify_attempts": 2,
        "immediate_verify_wait": 0.05,
    },
    {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "fee_account",
        "kind": "blank",
        "edit_mode": "selected",
        "focus_via_col": 5,
        "immediate_verify": True,
        "immediate_verify_attempts": 3,
        "immediate_verify_wait": 0.2,
    },
    {
        "col": 5,
        "name": "科目",
        "value_key": "fee_subject",
        "kind": "code_prefix",
        "input_mode": "paste",
        "sensitive_neighbor_cols": [4, 6, 7, 8],
        "immediate_verify": True,
        "immediate_verify_attempts": 3,
        "immediate_verify_wait": 0.2,
    },
    {
        "col": 7,
        "name": "贷方原币金额",
        "value_key": "fee_amount",
        "kind": "amount",
        "input_mode": "paste",
        "sensitive_neighbor_cols": [6],
        "immediate_verify": True,
        "immediate_verify_attempts": 2,
        "immediate_verify_wait": 0.15,
    },
    {
        "col": 11,
        "name": "结算方式",
        "value_key": "settlement",
        "commit_key": "Enter",
        "input_mode": "paste",
        "pre_write_stabilize": True,
        "pre_write_stabilize_wait": 0.08,
        "immediate_verify": True,
        "immediate_verify_attempts": 2,
        "immediate_verify_wait": 0.05,
    },
]
ACCOUNT_COL = 4
BUSINESS_TYPE_COL = 1
SUBJECT_COL = 5
EXCHANGE_RATE_COL = 6
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


def parse_decimal_text(value):
    text = normalize_text(value).replace(",", "").replace(" ", "")
    if not text:
        return None
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return None


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


def normalize_currency_name(value):
    text = normalize_text(value).upper()
    if text in {"USD", "美元"}:
        return "USD"
    if text in {"CNY", "RMB", "人民币"}:
        return "CNY"
    return text


def validate_main_row_exchange_rate(cells, currency, amount, row_index=0):
    actual = normalize_text((cells or {}).get(str(EXCHANGE_RATE_COL)))
    expected_amount = parse_decimal_text(amount)
    normalized_currency = normalize_currency_name(currency)
    if not actual:
        return {
            "ok": False,
            "row": row_index,
            "col": EXCHANGE_RATE_COL,
            "currency": currency,
            "normalized_currency": normalized_currency,
            "actual": actual,
            "normalized_rate": None,
            "amount": normalize_amount_text(amount),
            "normalized_amount": (
                str(expected_amount) if expected_amount is not None else None
            ),
            "reason": "汇率为空或未被读出，保存前汇率校验未通过",
            "policy": "NC 自动带出汇率，自动化只校验不写入；保存前必须读到合法汇率",
        }
    actual_decimal = parse_decimal_text(actual)
    if actual_decimal is None:
        return {
            "ok": False,
            "row": row_index,
            "col": EXCHANGE_RATE_COL,
            "currency": currency,
            "normalized_currency": normalized_currency,
            "actual": actual,
            "normalized_rate": None,
            "amount": normalize_amount_text(amount),
            "normalized_amount": (
                str(expected_amount) if expected_amount is not None else None
            ),
            "reason": f"汇率无法解析为有效数字：{actual!r}",
            "policy": "NC 自动带出汇率，自动化只校验不写入；保存前必须读到合法汇率",
        }
    if expected_amount is not None and actual_decimal == expected_amount:
        return {
            "ok": False,
            "row": row_index,
            "col": EXCHANGE_RATE_COL,
            "currency": currency,
            "normalized_currency": normalized_currency,
            "actual": actual,
            "normalized_rate": str(actual_decimal),
            "amount": normalize_amount_text(amount),
            "normalized_amount": str(expected_amount),
            "reason": "汇率列值等于本次录入金额，疑似金额误写入汇率列",
            "policy": "NC 自动带出汇率，自动化只校验不写入；金额误入汇率列必须失败并取消重开",
        }
    if actual_decimal.copy_abs() >= Decimal("100"):
        return {
            "ok": False,
            "row": row_index,
            "col": EXCHANGE_RATE_COL,
            "currency": currency,
            "normalized_currency": normalized_currency,
            "actual": actual,
            "normalized_rate": str(actual_decimal),
            "amount": normalize_amount_text(amount),
            "normalized_amount": (
                str(expected_amount) if expected_amount is not None else None
            ),
            "reason": f"汇率列值异常偏大：{actual!r}，疑似金额误写入汇率列",
            "policy": "NC 自动带出汇率，自动化只校验不写入；保存前必须读到合法汇率",
        }
    if normalized_currency == "USD":
        ok = Decimal("6") < actual_decimal < Decimal("10")
        reason = (
            None if ok else f"美元汇率不在有效区间：{actual!r}，必须大于 6 且小于 10"
        )
    elif normalized_currency == "CNY":
        ok = actual_decimal == Decimal("1")
        reason = None if ok else f"人民币汇率必须等于 1：{actual!r}"
    else:
        ok = False
        reason = f"不支持的币种汇率校验：{currency!r}"
    return {
        "ok": ok,
        "row": row_index,
        "col": EXCHANGE_RATE_COL,
        "currency": currency,
        "normalized_currency": normalized_currency,
        "actual": actual,
        "normalized_rate": str(actual_decimal),
        "amount": normalize_amount_text(amount),
        "normalized_amount": (
            str(expected_amount) if expected_amount is not None else None
        ),
        "reason": reason,
        "policy": "NC 自动带出汇率，自动化只校验不写入；保存前按币种校验汇率",
    }


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


def field_mismatch_reason(step, actual, prefix="读回值未匹配目标值"):
    return (
        f"{prefix}：字段={step.get('name')}，行={int(step.get('row') or 0) + 1}，"
        f"列={step.get('col')}，期望={step.get('value')!r}，实际={actual!r}"
    )


def apply_readback_to_steps(steps, cells):
    for step in steps:
        actual = cells.get(str(step["col"]))
        step["actual"] = actual
        ok = bool(step.get("input_ok")) and field_matches(
            actual, step.get("raw_value") or step["value"], step.get("kind")
        )
        step["ok"] = ok
        step["blocked"] = not ok
        step["reason"] = (
            None if ok else field_mismatch_reason(step, actual, "整行校验失败")
        )


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
