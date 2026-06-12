# 职责：提供收款单日期、金额、币种、展示文本和重复键的纯解析/格式化函数
# 不做什么：不读取 Excel 文件，不匹配 NC 结果，不写 Sheet2，不执行 JAB/GUI 动作
# 允许依赖层：标准库基础类型和调用方传入的规范化函数
# 谁不应该 import：底层 JAB/NC workflow 模块不应 import

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"日期格式无法识别: {value!r}")


def parse_amount(value):
    if value is None or str(value).strip() == "":
        raise ValueError("原始金额为空")
    text = re.sub(r"\s+", "", str(value).strip().replace(",", ""))
    if re.fullmatch(r"\([+-]?\d+(?:\.\d+)?\)", text):
        inner = text[1:-1].lstrip("+")
        text = inner if inner.startswith("-") else f"-{inner}"
    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"原始金额格式无法识别: {value!r}") from exc


def normalize_receipt_currency(value):
    text = str(value or "").strip().upper()
    if text in {"USD", "美元"}:
        return "美元"
    if text in {"RMB", "CNY", "人民币"}:
        return "人民币"
    return ""


def normalize_lookup_key(value):
    return "".join(str(value or "").strip().casefold().split())


def format_receipt_value(value):
    text = str(value if value is not None else "").strip()
    return text or "空"


def format_receipt_values(values, limit=3):
    items = [format_receipt_value(value) for value in values or []]
    if not items:
        return "空"
    shown = items[:limit]
    if len(items) > limit:
        shown.append(f"...共{len(items)}个")
    return "、".join(shown)


def make_receipt_duplicate_key(
    organization_code,
    receipt_date,
    bank,
    currency,
    customer_code,
    payer_name,
    amount,
    normalize_lookup_key,
    normalize_counterparty,
):
    return (
        str(organization_code or "").strip(),
        receipt_date.isoformat()
        if isinstance(receipt_date, date)
        else str(receipt_date),
        normalize_lookup_key(bank),
        str(currency or "").strip(),
        normalize_lookup_key(customer_code),
        normalize_counterparty(payer_name),
        str(amount),
    )
