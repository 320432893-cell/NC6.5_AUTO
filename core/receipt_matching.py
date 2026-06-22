# 职责：提供收款单 Excel 行与 NC 结果行的纯匹配规则和匹配原因文本
# 不做什么：不读取/写入 Excel，不查询 NC，不执行 JAB/GUI 动作，不解析配置
# 允许依赖层：收款单数据模型和解析/格式化纯函数
# 谁不应该 import：底层 JAB 操作模块不应 import

import re
import unicodedata

from core.receipt_amounts import receipt_nc_amount
from core.receipt_models import ReceiptMatchIssue
from core.receipt_parsing import format_receipt_value, format_receipt_values


PUNCTUATION_RE = re.compile(r"[^0-9A-Z\u4e00-\u9fff]+")


class ReceiptEntryMatcher:
    def match(self, excel_rows, nc_rows):
        return _match_receipts(excel_rows, nc_rows, name_attr="customer")


class ReceiptEntryDryRunMatcher:
    def match(self, excel_rows, nc_rows):
        return _match_receipts(excel_rows, nc_rows, name_attr="name")


def _match_receipts(excel_rows, nc_rows, name_attr):
    index = {}
    for nc_row in nc_rows:
        key = nc_row.original_amount
        index.setdefault(key, []).append(nc_row)

    matched = {}
    issues = []
    for excel_row in excel_rows:
        # 按 NC 原币口径(raw+fee)匹配:NC 结果行 original_amount 含手续费,
        # 仅用 raw_amount 会让有手续费的收款系统性假报未匹配
        amount_candidates = index.get(receipt_nc_amount(excel_row), [])
        candidates = [
            nc_row
            for nc_row in amount_candidates
            if names_match(excel_row.payer_name, getattr(nc_row, name_attr))
        ]
        if len(candidates) == 1:
            matched[excel_row.row] = candidates[0]
        else:
            if candidates:
                reason = format_receipt_duplicate_reason(len(candidates))
                issue_rows = [row.row_index for row in candidates]
            elif amount_candidates:
                reason = format_receipt_amount_name_mismatch_reason(
                    excel_amount=excel_row.raw_amount,
                    excel_name=excel_row.payer_name,
                    nc_names=[getattr(row, name_attr) for row in amount_candidates],
                )
                issue_rows = [row.row_index for row in amount_candidates]
            else:
                name_candidates = [
                    nc_row
                    for nc_row in nc_rows
                    if names_match(excel_row.payer_name, getattr(nc_row, name_attr))
                ]
                reason = (
                    format_receipt_name_amount_mismatch_reason(
                        excel_amount=excel_row.raw_amount,
                        excel_name=excel_row.payer_name,
                        nc_amounts=[row.original_amount for row in name_candidates],
                    )
                    if name_candidates
                    else format_receipt_not_found_reason(
                        excel_row=excel_row.row,
                        excel_amount=excel_row.raw_amount,
                        excel_name=excel_row.payer_name,
                    )
                )
                issue_rows = [row.row_index for row in name_candidates]
            issues.append(
                ReceiptMatchIssue(
                    excel_row=excel_row.row,
                    reason=reason,
                    nc_rows=issue_rows,
                )
            )
    return matched, issues


def format_receipt_duplicate_reason(count):
    return f"重复{count}条：名称和金额相同，需人工确认"


def format_receipt_amount_name_mismatch_reason(
    excel_amount=None, excel_name=None, nc_names=None
):
    if excel_amount is None and excel_name is None and not nc_names:
        return "金额匹配但名称不一致，需人工确认"
    return (
        "金额匹配但名称不一致，需人工确认："
        f"Excel金额={format_receipt_value(excel_amount)}；"
        f"Excel对手方={format_receipt_value(excel_name)}；"
        f"NC对手方={format_receipt_values(nc_names)}"
    )


def format_receipt_name_amount_mismatch_reason(
    excel_amount=None, excel_name=None, nc_amounts=None
):
    if excel_amount is None and excel_name is None and not nc_amounts:
        return "名称匹配但金额不一致，需人工确认"
    return (
        "名称匹配但金额不一致，需人工确认："
        f"Excel对手方={format_receipt_value(excel_name)}；"
        f"Excel金额={format_receipt_value(excel_amount)}；"
        f"NC金额={format_receipt_values(nc_amounts)}"
    )


# 诊断口径分类标记：保持稳定前缀，供写回分类按前缀判定“未命中”而非异常。
RECEIPT_NOT_FOUND_MARKER = "金额和对手方均未匹配"


def format_receipt_not_found_reason(
    excel_row=None, excel_amount=None, excel_name=None
):
    if excel_row is None and excel_amount is None and excel_name is None:
        return RECEIPT_NOT_FOUND_MARKER
    return (
        f"{RECEIPT_NOT_FOUND_MARKER}："
        f"Excel第{excel_row}行 金额={format_receipt_value(excel_amount)} "
        f"对手方={format_receipt_value(excel_name)} 在结果表均无匹配；请核对单据。"
    )


def normalize_counterparty(value):
    text = unicodedata.normalize("NFKC", str(value or "")).upper()
    text = re.sub(r"^\s*\d+\s*/\s*", "", text)
    return PUNCTUATION_RE.sub("", text)


def names_match(left, right):
    left_key = normalize_counterparty(left)
    right_key = normalize_counterparty(right)
    if not left_key or not right_key:
        return False
    return left_key in right_key or right_key in left_key
