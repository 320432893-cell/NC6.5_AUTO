# 职责：提供收款单 Excel 行与 NC 结果行的纯匹配规则和匹配原因文本
# 不做什么：不读取/写入 Excel，不查询 NC，不执行 JAB/GUI 动作，不解析配置
# 允许依赖层：收款单数据模型和解析/格式化纯函数
# 谁不应该 import：底层 JAB 操作模块不应 import

import re
from difflib import SequenceMatcher
import unicodedata

from core.receipt_models import ReceiptMatchIssue
from core.receipt_parsing import format_receipt_value, format_receipt_values


PUNCTUATION_RE = re.compile(r"[^0-9A-Z\u4e00-\u9fff]+")
TOKEN_RE = re.compile(r"[0-9A-Z\u4e00-\u9fff]+")
DEFAULT_NAME_MATCH_THRESHOLD = 80


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
        amount_candidates = index.get(excel_row.raw_amount, [])
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
                    else format_receipt_not_found_reason()
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
        return "金额匹配但名称不匹配，需人工确认"
    return (
        "金额匹配但名称不匹配，需人工确认："
        f"金额={format_receipt_value(excel_amount)}；"
        f"用于匹配的名称/期望名称={format_receipt_value(excel_name)}；"
        f"NC同金额行实际名称={format_receipt_values(nc_names)}"
    )


def format_receipt_name_amount_mismatch_reason(
    excel_amount=None, excel_name=None, nc_amounts=None
):
    if excel_amount is None and excel_name is None and not nc_amounts:
        return "名称匹配但金额不匹配，需人工确认"
    return (
        "名称匹配但金额不匹配，需人工确认："
        f"用于匹配的名称={format_receipt_value(excel_name)}；"
        f"期望金额={format_receipt_value(excel_amount)}；"
        f"NC同名称行实际金额={format_receipt_values(nc_amounts)}"
    )


def format_receipt_not_found_reason():
    return "金额和对手方均未匹配"


def normalize_counterparty(value):
    text = unicodedata.normalize("NFKC", str(value or "")).upper()
    text = re.sub(r"^\s*\d+\s*/\s*", "", text)
    return PUNCTUATION_RE.sub("", text)


def counterparty_tokens(value):
    text = unicodedata.normalize("NFKC", str(value or "")).upper()
    text = re.sub(r"^\s*\d+\s*/\s*", "", text)
    return TOKEN_RE.findall(text)


def counterparty_similarity(left, right):
    left_key = normalize_counterparty(left)
    right_key = normalize_counterparty(right)
    if not left_key or not right_key:
        return 0
    if left_key == right_key:
        return 100
    return round(SequenceMatcher(None, left_key, right_key).ratio() * 100)


def counterparty_match_details(left, right, threshold=DEFAULT_NAME_MATCH_THRESHOLD):
    score = counterparty_similarity(left, right)
    threshold = int(threshold)
    if not left or not right:
        return {
            "ok": False,
            "score": score,
            "threshold": threshold,
            "method": "empty",
            "reason": "名称为空",
        }
    if score >= threshold:
        return {
            "ok": True,
            "score": score,
            "threshold": threshold,
            "method": "similarity",
            "reason": None,
        }
    alias = counterparty_alias_match(left, right)
    if alias.get("ok"):
        return {
            **alias,
            "score": score,
            "threshold": threshold,
        }
    return {
        "ok": False,
        "score": score,
        "threshold": threshold,
        "method": "similarity",
        "reason": f"相似度={score}<阈值{threshold}",
    }


def counterparty_alias_match(left, right):
    left_key = normalize_counterparty(left)
    right_key = normalize_counterparty(right)
    if not left_key or not right_key:
        return {"ok": False, "method": "alias", "reason": "归一化名称为空"}
    if left_key == right_key:
        return {"ok": True, "method": "normalized_equal", "reason": None}
    long_text, short_text = (
        (str(left or ""), str(right or ""))
        if len(left_key) >= len(right_key)
        else (str(right or ""), str(left or ""))
    )
    long_key = normalize_counterparty(long_text)
    short_key = normalize_counterparty(short_text)
    long_tokens = counterparty_tokens(long_text)
    short_tokens = counterparty_tokens(short_text)
    if len(short_key) < 4:
        return {
            "ok": False,
            "method": "alias",
            "reason": "简称过短，不做包含放行",
        }
    if len(short_tokens) >= 2 and token_sequence_count(long_tokens, short_tokens) > 1:
        return {
            "ok": False,
            "method": "alias",
            "reason": "长名称中重复出现短名称，不做包含放行",
        }
    if long_key.startswith(short_key):
        return {
            "ok": True,
            "method": "normalized_prefix",
            "reason": f"归一化长名称以前缀包含短名称：{short_key}",
        }
    if (
        long_tokens
        and short_tokens
        and len(short_tokens) == 1
        and len(short_tokens[0]) >= 4
        and long_tokens[0].startswith(short_tokens[0])
    ):
        return {
            "ok": True,
            "method": "first_token_prefix",
            "reason": f"首词简称匹配：{short_tokens[0]} -> {long_tokens[0]}",
        }
    if len(short_tokens) >= 2 and token_sequence_count(long_tokens, short_tokens) == 1:
        return {
            "ok": True,
            "method": "token_sequence_contains",
            "reason": f"长名称按词包含短名称：{' '.join(short_tokens)}",
        }
    return {
        "ok": False,
        "method": "alias",
        "reason": "未命中安全包含/简称规则",
    }


def token_sequence_count(tokens, needle):
    if not tokens or not needle or len(needle) > len(tokens):
        return 0
    count = 0
    width = len(needle)
    for index in range(0, len(tokens) - width + 1):
        if tokens[index : index + width] == needle:
            count += 1
    return count


def names_match(left, right, threshold=DEFAULT_NAME_MATCH_THRESHOLD):
    return bool(counterparty_match_details(left, right, threshold).get("ok"))
