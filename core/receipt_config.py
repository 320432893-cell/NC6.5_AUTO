# 职责：解析收款单配置并提供配置查询对象
# 不做什么：不打开/保存 Excel，不抽取 NC 表格，不执行匹配，不执行 JAB/GUI 动作
# 允许依赖层：收款单数据模型和解析纯函数
# 谁不应该 import：底层 JAB 操作模块不应 import

from datetime import date

from core.receipt_models import ReceiptAccount, ReceiptBank, ReceiptOrganization
from core.receipt_parsing import normalize_lookup_key, parse_date


class ReceiptEntryConfig:
    def __init__(self, config):
        receipt_cfg = config.get("receipt_entry") or {}
        self.receipt_cfg = receipt_cfg
        self.schema_version = int(receipt_cfg.get("schema_version", 1))
        self.excel_cfg = receipt_cfg.get("excel") or {}
        self.candidate_cfg = receipt_cfg.get("candidate_check") or {}
        self.detail_entry_policy = receipt_cfg.get("detail_entry_policy") or {}
        self.organizations = {
            item["code"]: ReceiptOrganization(
                code=item["code"],
                name=item["name"],
                short_name=item["short_name"],
            )
            for item in receipt_cfg.get("finance_organizations", [])
        }
        self.banks = {
            item["id"]: ReceiptBank(
                id=item["id"],
                name=item["name"],
                aliases=tuple(item.get("aliases") or ()),
                enabled=bool(item.get("enabled", True)),
            )
            for item in receipt_cfg.get("banks", [])
        }
        self.accounts = [
            ReceiptAccount(
                organization_code=item["organization_code"],
                organization_short_name=item["organization_short_name"],
                account_label=item["account_label"],
                account_no=item["account_no"],
                header_currency_code=item.get("header_currency_code", ""),
                id=item.get("id") or default_account_id(item),
                bank_id=item.get("bank_id", ""),
                display_name=item.get("display_name", ""),
                aliases=tuple(item.get("aliases") or ()),
                excel_bank_aliases=tuple(item.get("excel_bank_aliases") or ()),
                nc_candidates_by_currency=normalize_candidate_map(
                    item.get("nc_candidates_by_currency")
                ),
                entry_policy=item.get("entry_policy") or {},
                enabled=bool(item.get("enabled", True)),
            )
            for item in receipt_cfg.get("accounts", [])
        ]
        self.accounts_by_label = self._build_account_lookup()

    def _build_account_lookup(self):
        result = {}
        for account in self.accounts:
            if not account.enabled:
                continue
            for name in account.lookup_names():
                key = normalize_lookup_key(name)
                if key:
                    result[key] = account
        return result

    @property
    def sheet_name(self):
        return self.excel_cfg.get("sheet_name", "💸Payments来款通知")

    @property
    def header_row(self):
        return int(self.excel_cfg.get("header_row", 1))

    @property
    def start_row(self):
        return int(self.excel_cfg.get("start_row", self.header_row + 1))

    @property
    def result_sheet_name(self):
        return self.excel_cfg.get("result_sheet_name", "收款单自动化结果")

    @property
    def start_date(self):
        return parse_date(self.excel_cfg.get("start_date", "2026-01-01"))

    @property
    def date_column(self):
        return self.excel_cfg.get("date_column", "到款日期")

    @property
    def payer_name_column(self):
        return self.excel_cfg.get("payer_name_column", "🟪银行来款名")

    @property
    def raw_amount_column(self):
        return self.excel_cfg.get("raw_amount_column", "🟪原始金额")

    @property
    def bank_column(self):
        return self.excel_cfg.get("bank_column", "银行")

    @property
    def currency_column(self):
        return self.excel_cfg.get("currency_column", "币种")

    @property
    def customer_code_column(self):
        return self.excel_cfg.get("customer_code_column", "客户编码")

    @property
    def fee_column(self):
        return self.excel_cfg.get("fee_column", "手续费")

    @property
    def organization_column(self):
        return self.excel_cfg.get("organization_column", "主体名称")

    @property
    def nc_done_column(self):
        return self.excel_cfg.get("nc_done_column", "是否NC已做过")

    @property
    def candidate_recent_months(self):
        return int(self.candidate_cfg.get("recent_months", 2))

    @property
    def candidate_from_date(self):
        value = self.candidate_cfg.get("from_date")
        if value in (None, ""):
            return None
        return parse_date(value)

    @property
    def candidate_only_blank_status(self):
        return bool(self.candidate_cfg.get("only_blank_status", True))

    @property
    def validation_policy(self):
        policy = self.receipt_cfg.get("validation_policy") or {}
        if policy.get("skip_invalid_rows"):
            return "skip_invalid_rows"
        return policy.get("mode", "strict")

    def candidate_start_date(self, today=None):
        explicit = self.candidate_from_date
        if explicit:
            return explicit
        return subtract_months(today or date.today(), self.candidate_recent_months)

    @property
    def result_columns(self):
        return (self.receipt_cfg.get("query") or {}).get("result_columns") or {}

    @property
    def result_column_indexes(self):
        default = {
            "document_no": 0,
            "document_date": 1,
            "customer": 4,
            "original_amount": 6,
            "payer_name": 19,
        }
        configured = (self.receipt_cfg.get("query") or {}).get("result_column_indexes")
        if not configured:
            return default
        return {**default, **configured}

    def organization_for_bank(self, bank):
        account = self.account_for_bank(bank)
        if not account:
            return None
        return self.organizations.get(account.organization_code)

    def account_for_bank(self, bank):
        return self.accounts_by_label.get(normalize_lookup_key(bank))

    def account_lookup_labels(self):
        labels = []
        for account in self.accounts:
            if not account.enabled:
                continue
            labels.extend(account.lookup_names())
        return sorted(dict.fromkeys(labels))


def subtract_months(value, months):
    if months < 0:
        raise ValueError(f"recent_months must be non-negative, got {months!r}")
    month_index = value.month - 1 - months
    year = value.year + month_index // 12
    month = month_index % 12 + 1
    day = min(value.day, days_in_month(year, month))
    return date(year, month, day)


def days_in_month(year, month):
    if month == 12:
        next_month = date(year + 1, 1, 1)
    else:
        next_month = date(year, month + 1, 1)
    return (next_month - date(year, month, 1)).days


def default_account_id(item):
    raw = "_".join(
        str(item.get(key) or "").strip()
        for key in ("organization_code", "account_label", "account_no")
    )
    return normalize_lookup_key(raw) or "account"


def normalize_candidate_map(value):
    if not isinstance(value, dict):
        return None
    result = {}
    for currency, candidates in value.items():
        if isinstance(candidates, str):
            items = [candidates]
        else:
            items = list(candidates or [])
        result[str(currency)] = tuple(
            str(item).strip() for item in items if str(item or "").strip()
        )
    return result
