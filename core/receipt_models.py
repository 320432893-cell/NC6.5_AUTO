# 职责：定义收款单预检、匹配和 NC 抽取使用的不可变数据模型
# 不做什么：不读取/写入 Excel，不解析配置，不匹配 NC，不执行 JAB/GUI 动作
# 允许依赖层：标准库基础类型
# 谁不应该 import：底层 JAB 操作模块不应 import

from dataclasses import dataclass
from datetime import date
from decimal import Decimal


@dataclass(frozen=True)
class ReceiptOrganization:
    code: str
    name: str
    short_name: str


@dataclass(frozen=True)
class ReceiptBank:
    id: str
    name: str
    aliases: tuple[str, ...] = ()
    enabled: bool = True


@dataclass(frozen=True)
class ReceiptAccount:
    organization_code: str
    organization_short_name: str
    account_label: str
    account_no: str
    header_currency_code: str = ""
    id: str = ""
    bank_id: str = ""
    aliases: tuple[str, ...] = ()
    excel_bank_aliases: tuple[str, ...] = ()
    nc_candidates_by_currency: dict[str, tuple[str, ...]] | None = None
    entry_policy: dict[str, object] | None = None
    enabled: bool = True

    def lookup_names(self):
        names = [
            self.account_label,
            *self.aliases,
            *self.excel_bank_aliases,
        ]
        return [name for name in names if str(name or "").strip()]

    def nc_candidates(self, currency=None):
        candidates: list[str] = []
        if self.nc_candidates_by_currency:
            if currency:
                candidates.extend(self.nc_candidates_by_currency.get(currency, ()))
            candidates.extend(self.nc_candidates_by_currency.get("*", ()))
        candidates.append(self.account_no)
        return list(dict.fromkeys(str(item).strip() for item in candidates if item))


@dataclass(frozen=True)
class ReceiptExcelRow:
    row: int
    receipt_date: date
    payer_name: str
    raw_amount: Decimal
    bank: str
    organization_code: str
    organization_name: str
    organization_short_name: str
    nc_done_status: str
    # 手续费:NC 原币金额=raw_amount+fee。默认 0 不破坏既有构造方;
    # dry-run 匹配须按 raw+fee 与 NC 对齐(见 receipt_amounts.receipt_nc_amount)
    fee: Decimal = Decimal("0.00")


@dataclass(frozen=True)
class ReceiptPlanRow:
    row: int
    receipt_date: date
    payer_name: str
    raw_amount: Decimal
    bank: str
    currency: str
    customer_code: str
    fee: Decimal
    organization_code: str
    organization_name: str
    organization_short_name: str
    account_id: str
    account_label: str
    account_no: str
    header_currency_code: str
    duplicate_key: tuple[str, ...]


@dataclass(frozen=True)
class ReceiptPlanIssue:
    excel_row: int | None
    stage: str
    issue_type: str
    field: str
    raw_value: str
    config_node: str
    message: str
    action: str


@dataclass(frozen=True)
class ReceiptNCRow:
    row_index: int
    document_date: date
    customer: str
    original_amount: Decimal


@dataclass(frozen=True)
class ReceiptNCIndexedRow:
    row_index: int
    document_date: date
    original_amount: Decimal
    name: str
    document_no: str
    table_index: int


@dataclass(frozen=True)
class ReceiptNCExtractIssue:
    table_index: int | None
    row_index: int | None
    reason: str


@dataclass(frozen=True)
class ReceiptMatchIssue:
    excel_row: int
    reason: str
    nc_rows: list[int]


@dataclass(frozen=True)
class ReceiptBatchResultRow:
    plan_row: ReceiptPlanRow
    local_status: str
    exception_reason: str = ""
    nc_customer_name: str = ""
    nc_document_no: str = ""
