from dataclasses import dataclass
from decimal import Decimal
from typing import Any, NotRequired, TypedDict


@dataclass(frozen=True)
class ExcelVoucherItem:
    row: int
    raw_key: Any
    raw_amount: Any
    raw_partner: Any
    amount: Decimal | None
    partner: str
    voucher: Any
    source: str
    parse_error: str

    def __getitem__(self, key: str) -> Any:
        if not hasattr(self, key):
            raise KeyError(key)
        return getattr(self, key)

    def get(self, key: str, default: Any = None) -> Any:
        return getattr(self, key, default)


class TableSnapshotRow(TypedDict, total=False):
    row_index: int
    amount_text: str
    amount: Decimal | None
    partner_text: str
    partner: str
    voucher_text: str
    extra_text: dict[int, str]


class PendingMatch(TypedDict):
    item: ExcelVoucherItem
    nc_row: int
    row_data: TableSnapshotRow


class GeneratedVoucherMatch(TypedDict):
    item: ExcelVoucherItem
    nc_row: int
    row_data: TableSnapshotRow


class VoucherPendingMatch(TypedDict):
    item: ExcelVoucherItem
    nc_row: int | None
    row_data: TableSnapshotRow


class VoucherSaveMatch(TypedDict):
    item: ExcelVoucherItem
    nc_row: int | None
    row_data: TableSnapshotRow
    table_index: int
    table_rows: int
    voucher_row: int
    voucher_cells: list[str]
    match_mode: NotRequired[str]
    fallback_reason: NotRequired[str]


class MatchIssue(TypedDict):
    item: ExcelVoucherItem
    reason: str
    rows: list[int]


BackfillUpdateValue = int | str
BackfillUpdates = dict[int, BackfillUpdateValue]


class BackfillAuditRecord(TypedDict):
    excel_row: int
    amount: str
    partner: str
    status: str
    update_value: BackfillUpdateValue
    generated_row: NotRequired[int]
    raw_voucher: NotRequired[str]
    issue_reason: NotRequired[str]
    nc_rows: NotRequired[list[int]]
