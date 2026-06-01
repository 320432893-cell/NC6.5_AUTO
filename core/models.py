from decimal import Decimal
from typing import Any, TypedDict


class ExcelVoucherItem(TypedDict):
    row: int
    raw_key: Any
    raw_amount: Any
    raw_partner: Any
    amount: Decimal | None
    partner: str
    voucher: Any
    source: str
    parse_error: str


class TableSnapshotRow(TypedDict, total=False):
    row_index: int
    amount: Decimal | None
    partner: str
    voucher_text: str
    extra_text: dict[int, str]


class PendingMatch(TypedDict):
    item: ExcelVoucherItem
    nc_row: int
    row_data: TableSnapshotRow


class VoucherPendingMatch(TypedDict):
    item: ExcelVoucherItem
    nc_row: int | None
    row_data: TableSnapshotRow


class MatchIssue(TypedDict):
    item: ExcelVoucherItem
    reason: str
    rows: list[int]
