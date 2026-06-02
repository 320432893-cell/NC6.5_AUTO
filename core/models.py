from dataclasses import dataclass
from decimal import Decimal
from typing import Any, NotRequired, TypedDict

from core.errors import ContractViolation


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

    def validate_for_processing(self, context: str = "") -> None:
        if self.parse_error:
            return

        errors = []
        if self.row <= 0:
            errors.append("Excel 行号必须为正整数")
        if self.amount is None:
            errors.append("金额不能为空")
        if not self.partner:
            errors.append("对手方不能为空")

        if errors:
            raise ContractViolation(
                "Excel 行数据不符合处理契约: "
                f"context={context or 'unknown'} excel_row={self.row} "
                f"amount={self.amount} partner={self.partner!r} "
                f"source={self.source!r} errors={'; '.join(errors)}"
            )


class TableSnapshotRow(TypedDict, total=False):
    row_index: int
    amount_text: str
    amount: Decimal | None
    partner_text: str
    partner: str
    voucher_text: str
    extra_text: dict[int, str]


@dataclass(frozen=True)
class PendingMatch:
    item: ExcelVoucherItem
    nc_row: int
    row_data: TableSnapshotRow

    def __getitem__(self, key: str) -> Any:
        if not hasattr(self, key):
            raise KeyError(key)
        return getattr(self, key)

    def get(self, key: str, default: Any = None) -> Any:
        return getattr(self, key, default)


@dataclass(frozen=True)
class GeneratedVoucherMatch:
    item: ExcelVoucherItem
    nc_row: int
    row_data: TableSnapshotRow

    def __getitem__(self, key: str) -> Any:
        if not hasattr(self, key):
            raise KeyError(key)
        return getattr(self, key)

    def get(self, key: str, default: Any = None) -> Any:
        return getattr(self, key, default)


@dataclass(frozen=True)
class VoucherPendingMatch:
    item: ExcelVoucherItem
    nc_row: int | None
    row_data: TableSnapshotRow

    def __getitem__(self, key: str) -> Any:
        if not hasattr(self, key):
            raise KeyError(key)
        return getattr(self, key)

    def get(self, key: str, default: Any = None) -> Any:
        return getattr(self, key, default)


@dataclass(frozen=True)
class VoucherSaveMatch:
    item: ExcelVoucherItem
    nc_row: int | None
    row_data: TableSnapshotRow
    table_index: int
    table_rows: int
    voucher_row: int
    voucher_cells: list[str]
    match_mode: str = ""
    fallback_reason: str = ""

    def __getitem__(self, key: str) -> Any:
        if not hasattr(self, key):
            raise KeyError(key)
        return getattr(self, key)

    def get(self, key: str, default: Any = None) -> Any:
        return getattr(self, key, default)

    def validate_for_save(self, context: str = "") -> None:
        errors = []
        if self.table_index < 0:
            errors.append("制单表索引必须非负")
        if self.table_rows <= 0:
            errors.append("制单表行数必须为正")
        if self.voucher_row < 0:
            errors.append("制单行号必须非负")
        if not self.voucher_cells:
            errors.append("制单行单元格不能为空")

        if errors:
            item = self.item
            raise ContractViolation(
                "制单保存匹配不符合契约: "
                f"context={context or 'unknown'} excel_row={item.row} "
                f"amount={item.amount} partner={item.partner!r} "
                f"nc_row={self.nc_row} table_index={self.table_index} "
                f"table_rows={self.table_rows} voucher_row={self.voucher_row} "
                f"errors={'; '.join(errors)}"
            )


@dataclass(frozen=True)
class MatchIssue:
    item: ExcelVoucherItem
    reason: str
    rows: list[int]

    def __getitem__(self, key: str) -> Any:
        if not hasattr(self, key):
            raise KeyError(key)
        return getattr(self, key)

    def get(self, key: str, default: Any = None) -> Any:
        return getattr(self, key, default)


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
