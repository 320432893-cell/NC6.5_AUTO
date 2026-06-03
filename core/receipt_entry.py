from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re
import unicodedata

import openpyxl

from core.errors import ExcelLockedError, WorkflowStateError


PUNCTUATION_RE = re.compile(r"[^0-9A-Z\u4e00-\u9fff]+")


@dataclass(frozen=True)
class ReceiptOrganization:
    code: str
    name: str
    short_name: str


@dataclass(frozen=True)
class ReceiptAccount:
    organization_code: str
    organization_short_name: str
    account_label: str
    account_no: str


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


class ReceiptEntryConfig:
    def __init__(self, config):
        receipt_cfg = config.get("receipt_entry") or {}
        self.receipt_cfg = receipt_cfg
        self.excel_cfg = receipt_cfg.get("excel") or {}
        self.candidate_cfg = receipt_cfg.get("candidate_check") or {}
        self.organizations = {
            item["code"]: ReceiptOrganization(
                code=item["code"],
                name=item["name"],
                short_name=item["short_name"],
            )
            for item in receipt_cfg.get("finance_organizations", [])
        }
        self.accounts = [
            ReceiptAccount(
                organization_code=item["organization_code"],
                organization_short_name=item["organization_short_name"],
                account_label=item["account_label"],
                account_no=item["account_no"],
            )
            for item in receipt_cfg.get("accounts", [])
        ]
        self.accounts_by_label = {
            normalize_lookup_key(account.account_label): account
            for account in self.accounts
        }

    @property
    def sheet_name(self):
        return self.excel_cfg.get("sheet_name", "💸Payments来款通知")

    @property
    def header_row(self):
        return int(self.excel_cfg.get("header_row", 1))

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
        account = self.accounts_by_label.get(normalize_lookup_key(bank))
        if not account:
            return None
        return self.organizations.get(account.organization_code)


class ReceiptEntryWorkbook:
    def __init__(self, config, excel_path=None):
        self.config = ReceiptEntryConfig(config)
        workbook_path = excel_path or self.config.excel_cfg.get("path")
        if not workbook_path:
            raise WorkflowStateError("receipt_entry.excel.path is required")
        self.excel_path = str(workbook_path)

    def preview_rows(self, today=None):
        wb = openpyxl.load_workbook(self.excel_path, read_only=False, data_only=True)
        try:
            ws = wb[self.config.sheet_name]
            columns = self._read_header(ws)
            rows, issues = self._load_rows(ws, columns)
            return rows, self.select_candidate_rows(rows, today=today), issues
        finally:
            wb.close()

    def ensure_output_columns_and_subjects(self, today=None):
        wb = openpyxl.load_workbook(self.excel_path, read_only=False)
        try:
            ws = wb[self.config.sheet_name]
            columns = self._read_header(ws)
            columns = self._ensure_column(ws, columns, self.config.organization_column)
            columns = self._ensure_column(ws, columns, self.config.nc_done_column)
            rows, issues = self._load_rows(ws, columns)
            org_col = columns[self.config.organization_column]
            for row in rows:
                ws.cell(row=row.row, column=org_col, value=row.organization_name)
            self._save_workbook(wb, "写入收款单主体预处理列")
            return rows, self.select_candidate_rows(rows, today=today), issues
        except PermissionError as exc:
            wb.close()
            raise ExcelLockedError(
                f"Excel 文件无法写入，可能正被 WPS/Excel 打开: path={self.excel_path}"
            ) from exc

    def select_candidate_rows(self, rows, today=None):
        candidate_start = self.config.candidate_start_date(today=today)
        candidates = []
        for row in rows:
            if row.receipt_date < candidate_start:
                continue
            if self.config.candidate_only_blank_status and row.nc_done_status:
                continue
            candidates.append(row)
        return candidates

    def _load_rows(self, ws, columns):
        required = [
            self.config.date_column,
            self.config.payer_name_column,
            self.config.raw_amount_column,
            self.config.bank_column,
        ]
        missing = [name for name in required if name not in columns]
        if missing:
            raise WorkflowStateError(f"收款 Excel 缺少必需列: {missing}")

        rows = []
        issues = []
        for row_index in range(self.config.header_row + 1, ws.max_row + 1):
            raw_date = ws.cell(row_index, columns[self.config.date_column]).value
            if raw_date in (None, ""):
                continue
            try:
                receipt_date = parse_date(raw_date)
            except ValueError as exc:
                issues.append(ReceiptMatchIssue(row_index, str(exc), []))
                continue
            if receipt_date < self.config.start_date:
                continue

            bank = ws.cell(row_index, columns[self.config.bank_column]).value
            payer_name = ws.cell(
                row_index, columns[self.config.payer_name_column]
            ).value
            raw_amount = ws.cell(
                row_index, columns[self.config.raw_amount_column]
            ).value
            nc_done_status = read_optional_cell(
                ws,
                row_index,
                columns.get(self.config.nc_done_column),
            )
            organization = self.config.organization_for_bank(bank)
            if not organization:
                issues.append(ReceiptMatchIssue(row_index, f"银行未配置: {bank!r}", []))
                continue
            try:
                amount = parse_amount(raw_amount)
            except ValueError as exc:
                issues.append(ReceiptMatchIssue(row_index, str(exc), []))
                continue
            payer_text = "" if payer_name is None else str(payer_name).strip()
            if not payer_text:
                issues.append(ReceiptMatchIssue(row_index, "银行来款名为空", []))
                continue
            rows.append(
                ReceiptExcelRow(
                    row=row_index,
                    receipt_date=receipt_date,
                    payer_name=payer_text,
                    raw_amount=amount,
                    bank="" if bank is None else str(bank).strip(),
                    organization_code=organization.code,
                    organization_name=organization.name,
                    organization_short_name=organization.short_name,
                    nc_done_status=nc_done_status,
                )
            )
        return rows, issues

    def _read_header(self, ws):
        columns = {}
        for column in range(1, ws.max_column + 1):
            value = ws.cell(self.config.header_row, column).value
            if value is None:
                continue
            text = str(value).strip()
            if text:
                columns[text] = column
        return columns

    def _ensure_column(self, ws, columns, name):
        if name in columns:
            return columns
        column = ws.max_column + 1
        ws.cell(row=self.config.header_row, column=column, value=name)
        return {**columns, name: column}

    def _save_workbook(self, wb, operation):
        try:
            wb.save(self.excel_path)
        except PermissionError as exc:
            raise ExcelLockedError(
                f"Excel 文件无法写入，可能正被 WPS/Excel 打开: "
                f"operation={operation} path={self.excel_path}"
            ) from exc
        finally:
            wb.close()


class ReceiptEntryMatcher:
    def match(self, excel_rows, nc_rows):
        index = {}
        for nc_row in nc_rows:
            key = nc_row.original_amount
            index.setdefault(key, []).append(nc_row)

        matched = {}
        issues = []
        for excel_row in excel_rows:
            candidates = [
                nc_row
                for nc_row in index.get(excel_row.raw_amount, [])
                if names_match(excel_row.payer_name, nc_row.customer)
            ]
            if len(candidates) == 1:
                matched[excel_row.row] = candidates[0]
            else:
                issues.append(
                    ReceiptMatchIssue(
                        excel_row=excel_row.row,
                        reason="未找到"
                        if not candidates
                        else format_receipt_duplicate_reason(len(candidates)),
                        nc_rows=[row.row_index for row in candidates],
                    )
                )
        return matched, issues


class ReceiptEntryDryRunMatcher:
    def match(self, excel_rows, nc_rows):
        index = {}
        for nc_row in nc_rows:
            key = nc_row.original_amount
            index.setdefault(key, []).append(nc_row)

        matched = {}
        issues = []
        for excel_row in excel_rows:
            candidates = [
                nc_row
                for nc_row in index.get(excel_row.raw_amount, [])
                if names_match(excel_row.payer_name, nc_row.name)
            ]
            if len(candidates) == 1:
                matched[excel_row.row] = candidates[0]
            else:
                issues.append(
                    ReceiptMatchIssue(
                        excel_row=excel_row.row,
                        reason="未找到"
                        if not candidates
                        else format_receipt_duplicate_reason(len(candidates)),
                        nc_rows=[row.row_index for row in candidates],
                    )
                )
        return matched, issues


def format_receipt_duplicate_reason(count):
    return f"重复{count}条：名称和金额相同，需人工确认"


class ReceiptNCResultExtractor:
    def __init__(self, config):
        self.config = ReceiptEntryConfig(config)

    def extract(self, tables):
        return extract_receipt_nc_rows(tables, self.config.result_columns)

    def extract_by_indexes(self, tables, name_column, amount_column=None):
        columns = {**self.config.result_column_indexes, "name": name_column}
        if amount_column is not None:
            columns["original_amount"] = amount_column
        return extract_receipt_nc_rows_by_indexes(tables, columns)


def extract_receipt_nc_rows_by_indexes(tables, columns):
    nc_rows = []
    issues = []
    main_tables = [
        table
        for table in tables
        if table.get("col_count", 0) >= max(columns.values()) + 1
        and table.get("row_count", 0) > 0
    ]
    if not main_tables:
        return [], [
            ReceiptNCExtractIssue(
                table_index=None,
                row_index=None,
                reason=f"未找到可按列位抽取的收款单结果表: columns={columns}",
            )
        ]

    seen_document_numbers = set()
    for table in main_tables:
        for row in table.get("rows") or []:
            row_index = int(row.get("row_index", -1))
            cells = row.get("cells") or []
            if is_blank_result_row(cells, columns):
                continue
            try:
                document_no = read_cell(cells, columns["document_no"])
                if document_no and document_no in seen_document_numbers:
                    continue
                if document_no:
                    seen_document_numbers.add(document_no)
                nc_rows.append(
                    ReceiptNCIndexedRow(
                        row_index=row_index,
                        table_index=int(table.get("table_index", -1)),
                        document_no=document_no,
                        document_date=parse_date(
                            read_cell(cells, columns["document_date"])
                        ),
                        original_amount=parse_amount(
                            read_cell(cells, columns["original_amount"])
                        ),
                        name=read_required_text(cells, columns["name"], "匹配名称为空"),
                    )
                )
            except ValueError as exc:
                issues.append(
                    ReceiptNCExtractIssue(
                        table_index=table.get("table_index"),
                        row_index=row_index,
                        reason=str(exc),
                    )
                )
    return nc_rows, issues


def extract_receipt_nc_rows(tables, result_columns):
    required = {
        "document_date": result_columns.get("document_date", "单据日期"),
        "original_amount": result_columns.get("original_amount", "原币金额"),
        "customer": result_columns.get("customer", "客户"),
    }
    nc_rows = []
    issues = []
    matched_header = False

    for table in tables:
        rows = table.get("rows") or []
        resolved = resolve_receipt_result_columns(rows, required)
        if not resolved:
            continue
        matched_header = True
        header_row_index, columns = resolved
        for row in rows:
            row_index = int(row.get("row_index", -1))
            if row_index <= header_row_index:
                continue
            cells = row.get("cells") or []
            if is_blank_result_row(cells, columns):
                continue
            try:
                nc_rows.append(
                    ReceiptNCRow(
                        row_index=row_index,
                        document_date=parse_date(
                            read_cell(cells, columns["document_date"])
                        ),
                        original_amount=parse_amount(
                            read_cell(cells, columns["original_amount"])
                        ),
                        customer=read_required_text(
                            cells,
                            columns["customer"],
                            "客户为空",
                        ),
                    )
                )
            except ValueError as exc:
                issues.append(
                    ReceiptNCExtractIssue(
                        table_index=table.get("table_index"),
                        row_index=row_index,
                        reason=str(exc),
                    )
                )

    if not matched_header:
        issues.append(
            ReceiptNCExtractIssue(
                table_index=None,
                row_index=None,
                reason=f"未找到包含结果列的收款单表头: {list(required.values())}",
            )
        )
    return nc_rows, issues


def resolve_receipt_result_columns(rows, required):
    required_keys = {
        field: normalize_lookup_key(label) for field, label in required.items()
    }
    for row in rows:
        cells = row.get("cells") or []
        header = {
            normalize_lookup_key(value): column
            for column, value in enumerate(cells)
            if str(value or "").strip()
        }
        if all(label in header for label in required_keys.values()):
            return int(row.get("row_index", -1)), {
                field: header[label] for field, label in required_keys.items()
            }
    return None


def is_blank_result_row(cells, columns):
    return all(not read_cell(cells, column).strip() for column in columns.values())


def read_cell(cells, column):
    if column >= len(cells):
        return ""
    return str(cells[column] or "").strip()


def read_required_text(cells, column, error_message):
    text = read_cell(cells, column)
    if not text:
        raise ValueError(error_message)
    return text


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


def read_optional_cell(ws, row, column):
    if not column:
        return ""
    value = ws.cell(row, column).value
    return "" if value is None else str(value).strip()


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


def normalize_lookup_key(value):
    return "".join(str(value or "").strip().casefold().split())


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
