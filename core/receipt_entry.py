from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re
import unicodedata

import openpyxl

from core.errors import ExcelLockedError, WorkflowStateError


PUNCTUATION_RE = re.compile(r"[^0-9A-Z\u4e00-\u9fff]+")

RESULT_SHEET_HEADERS = [
    "原Sheet1行号",
    "执行主体名称",
    "到款日期",
    "客户编码",
    "币种",
    "银行来款名",
    "实收金额",
    "手续费",
    "总金额",
    "收款银行账户",
    "本地预检状态",
    "异常原因",
]

DEPRECATED_RESULT_SHEET_HEADERS = {
    "执行主体编码",
    "执行主体简称",
    "银行",
    "账户配置ID",
    "异常阶段",
    "异常类型",
    "异常字段",
    "原始值",
    "配置节点",
    "异常说明",
    "处理动作",
    "录入结果",
    "保存结果",
    "后验查询结果",
}


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
    display_name: str = ""
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

    def build_local_plan(self, write_sheet=False):
        wb = openpyxl.load_workbook(self.excel_path, read_only=False, data_only=True)
        try:
            ws = wb[self.config.sheet_name]
            columns = self._read_header(ws)
            rows, issues = self._build_plan_rows(ws, columns)
            duplicate_issues = self._detect_duplicate_rows(rows)
            issues.extend(duplicate_issues)
        finally:
            wb.close()
        if write_sheet:
            writable = openpyxl.load_workbook(self.excel_path, read_only=False)
            try:
                self._write_plan_sheet(writable, rows, issues)
            finally:
                writable.close()
        return rows, issues, self._summarize_plan(rows, issues)

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

    def write_nc_done_statuses(self, row_statuses):
        updates = {
            int(row): str(status)
            for row, status in row_statuses.items()
            if status is not None and str(status).strip()
        }
        if not updates:
            return {"updated": 0, "rows": []}

        wb = openpyxl.load_workbook(self.excel_path, read_only=False)
        try:
            ws = wb[self.config.sheet_name]
            columns = self._read_header(ws)
            columns = self._ensure_column(ws, columns, self.config.nc_done_column)
            status_col = columns[self.config.nc_done_column]
            invalid_rows = [
                row
                for row in updates
                if row <= self.config.header_row or row > ws.max_row
            ]
            if invalid_rows:
                wb.close()
                raise WorkflowStateError(
                    f"收款 Excel 状态写入行号无效: {sorted(invalid_rows)}"
                )
            for row, status in sorted(updates.items()):
                ws.cell(row=row, column=status_col, value=status)
            self._save_workbook(wb, "写入收款单NC状态")
            return {"updated": len(updates), "rows": sorted(updates)}
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

    def _build_plan_rows(self, ws, columns):
        issues = []
        rows = []
        required = [
            self.config.date_column,
            self.config.payer_name_column,
            self.config.raw_amount_column,
            self.config.bank_column,
            self.config.currency_column,
            self.config.customer_code_column,
        ]
        missing = [name for name in required if name not in columns]
        for name in missing:
            issues.append(
                ReceiptPlanIssue(
                    excel_row=None,
                    stage="配置识别",
                    issue_type="EXCEL_REQUIRED_COLUMN_MISSING",
                    field=name,
                    raw_value="",
                    config_node="receipt_entry.excel",
                    message=f"Sheet1 缺少必需列 {name!r}",
                    action="停止整批",
                )
            )
        if missing:
            return [], issues
        if self.config.start_row <= self.config.header_row:
            issues.append(
                ReceiptPlanIssue(
                    excel_row=None,
                    stage="配置识别",
                    issue_type="EXCEL_START_ROW_INVALID",
                    field="start_row",
                    raw_value=str(self.config.start_row),
                    config_node="receipt_entry.excel.start_row",
                    message=(
                        "receipt_entry.excel.start_row 必须大于 "
                        f"header_row={self.config.header_row}"
                    ),
                    action="停止整批",
                )
            )
            return [], issues

        for row_index in range(self.config.start_row, ws.max_row + 1):
            raw_date = ws.cell(row_index, columns[self.config.date_column]).value
            if raw_date in (None, ""):
                continue
            row, row_issues = self._build_plan_row(ws, columns, row_index)
            issues.extend(row_issues)
            if row is not None:
                rows.append(row)
        return rows, issues

    def _build_plan_row(self, ws, columns, row_index):
        row_issues = []
        raw_date = ws.cell(row_index, columns[self.config.date_column]).value
        payer_name = read_optional_cell(
            ws, row_index, columns.get(self.config.payer_name_column)
        )
        raw_amount = ws.cell(row_index, columns[self.config.raw_amount_column]).value
        bank = read_optional_cell(ws, row_index, columns.get(self.config.bank_column))
        currency = read_optional_cell(
            ws, row_index, columns.get(self.config.currency_column)
        )
        customer_code = read_optional_cell(
            ws, row_index, columns.get(self.config.customer_code_column)
        )
        fee_raw = (
            ws.cell(row_index, columns[self.config.fee_column]).value
            if self.config.fee_column in columns
            else None
        )

        receipt_date = None
        amount = None
        fee = Decimal("0.00")
        account = None
        organization = None

        try:
            receipt_date = parse_date(raw_date)
        except ValueError:
            row_issues.append(
                plan_issue(
                    row_index,
                    "本地数据校验",
                    "DATE_INVALID",
                    self.config.date_column,
                    raw_date,
                    "receipt_entry.excel.date_column",
                    f"到款日期格式无法识别: {raw_date!r}",
                )
            )

        if not payer_name:
            row_issues.append(
                plan_issue(
                    row_index,
                    "本地数据校验",
                    "PAYER_NAME_EMPTY",
                    self.config.payer_name_column,
                    payer_name,
                    "receipt_entry.excel.payer_name_column",
                    "银行来款名为空，无法用于后验匹配",
                )
            )

        try:
            amount = parse_amount(raw_amount)
            if amount <= 0:
                row_issues.append(
                    plan_issue(
                        row_index,
                        "本地数据校验",
                        "AMOUNT_ZERO_OR_NEGATIVE",
                        self.config.raw_amount_column,
                        raw_amount,
                        "receipt_entry.excel.raw_amount_column",
                        f"原始金额必须大于 0，当前为 {amount}",
                    )
                )
        except ValueError as exc:
            row_issues.append(
                plan_issue(
                    row_index,
                    "本地数据校验",
                    "AMOUNT_INVALID",
                    self.config.raw_amount_column,
                    raw_amount,
                    "receipt_entry.excel.raw_amount_column",
                    str(exc),
                )
            )

        if not bank:
            row_issues.append(
                plan_issue(
                    row_index,
                    "配置识别",
                    "BANK_EMPTY",
                    self.config.bank_column,
                    bank,
                    "receipt_entry.excel.bank_column",
                    "银行为空，无法匹配 receipt_entry.accounts",
                )
            )
        else:
            account = self.config.account_for_bank(bank)
            if not account:
                row_issues.append(
                    plan_issue(
                        row_index,
                        "配置识别",
                        "BANK_ACCOUNT_NOT_CONFIGURED",
                        self.config.bank_column,
                        bank,
                        (
                            "receipt_entry.accounts[*].account_label/"
                            "aliases/excel_bank_aliases"
                        ),
                        (
                            f"Sheet1 银行={bank!r} 未匹配任何账户配置；"
                            f"可用配置值={self.config.account_lookup_labels()}"
                        ),
                    )
                )
            elif not account.enabled:
                row_issues.append(
                    plan_issue(
                        row_index,
                        "配置识别",
                        "BANK_ACCOUNT_DISABLED",
                        self.config.bank_column,
                        bank,
                        f"receipt_entry.accounts[{account.id}].enabled",
                        f"银行={bank!r} 匹配到账户 {account.id!r}，但账户已禁用",
                    )
                )
            else:
                organization = self.config.organizations.get(account.organization_code)
                if not organization:
                    row_issues.append(
                        plan_issue(
                            row_index,
                            "配置识别",
                            "ORG_NOT_CONFIGURED",
                            "organization_code",
                            account.organization_code,
                            "receipt_entry.finance_organizations",
                            (
                                f"账户 {account.id!r} 的 organization_code="
                                f"{account.organization_code!r} 不存在"
                            ),
                        )
                    )

        currency_name = normalize_receipt_currency(currency)
        if not currency:
            row_issues.append(
                plan_issue(
                    row_index,
                    "本地数据校验",
                    "CURRENCY_EMPTY",
                    self.config.currency_column,
                    currency,
                    "receipt_entry.excel.currency_column",
                    "币种为空，无法选择 NC 明细币种和账号候选",
                )
            )
        elif not currency_name:
            row_issues.append(
                plan_issue(
                    row_index,
                    "本地数据校验",
                    "CURRENCY_UNSUPPORTED",
                    self.config.currency_column,
                    currency,
                    "receipt_entry.excel.currency_column",
                    f"币种={currency!r} 不在支持列表 USD/RMB/CNY/美元/人民币",
                )
            )

        if not customer_code:
            row_issues.append(
                plan_issue(
                    row_index,
                    "本地数据校验",
                    "CUSTOMER_CODE_EMPTY",
                    self.config.customer_code_column,
                    customer_code,
                    "receipt_entry.excel.customer_code_column",
                    "客户编码为空，不能写收款单表头客户字段",
                )
            )

        if fee_raw not in (None, ""):
            try:
                fee = parse_amount(fee_raw)
                if fee < 0:
                    row_issues.append(
                        plan_issue(
                            row_index,
                            "本地数据校验",
                            "FEE_NEGATIVE",
                            self.config.fee_column,
                            fee_raw,
                            "receipt_entry.excel.fee_column",
                            f"手续费不能小于 0，当前为 {fee}",
                        )
                    )
            except ValueError as exc:
                row_issues.append(
                    plan_issue(
                        row_index,
                        "本地数据校验",
                        "FEE_INVALID",
                        self.config.fee_column,
                        fee_raw,
                        "receipt_entry.excel.fee_column",
                        str(exc).replace("原始金额", "手续费"),
                    )
                )

        if account and currency_name and not account.nc_candidates(currency_name):
            row_issues.append(
                plan_issue(
                    row_index,
                    "配置识别",
                    "DETAIL_ACCOUNT_CANDIDATE_MISSING",
                    self.config.bank_column,
                    bank,
                    f"receipt_entry.accounts[{account.id}].nc_candidates_by_currency",
                    (
                        f"账户 {account.id!r} 在币种 {currency_name!r} 下没有可用 "
                        "NC 账号候选"
                    ),
                )
            )

        if row_issues:
            return None, row_issues
        if (
            receipt_date is None
            or amount is None
            or account is None
            or organization is None
            or currency_name is None
        ):
            raise WorkflowStateError("收款单本地预检内部状态不完整，无法生成运行计划")
        duplicate_key = make_receipt_duplicate_key(
            organization.code,
            receipt_date,
            bank,
            currency_name,
            customer_code,
            payer_name,
            amount,
        )
        return (
            ReceiptPlanRow(
                row=row_index,
                receipt_date=receipt_date,
                payer_name=payer_name,
                raw_amount=amount,
                bank=bank,
                currency=currency_name,
                customer_code=customer_code,
                fee=fee,
                organization_code=organization.code,
                organization_name=organization.name,
                organization_short_name=organization.short_name,
                account_id=account.id,
                account_label=account.account_label,
                account_no=account.account_no,
                header_currency_code=account.header_currency_code,
                duplicate_key=duplicate_key,
            ),
            [],
        )

    def _detect_duplicate_rows(self, rows):
        grouped = {}
        for row in rows:
            grouped.setdefault(row.duplicate_key, []).append(row)
        issues = []
        for key, group in grouped.items():
            if len(group) <= 1:
                continue
            row_numbers = [row.row for row in group]
            key_text = " + ".join(key)
            for row in group:
                issues.append(
                    ReceiptPlanIssue(
                        excel_row=row.row,
                        stage="本地重复校验",
                        issue_type="DUPLICATE_EXCEL_ROWS",
                        field="重复键",
                        raw_value=key_text,
                        config_node="local.duplicate_key",
                        message=(
                            f"本批 Sheet1 存在重复行；重复键={key_text}；"
                            f"重复原行号={row_numbers}；为避免重复制单，整组未录入。"
                        ),
                        action="跳过重复组",
                    )
                )
        return issues

    def _summarize_plan(self, rows, issues):
        duplicate_rows = {
            issue.excel_row
            for issue in issues
            if issue.issue_type == "DUPLICATE_EXCEL_ROWS" and issue.excel_row
        }
        runnable = [row for row in rows if row.row not in duplicate_rows]
        grouped = {}
        for row in runnable:
            grouped.setdefault(row.organization_code, []).append(row.row)
        return {
            "rows": len(rows),
            "issues": len(issues),
            "runnable_rows": len(runnable),
            "duplicate_rows": sorted(duplicate_rows),
            "organizations": {key: value for key, value in sorted(grouped.items())},
            "validation_policy": self.config.validation_policy,
            "can_run": not issues
            or self.config.validation_policy == "skip_invalid_rows",
        }

    def _write_plan_sheet(self, wb, rows, issues):
        name = self.config.result_sheet_name
        if name not in wb.sheetnames:
            ws = wb.create_sheet(name)
        else:
            ws = wb[name]
        columns = ensure_result_sheet_headers(ws, self.config.header_row)
        append_start_row = ws.max_row + 1
        if append_start_row <= self.config.header_row:
            append_start_row = self.config.header_row + 1
        issues_by_row = {}
        global_issues = []
        for issue in issues:
            if issue.excel_row is None:
                global_issues.append(issue)
            else:
                issues_by_row.setdefault(issue.excel_row, []).append(issue)
        rows_by_number = {row.row: row for row in rows}
        emitted_rows = set()
        for row in sorted(rows, key=plan_sheet_sort_key):
            emitted_rows.add(row.row)
            row_issues = issues_by_row.get(row.row, [])
            if row_issues:
                for issue in row_issues:
                    append_start_row = append_plan_sheet_row(
                        ws,
                        columns,
                        append_start_row,
                        plan_sheet_row(row, issue, "异常"),
                    )
            else:
                append_start_row = append_plan_sheet_row(
                    ws,
                    columns,
                    append_start_row,
                    plan_sheet_row(row, None, "通过"),
                )
        orphan_issue_items = [
            (row_number, row_issues)
            for row_number, row_issues in issues_by_row.items()
            if row_number not in emitted_rows
        ]
        orphan_issue_items.sort(key=lambda item: orphan_issue_sort_key(item, rows_by_number))
        for row_number, row_issues in orphan_issue_items:
            if row_number in emitted_rows:
                continue
            for issue in row_issues:
                append_start_row = append_plan_sheet_row(
                    ws,
                    columns,
                    append_start_row,
                    plan_sheet_row(rows_by_number.get(row_number), issue, "异常"),
                )
        for issue in global_issues:
            append_start_row = append_plan_sheet_row(
                ws,
                columns,
                append_start_row,
                plan_sheet_row(None, issue, "异常"),
            )
        wb.save(self.excel_path)

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
            amount_candidates = index.get(excel_row.raw_amount, [])
            candidates = [
                nc_row
                for nc_row in amount_candidates
                if names_match(excel_row.payer_name, nc_row.customer)
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
                        nc_names=[row.customer for row in amount_candidates],
                    )
                    issue_rows = [row.row_index for row in amount_candidates]
                else:
                    name_candidates = [
                        nc_row
                        for nc_row in nc_rows
                        if names_match(excel_row.payer_name, nc_row.customer)
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


class ReceiptEntryDryRunMatcher:
    def match(self, excel_rows, nc_rows):
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
                if names_match(excel_row.payer_name, nc_row.name)
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
                        nc_names=[row.name for row in amount_candidates],
                    )
                    issue_rows = [row.row_index for row in amount_candidates]
                else:
                    name_candidates = [
                        nc_row
                        for nc_row in nc_rows
                        if names_match(excel_row.payer_name, nc_row.name)
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


def format_receipt_not_found_reason():
    return "金额和对手方均未匹配"


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


def plan_issue(
    excel_row,
    stage,
    issue_type,
    field,
    raw_value,
    config_node,
    message,
    action="跳过本行",
):
    return ReceiptPlanIssue(
        excel_row=excel_row,
        stage=stage,
        issue_type=issue_type,
        field=field,
        raw_value=format_receipt_value(raw_value),
        config_node=config_node,
        message=message,
        action=action,
    )


def normalize_receipt_currency(value):
    text = str(value or "").strip().upper()
    if text in {"USD", "美元"}:
        return "美元"
    if text in {"RMB", "CNY", "人民币"}:
        return "人民币"
    return ""


def make_receipt_duplicate_key(
    organization_code,
    receipt_date,
    bank,
    currency,
    customer_code,
    payer_name,
    amount,
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


def plan_sheet_row(row, issue, status):
    if row is None:
        base = [""] * 10
    else:
        total_amount = row.raw_amount + row.fee
        base = [
            row.row,
            row.organization_name,
            row.receipt_date.isoformat(),
            row.customer_code,
            row.currency,
            row.payer_name,
            str(row.raw_amount),
            str(row.fee),
            str(total_amount),
            row.account_no,
        ]
    issue_reason = format_plan_issue_reason(issue)
    return [
        *base,
        status,
        issue_reason,
    ]


def format_plan_issue_reason(issue):
    if issue is None:
        return ""
    return f"本地预检：{plan_issue_summary(issue)}"


def plan_issue_summary(issue):
    summaries = {
        "EXCEL_REQUIRED_COLUMN_MISSING": "缺少必需列",
        "EXCEL_START_ROW_INVALID": "起始行配置错误",
        "DATE_INVALID": "到款日期格式错误",
        "PAYER_NAME_EMPTY": "银行来款名为空",
        "AMOUNT_ZERO_OR_NEGATIVE": "实收金额必须大于0",
        "AMOUNT_INVALID": "实收金额格式错误",
        "BANK_EMPTY": "银行为空",
        "BANK_ACCOUNT_NOT_CONFIGURED": "银行未配置",
        "BANK_ACCOUNT_DISABLED": "银行账户已禁用",
        "ORG_NOT_CONFIGURED": "执行主体未配置",
        "CURRENCY_EMPTY": "币种为空",
        "CURRENCY_UNSUPPORTED": "币种不支持",
        "CUSTOMER_CODE_EMPTY": "客户编码为空",
        "FEE_NEGATIVE": "手续费不能小于0",
        "FEE_INVALID": "手续费格式错误",
        "DETAIL_ACCOUNT_CANDIDATE_MISSING": "收款银行账户候选缺失",
        "DUPLICATE_EXCEL_ROWS": "本批存在重复行",
    }
    return summaries.get(issue.issue_type) or str(issue.message or "预检失败")


def ensure_result_sheet_headers(ws, header_row):
    for column in range(ws.max_column, 0, -1):
        value = ws.cell(header_row, column).value
        text = str(value or "").strip()
        if text in DEPRECATED_RESULT_SHEET_HEADERS:
            ws.delete_cols(column)

    columns = {}
    for column in range(1, ws.max_column + 1):
        value = ws.cell(header_row, column).value
        text = str(value or "").strip()
        if text:
            columns[text] = column
    next_column = max(columns.values(), default=0) + 1
    for header in RESULT_SHEET_HEADERS:
        if header in columns:
            continue
        ws.cell(row=header_row, column=next_column, value=header)
        columns[header] = next_column
        next_column += 1
    return columns


def append_plan_sheet_row(ws, columns, row_number, values):
    row_by_header = dict(zip(RESULT_SHEET_HEADERS, values, strict=True))
    for header, value in row_by_header.items():
        ws.cell(row=row_number, column=columns[header], value=value)
    return row_number + 1


def plan_sheet_sort_key(row):
    return (
        str(row.organization_code or ""),
        str(row.organization_name or ""),
        int(row.row or 0),
    )


def orphan_issue_sort_key(item, rows_by_number):
    row_number, _row_issues = item
    row = rows_by_number.get(row_number)
    if row is not None:
        return plan_sheet_sort_key(row)
    return ("ZZZ", "", int(row_number or 0))


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
