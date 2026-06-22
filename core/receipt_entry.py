from decimal import Decimal

import openpyxl

from core.errors import ExcelLockedError, WorkflowStateError
from core.receipt_config import (
    ReceiptEntryConfig as ReceiptEntryConfig,
    days_in_month as days_in_month,
    default_account_id as default_account_id,
    normalize_candidate_map as normalize_candidate_map,
    subtract_months as subtract_months,
)
from core.receipt_matching import (
    ReceiptEntryDryRunMatcher as ReceiptEntryDryRunMatcher,
    ReceiptEntryMatcher as ReceiptEntryMatcher,
    format_receipt_amount_name_mismatch_reason as format_receipt_amount_name_mismatch_reason,
    format_receipt_duplicate_reason as format_receipt_duplicate_reason,
    format_receipt_name_amount_mismatch_reason as format_receipt_name_amount_mismatch_reason,
    format_receipt_not_found_reason as format_receipt_not_found_reason,
    names_match as names_match,
    normalize_counterparty as normalize_counterparty,
)
from core.receipt_models import (
    ReceiptExcelRow,
    ReceiptMatchIssue,
    ReceiptNCExtractIssue as ReceiptNCExtractIssue,
    ReceiptNCIndexedRow as ReceiptNCIndexedRow,
    ReceiptNCRow as ReceiptNCRow,
    ReceiptPlanIssue as ReceiptPlanIssue,
)
from core.receipt_nc_extract import (
    ReceiptNCResultExtractor as ReceiptNCResultExtractor,
    extract_receipt_nc_rows as extract_receipt_nc_rows,
    extract_receipt_nc_rows_by_indexes as extract_receipt_nc_rows_by_indexes,
    is_blank_result_row as is_blank_result_row,
    read_cell as read_cell,
    read_required_text as read_required_text,
    resolve_receipt_result_columns as resolve_receipt_result_columns,
)
from core.receipt_parsing import (
    parse_amount,
    parse_date,
)
from core.receipt_plan import (
    build_plan_rows,
)
from core.receipt_plan_issue import (
    detect_duplicate_rows,
    plan_issue as plan_issue,
    read_optional_cell,
    summarize_plan,
)

# 兼容别名：旧调用方仍可从 core.receipt_entry import；待生产脚本和测试迁到
# core.receipt_models/core.receipt_matching/core.receipt_nc_extract/
# core.receipt_parsing/core.receipt_sheet 后删除这些 re-export。
from core.receipt_sheet import (
    DEPRECATED_RESULT_SHEET_HEADERS as DEPRECATED_RESULT_SHEET_HEADERS,
    RESULT_SHEET_HEADERS as RESULT_SHEET_HEADERS,
    append_plan_sheet_row as append_plan_sheet_row,
    ensure_result_sheet_headers as ensure_result_sheet_headers,
    orphan_issue_sort_key as orphan_issue_sort_key,
    plan_sheet_row as plan_sheet_row,
    plan_sheet_sort_key as plan_sheet_sort_key,
    rewrite_batch_result_sheet,
    rewrite_plan_sheet,
)


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
            rows, issues = build_plan_rows(self.config, ws, columns)
            duplicate_issues = detect_duplicate_rows(rows)
            issues.extend(duplicate_issues)
        finally:
            wb.close()
        if write_sheet:
            writable = openpyxl.load_workbook(self.excel_path, read_only=False)
            try:
                self._write_plan_sheet(writable, rows, issues)
            finally:
                writable.close()
        return rows, issues, summarize_plan(rows, issues, self.config.validation_policy)

    def write_plan_sheet(self, rows, issues):
        writable = openpyxl.load_workbook(self.excel_path, read_only=False)
        try:
            self._write_plan_sheet(writable, rows, issues)
        finally:
            writable.close()

    def write_batch_result_sheet(self, results):
        writable = openpyxl.load_workbook(self.excel_path, read_only=False)
        try:
            rewrite_batch_result_sheet(
                writable,
                self.config.result_sheet_name,
                self.config.header_row,
                results,
            )
            writable.save(self.excel_path)
        finally:
            writable.close()

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
            # 读手续费,使 dry-run 匹配按 NC 口径(raw+fee)对齐;空/异常按 0 兜底
            # (严格的费用校验是录入计划 build_plan_rows 的职责,此处仅供匹配口径)
            fee_raw = read_optional_cell(
                ws, row_index, columns.get(self.config.fee_column)
            )
            try:
                fee = (
                    parse_amount(fee_raw)
                    if fee_raw not in (None, "")
                    else Decimal("0.00")
                )
            except ValueError:
                fee = Decimal("0.00")
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
                    fee=fee,
                )
            )
        return rows, issues

    def _write_plan_sheet(self, wb, rows, issues):
        rewrite_plan_sheet(
            wb,
            self.config.result_sheet_name,
            self.config.header_row,
            rows,
            issues,
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
