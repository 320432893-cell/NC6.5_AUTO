# 职责：从已读取的 NC 结果表结构中抽取收款单行和抽取问题
# 不做什么：不打开 NC 页面，不读取/写入 Excel，不解析全量配置，不执行 JAB/GUI 动作
# 允许依赖层：收款单数据模型、解析纯函数和调用方传入的列配置
# 谁不应该 import：底层 JAB 操作模块不应 import

from core.receipt_models import (
    ReceiptNCExtractIssue,
    ReceiptNCIndexedRow,
    ReceiptNCRow,
)
from core.receipt_parsing import normalize_lookup_key, parse_amount, parse_date
from core.receipt_config import ReceiptEntryConfig


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
