from datetime import date
from decimal import Decimal

from openpyxl import Workbook, load_workbook

from core.receipt_entry import (
    ReceiptEntryConfig,
    ReceiptEntryMatcher,
    ReceiptEntryWorkbook,
    ReceiptExcelRow,
    ReceiptNCRow,
    names_match,
    normalize_counterparty,
)


def receipt_config(path="unused.xlsx"):
    return {
        "receipt_entry": {
            "state_label": "收款单录入",
            "excel": {
                "path": str(path),
                "sheet_name": "💸Payments来款通知",
                "header_row": 1,
                "start_date": "2026-01-01",
                "date_column": "到款日期",
                "payer_name_column": "🟪银行来款名",
                "raw_amount_column": "🟪原始金额",
                "bank_column": "银行",
                "organization_column": "主体名称",
                "nc_done_column": "是否NC已做过",
            },
            "finance_organizations": [
                {
                    "code": "A001",
                    "name": "上海移为通信技术股份有限公司",
                    "short_name": "移为",
                },
                {
                    "code": "A006",
                    "name": "上海移为通信技术（香港）有限公司",
                    "short_name": "移为香港",
                },
            ],
            "accounts": [
                {
                    "organization_code": "A001",
                    "organization_short_name": "移为",
                    "account_label": "PayPal",
                    "account_no": "paypal",
                },
                {
                    "organization_code": "A006",
                    "organization_short_name": "移为香港",
                    "account_label": "香港花旗",
                    "account_no": "1778667904",
                },
            ],
        }
    }


def test_bank_label_maps_to_organization_case_insensitive():
    config = ReceiptEntryConfig(receipt_config())

    organization = config.organization_for_bank("Paypal")

    assert organization is not None
    assert organization.code == "A001"
    assert organization.name == "上海移为通信技术股份有限公司"


def test_ensure_output_columns_and_subjects(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行"])
    ws.append([date(2026, 1, 16), "lamine Mohamed", 225.68, "Paypal"])
    ws.append([date(2025, 12, 31), "old", 1, "Paypal"])
    wb.save(path)
    wb.close()

    rows, issues = ReceiptEntryWorkbook(
        receipt_config(path)
    ).ensure_output_columns_and_subjects()

    assert issues == []
    assert len(rows) == 1
    assert rows[0].organization_code == "A001"

    saved = load_workbook(path)
    ws = saved["💸Payments来款通知"]
    headers = [ws.cell(1, column).value for column in range(1, ws.max_column + 1)]
    assert headers == [
        "到款日期",
        "🟪银行来款名",
        "🟪原始金额",
        "银行",
        "主体名称",
        "是否NC已做过",
    ]
    assert ws.cell(2, 5).value == "上海移为通信技术股份有限公司"
    assert ws.cell(3, 5).value is None
    saved.close()


def test_counterparty_normalization_ignores_prefix_and_punctuation():
    assert normalize_counterparty("1/AZUGA INC. AZUGA INC - OPERATING") == (
        "AZUGAINCAZUGAINCOPERATING"
    )
    assert names_match("1/AZUGA INC. AZUGA INC - OPERATING", "AZUGA INC")


def test_receipt_matcher_matches_date_amount_and_contained_name():
    excel_row = ReceiptExcelRow(
        row=10,
        receipt_date=date(2026, 1, 16),
        payer_name="1/AZUGA INC. AZUGA INC - OPERATING",
        raw_amount=Decimal("68700.00"),
        bank="大陆花旗",
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
    )
    nc_row = ReceiptNCRow(
        row_index=3,
        document_date=date(2026, 1, 16),
        customer="AZUGA INC",
        original_amount=Decimal("68700.00"),
    )

    matched, issues = ReceiptEntryMatcher().match([excel_row], [nc_row])

    assert matched == {10: nc_row}
    assert issues == []


def test_receipt_matcher_reports_duplicate_as_exception_issue():
    excel_row = ReceiptExcelRow(
        row=10,
        receipt_date=date(2026, 1, 16),
        payer_name="AZUGA INC",
        raw_amount=Decimal("68700.00"),
        bank="大陆花旗",
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
    )
    nc_rows = [
        ReceiptNCRow(
            row_index=3,
            document_date=date(2026, 1, 16),
            customer="AZUGA INC",
            original_amount=Decimal("68700.00"),
        ),
        ReceiptNCRow(
            row_index=4,
            document_date=date(2026, 1, 16),
            customer="AZUGA INC",
            original_amount=Decimal("68700.00"),
        ),
    ]

    matched, issues = ReceiptEntryMatcher().match([excel_row], nc_rows)

    assert matched == {}
    assert len(issues) == 1
    assert issues[0].reason == "重复2条"
    assert issues[0].nc_rows == [3, 4]
