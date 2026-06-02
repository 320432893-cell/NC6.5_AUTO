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
            "candidate_check": {
                "recent_months": 2,
                "from_date": None,
                "only_blank_status": True,
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

    rows, candidates, issues = ReceiptEntryWorkbook(
        receipt_config(path)
    ).ensure_output_columns_and_subjects(today=date(2026, 1, 20))

    assert issues == []
    assert len(rows) == 1
    assert candidates == rows
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


def test_candidate_rows_use_recent_months_and_blank_status(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行", "是否NC已做过"])
    ws.append([date(2026, 3, 31), "old recent excluded", 100, "Paypal", None])
    ws.append([date(2026, 4, 2), "already done", 200, "Paypal", "已做过"])
    ws.append([date(2026, 4, 2), "candidate", 300, "Paypal", None])
    wb.save(path)
    wb.close()

    rows, candidates, issues = ReceiptEntryWorkbook(receipt_config(path)).preview_rows(
        today=date(2026, 6, 2)
    )

    assert issues == []
    assert len(rows) == 3
    assert [row.payer_name for row in candidates] == ["candidate"]


def test_candidate_from_date_overrides_recent_months(tmp_path):
    path = tmp_path / "payments.xlsx"
    config = receipt_config(path)
    config["receipt_entry"]["candidate_check"]["from_date"] = "2026-05-01"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行"])
    ws.append([date(2026, 4, 30), "old", 100, "Paypal"])
    ws.append([date(2026, 5, 1), "candidate", 200, "Paypal"])
    wb.save(path)
    wb.close()

    rows, candidates, issues = ReceiptEntryWorkbook(config).preview_rows(
        today=date(2026, 6, 2)
    )

    assert issues == []
    assert len(rows) == 2
    assert [row.payer_name for row in candidates] == ["candidate"]


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
        nc_done_status="",
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
        nc_done_status="",
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
