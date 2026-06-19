# 生命周期：持久维护
# 覆盖的业务场景：收款单录入的对手方匹配：名称归一化、金额/名称匹配与 dry-run 报告
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_entry_counterparty_match.py


from tests._receipt_entry_helpers import (
    Decimal,
    ReceiptEntryDryRunMatcher,
    ReceiptEntryMatcher,
    ReceiptExcelRow,
    ReceiptNCIndexedRow,
    ReceiptNCResultExtractor,
    ReceiptNCRow,
    date,
    format_receipt_amount_name_mismatch_reason,
    format_receipt_duplicate_reason,
    format_receipt_name_amount_mismatch_reason,
    format_receipt_not_found_reason,
    names_match,
    normalize_counterparty,
)


def test_counterparty_normalization_ignores_prefix_and_punctuation():
    assert normalize_counterparty("1/AZUGA INC. AZUGA INC - OPERATING") == (
        "AZUGAINCAZUGAINCOPERATING"
    )
    assert names_match("1/AZUGA INC. AZUGA INC - OPERATING", "AZUGA INC")


def test_receipt_matcher_matches_amount_and_name_even_when_dates_differ():
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
        document_date=date(2026, 1, 17),
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
    assert issues[0].reason == format_receipt_duplicate_reason(len(nc_rows))


def test_dry_run_matcher_reports_same_name_different_amount():
    excel_row = ReceiptExcelRow(
        row=10,
        receipt_date=date(2026, 3, 31),
        payer_name="Christoff Pretorius",
        raw_amount=Decimal("100.00"),
        bank="大陆花旗",
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
        nc_done_status="",
    )
    nc_row = ReceiptNCIndexedRow(
        row_index=7,
        table_index=2,
        document_no="D7",
        document_date=date(2026, 4, 1),
        original_amount=Decimal("120.00"),
        name="Christoff Pretorius",
    )

    matched, issues = ReceiptEntryDryRunMatcher().match([excel_row], [nc_row])

    assert matched == {}
    assert len(issues) == 1
    assert issues[0].reason == format_receipt_name_amount_mismatch_reason(
        excel_amount=Decimal("100.00"),
        excel_name="Christoff Pretorius",
        nc_amounts=[Decimal("120.00")],
    )
    assert issues[0].nc_rows == [7]


def test_dry_run_matcher_reports_same_amount_different_name():
    excel_row = ReceiptExcelRow(
        row=10,
        receipt_date=date(2026, 3, 31),
        payer_name="Christoff Pretorius",
        raw_amount=Decimal("100.00"),
        bank="大陆花旗",
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
        nc_done_status="",
    )
    nc_row = ReceiptNCIndexedRow(
        row_index=8,
        table_index=2,
        document_no="D8",
        document_date=date(2026, 4, 1),
        original_amount=Decimal("100.00"),
        name="Different Payer",
    )

    matched, issues = ReceiptEntryDryRunMatcher().match([excel_row], [nc_row])

    assert matched == {}
    assert len(issues) == 1
    assert issues[0].reason == format_receipt_amount_name_mismatch_reason(
        excel_amount=Decimal("100.00"),
        excel_name="Christoff Pretorius",
        nc_names=["Different Payer"],
    )
    assert issues[0].nc_rows == [8]


def test_dry_run_matcher_reports_no_amount_or_name_match():
    excel_row = ReceiptExcelRow(
        row=10,
        receipt_date=date(2026, 3, 31),
        payer_name="Christoff Pretorius",
        raw_amount=Decimal("100.00"),
        bank="大陆花旗",
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
        nc_done_status="",
    )
    nc_row = ReceiptNCIndexedRow(
        row_index=9,
        table_index=2,
        document_no="D9",
        document_date=date(2026, 4, 1),
        original_amount=Decimal("120.00"),
        name="Different Payer",
    )

    matched, issues = ReceiptEntryDryRunMatcher().match([excel_row], [nc_row])

    assert matched == {}
    assert len(issues) == 1
    reason = issues[0].reason
    # 业务意图：保留“未命中”诊断标记作为写回分类前缀，并带上行号/金额/对手方/下一步。
    assert reason.startswith(format_receipt_not_found_reason())
    assert "第10行" in reason
    assert "100.00" in reason
    assert "Christoff Pretorius" in reason
    assert "核对单据" in reason
    assert issues[0].nc_rows == []


def test_dry_run_matcher_compares_chosen_nc_name_field():
    excel_row = ReceiptExcelRow(
        row=10,
        receipt_date=date(2026, 3, 31),
        payer_name="BANK PAYER",
        raw_amount=Decimal("37165.00"),
        bank="大陆花旗",
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
        nc_done_status="",
    )
    nc_row = ReceiptNCResultExtractor(
        {
            "receipt_entry": {
                "query": {
                    "result_column_indexes": {
                        "document_no": 0,
                        "document_date": 1,
                        "customer": 4,
                        "original_amount": 6,
                        "payer_name": 19,
                    }
                },
                "finance_organizations": [],
                "accounts": [],
            }
        }
    ).extract_by_indexes(
        [
            {
                "table_index": 2,
                "row_count": 1,
                "col_count": 41,
                "rows": [
                    {
                        "row_index": 0,
                        "cells": [
                            "D1",
                            "2026-03-31",
                            "",
                            "",
                            "NC CUSTOMER",
                            "",
                            "37,165.00",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "BANK PAYER",
                        ],
                    }
                ],
            }
        ],
        19,
    )[0][0]

    matched, issues = ReceiptEntryDryRunMatcher().match([excel_row], [nc_row])

    assert matched == {10: nc_row}
    assert issues == []


def test_dry_run_matcher_ignores_date_and_matches_amount_name():
    excel_row = ReceiptExcelRow(
        row=10,
        receipt_date=date(2026, 3, 31),
        payer_name="BANK PAYER",
        raw_amount=Decimal("37165.00"),
        bank="大陆花旗",
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
        nc_done_status="",
    )
    nc_row = ReceiptNCIndexedRow(
        row_index=0,
        table_index=2,
        document_no="D1",
        document_date=date(2026, 4, 1),
        original_amount=Decimal("37165.00"),
        name="BANK PAYER",
    )

    matched, issues = ReceiptEntryDryRunMatcher().match([excel_row], [nc_row])

    assert matched == {10: nc_row}
    assert issues == []


def test_dry_run_matcher_reports_duplicate_amount_name():
    excel_row = ReceiptExcelRow(
        row=10,
        receipt_date=date(2026, 3, 31),
        payer_name="BANK PAYER",
        raw_amount=Decimal("37165.00"),
        bank="大陆花旗",
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
        nc_done_status="",
    )
    nc_rows = [
        ReceiptNCIndexedRow(
            row_index=0,
            table_index=2,
            document_no="D1",
            document_date=date(2026, 4, 1),
            original_amount=Decimal("37165.00"),
            name="BANK PAYER",
        ),
        ReceiptNCIndexedRow(
            row_index=1,
            table_index=2,
            document_no="D2",
            document_date=date(2026, 4, 2),
            original_amount=Decimal("37165.00"),
            name="1/BANK PAYER",
        ),
    ]

    matched, issues = ReceiptEntryDryRunMatcher().match([excel_row], nc_rows)

    assert matched == {}
    assert len(issues) == 1
    assert issues[0].reason == format_receipt_duplicate_reason(len(nc_rows))
