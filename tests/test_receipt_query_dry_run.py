# 生命周期：持久维护
# 覆盖的业务场景：收款单查询的 dry-run 匹配报告：状态写入、已填充包含、增量复用与全量变体
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake JAB 替身）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_query_dry_run.py


from tests._receipt_query_helpers import (
    Decimal,
    ReceiptExcelRow,
    ReceiptMatchIssue,
    ReceiptNCExtractIssue,
    ReceiptNCIndexedRow,
    ReceiptNCResultExtractor,
    Workbook,
    build_dry_run_match_report,
    build_dry_run_match_report_from_preview,
    date,
    format_receipt_duplicate_reason,
    format_receipt_name_amount_mismatch_reason,
    load_workbook,
    receipt_config,
)


def test_dry_run_match_report_writes_specific_statuses(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行", "是否NC已做过"])
    ws.append([date(2026, 5, 1), "MATCHED INC", 100, "Paypal", None])
    ws.append([date(2026, 5, 2), "MISSING INC", 200, "Paypal", None])
    ws.append([date(2026, 5, 3), "DUP INC", 300, "Paypal", None])
    ws.append([date(2026, 5, 4), "Christoff Pretorius", 400, "Paypal", None])
    wb.save(path)
    wb.close()
    config = receipt_config(path)
    tables = [
        {
            "table_index": 2,
            "row_count": 3,
            "col_count": 8,
            "rows": [
                {
                    "row_index": 0,
                    "cells": [
                        "D1",
                        "2026-05-01",
                        "MATCHED INC",
                        "",
                        "",
                        "",
                        "",
                        "100.00",
                    ],
                },
                {
                    "row_index": 1,
                    "cells": [
                        "D2",
                        "2026-05-03",
                        "DUP INC",
                        "",
                        "",
                        "",
                        "",
                        "300.00",
                    ],
                },
                {
                    "row_index": 2,
                    "cells": [
                        "D3",
                        "2026-05-03",
                        "1/DUP INC",
                        "",
                        "",
                        "",
                        "",
                        "300.00",
                    ],
                },
                {
                    "row_index": 3,
                    "cells": [
                        "D4",
                        "2026-05-04",
                        "Christoff Pretorius",
                        "",
                        "",
                        "",
                        "",
                        "450.00",
                    ],
                },
            ],
        }
    ]

    report = build_dry_run_match_report(
        config,
        ReceiptNCResultExtractor(config),
        tables,
        org_code="A001",
        business_date=date(2026, 6, 1),
        write_back=True,
    )

    assert report["write_back"]["updated"] == 4
    assert report["write_back"]["matched_rows"] == [2]
    assert report["write_back"]["not_found_rows"] == [3]
    assert report["write_back"]["duplicate_rows"] == [4]
    assert report["write_back"]["exception_rows"] == [4, 5]
    assert report["write_back"]["skipped_duplicate_rows"] == [4]
    saved = load_workbook(path)
    ws = saved["💸Payments来款通知"]
    assert ws.cell(2, 5).value == "已做过"
    assert ws.cell(3, 5).value == "未做过"
    assert ws.cell(4, 5).value == format_receipt_duplicate_reason(2)
    assert ws.cell(5, 5).value == format_receipt_name_amount_mismatch_reason(
        excel_amount=Decimal("400.00"),
        excel_name="Christoff Pretorius",
        nc_amounts=["450.00"],
    )
    saved.close()


def test_dry_run_match_report_can_include_already_filled_statuses(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行", "是否NC已做过"])
    ws.append([date(2026, 5, 1), "MATCHED INC", 100, "Paypal", "未做过"])
    wb.save(path)
    wb.close()
    config = receipt_config(path)
    config["receipt_entry"]["candidate_check"]["only_blank_status"] = False
    tables = [
        {
            "table_index": 2,
            "row_count": 1,
            "col_count": 8,
            "rows": [
                {
                    "row_index": 0,
                    "cells": [
                        "D1",
                        "2026-05-01",
                        "MATCHED INC",
                        "",
                        "",
                        "",
                        "",
                        "100.00",
                    ],
                },
            ],
        }
    ]

    report = build_dry_run_match_report(
        config,
        ReceiptNCResultExtractor(config),
        tables,
        org_code="A001",
        business_date=date(2026, 6, 1),
        write_back=True,
    )

    assert report["org_candidates"] == 1
    assert report["write_back"]["updated"] == 1
    saved = load_workbook(path)
    ws = saved["💸Payments来款通知"]
    assert ws.cell(2, 5).value == "已做过"
    saved.close()


def test_dry_run_match_report_reuses_incremental_configured_match():
    class CountingExtractor:
        def __init__(self):
            self.config = ReceiptNCResultExtractor(receipt_config("unused.xlsx")).config
            self.calls = []

        def extract_by_indexes(self, tables, name_column, amount_column=None):
            self.calls.append((name_column, amount_column))
            return [], []

    excel_row = ReceiptExcelRow(
        row=2,
        receipt_date=date(2026, 5, 1),
        payer_name="MATCHED INC",
        raw_amount=Decimal("100.00"),
        bank="Paypal",
        organization_code="A001",
        organization_name="A001",
        organization_short_name="A001",
        nc_done_status="",
    )
    nc_row = ReceiptNCIndexedRow(
        row_index=0,
        document_date=date(2026, 5, 1),
        original_amount=Decimal("100.00"),
        name="MATCHED INC",
        document_no="D1",
        table_index=2,
    )
    snapshot = {
        "nc_rows": [nc_row],
        "extract_issues": [],
        "matched": {2: nc_row},
        "match_issues": [],
    }
    extractor = CountingExtractor()

    report = build_dry_run_match_report_from_preview(
        receipt_config("unused.xlsx"),
        extractor,
        tables=[],
        org_code="A001",
        business_date=date(2026, 5, 1),
        rows=[excel_row],
        candidates=[excel_row],
        excel_issues=[],
        target_rows=[excel_row],
        configured_match_snapshot=snapshot,
    )

    assert extractor.calls == []
    assert len(report["variants"]) == 1
    assert report["variants"][0]["source"] == "incremental"
    assert report["variants"][0]["matches"] == 1
    assert report["write_back"]["matched_rows"] == [2]


def test_dry_run_match_report_can_run_all_variants_when_enabled():
    class CountingExtractor:
        def __init__(self):
            self.config = ReceiptNCResultExtractor(receipt_config("unused.xlsx")).config
            self.calls = []

        def extract_by_indexes(self, tables, name_column, amount_column=None):
            self.calls.append((name_column, amount_column))
            return [], [
                ReceiptNCExtractIssue(
                    table_index=None,
                    row_index=None,
                    reason="test",
                )
            ]

    excel_row = ReceiptExcelRow(
        row=2,
        receipt_date=date(2026, 5, 1),
        payer_name="MATCHED INC",
        raw_amount=Decimal("100.00"),
        bank="Paypal",
        organization_code="A001",
        organization_name="A001",
        organization_short_name="A001",
        nc_done_status="",
    )
    config = receipt_config("unused.xlsx")
    config["receipt_entry"]["query"]["dry_run_all_variants"] = True
    extractor = CountingExtractor()

    report = build_dry_run_match_report_from_preview(
        config,
        extractor,
        tables=[],
        org_code="A001",
        business_date=date(2026, 5, 1),
        rows=[excel_row],
        candidates=[excel_row],
        excel_issues=[],
        target_rows=[excel_row],
        configured_match_snapshot={
            "nc_rows": [],
            "extract_issues": [],
            "matched": {},
            "match_issues": [
                ReceiptMatchIssue(
                    excel_row=2,
                    reason="未找到",
                    nc_rows=[],
                )
            ],
        },
    )

    assert len(report["variants"]) > 1
    assert report["variants"][0]["source"] == "incremental"
    assert extractor.calls
