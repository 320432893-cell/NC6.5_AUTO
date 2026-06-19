# 生命周期：持久维护
# 覆盖的业务场景：收款单录入的 NC 结果抽取：按表头标签抽取、坏值/缺表头诊断、配置列名
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_entry_nc_extract.py


from tests._receipt_entry_helpers import (
    Decimal,
    RESULT_SHEET_HEADERS,
    ReceiptNCResultExtractor,
    ReceiptNCRow,
    date,
    extract_receipt_nc_rows,
    parse_amount,
    parse_amount_from_new_module,
    receipt_config,
)


def test_extract_receipt_nc_rows_uses_header_labels_not_fixed_indexes():
    tables = [
        {
            "table_index": 2,
            "rows": [
                {
                    "row_index": 0,
                    "cells": ["选择", "客户", "备注", "原币金额", "单据日期"],
                },
                {
                    "row_index": 1,
                    "cells": ["", "AZUGA INC", "", "68,700.00", "2026-01-16"],
                },
                {
                    "row_index": 2,
                    "cells": ["", "", "", "", ""],
                },
            ],
        }
    ]

    rows, issues = extract_receipt_nc_rows(
        tables,
        {
            "document_date": "单据日期",
            "original_amount": "原币金额",
            "customer": "客户",
        },
    )

    assert issues == []
    assert rows == [
        ReceiptNCRow(
            row_index=1,
            document_date=date(2026, 1, 16),
            customer="AZUGA INC",
            original_amount=Decimal("68700.00"),
        )
    ]


def test_extract_receipt_nc_rows_reports_bad_result_values():
    tables = [
        {
            "table_index": 2,
            "rows": [
                {
                    "row_index": 0,
                    "cells": ["单据日期", "客户", "原币金额"],
                },
                {
                    "row_index": 1,
                    "cells": ["2026-01-16", "AZUGA INC", "not money"],
                },
            ],
        }
    ]

    rows, issues = extract_receipt_nc_rows(
        tables,
        {
            "document_date": "单据日期",
            "original_amount": "原币金额",
            "customer": "客户",
        },
    )

    assert rows == []
    assert len(issues) == 1
    assert issues[0].table_index == 2
    assert issues[0].row_index == 1
    assert "原始金额格式无法识别" in issues[0].reason


def test_parse_amount_accepts_nc_negative_spacing():
    assert parse_amount("- 1,368.10") == Decimal("-1368.10")
    assert parse_amount("(1,368.10)") == Decimal("-1368.10")


def test_receipt_split_modules_export_new_paths():
    assert parse_amount_from_new_module("- 1,368.10") == Decimal("-1368.10")
    assert "异常原因" in RESULT_SHEET_HEADERS


def test_extract_receipt_nc_rows_reports_missing_header():
    rows, issues = extract_receipt_nc_rows(
        [{"table_index": 1, "rows": [{"row_index": 0, "cells": ["日期", "金额"]}]}],
        {
            "document_date": "单据日期",
            "original_amount": "原币金额",
            "customer": "客户",
        },
    )

    assert rows == []
    assert len(issues) == 1
    assert issues[0].table_index is None
    assert "未找到包含结果列" in issues[0].reason


def test_receipt_nc_result_extractor_uses_configured_result_column_names():
    config = receipt_config()
    config["receipt_entry"]["query"] = {
        "result_columns": {
            "document_date": "日期",
            "original_amount": "金额",
            "customer": "往来客户",
        }
    }
    tables = [
        {
            "table_index": 3,
            "rows": [
                {"row_index": 0, "cells": ["往来客户", "金额", "日期"]},
                {"row_index": 1, "cells": ["客户A", "100", "2026/01/17"]},
            ],
        }
    ]

    rows, issues = ReceiptNCResultExtractor(config).extract(tables)

    assert issues == []
    assert rows == [
        ReceiptNCRow(
            row_index=1,
            document_date=date(2026, 1, 17),
            customer="客户A",
            original_amount=Decimal("100.00"),
        )
    ]
