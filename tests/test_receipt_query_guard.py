# 生命周期：持久维护
# 覆盖的业务场景：收款单查询的 guard 操作：父页状态校验与结果表单据类型校验
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake JAB 替身）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_query_guard.py


from tests._receipt_query_helpers import (
    FakeGuardJAB,
    ReceiptPageGuardError,
    guard_receipt_parent_page,
    guard_receipt_result_tables,
    pytest,
)


def test_guard_receipt_parent_page_requires_state_label():
    with pytest.raises(ReceiptPageGuardError, match="收款单录入"):
        guard_receipt_parent_page(
            FakeGuardJAB(found=False),
            {"receipt_entry": {"state_label": "收款单录入"}},
            {},
        )


def test_guard_receipt_parent_page_releases_found_context():
    jab = FakeGuardJAB(found=True)

    report = guard_receipt_parent_page(
        jab,
        {"receipt_entry": {"state_label": "收款单录入"}},
        {},
    )

    assert report["ok"] is True
    assert jab.released == [(22, [11])]


def test_guard_receipt_result_tables_blocks_wrong_document_type():
    tables = [
        {
            "table_index": 4,
            "row_count": 1,
            "col_count": 41,
            "rows": [
                {
                    "row_index": 0,
                    "cells": [
                        "D1",
                        "2026-05-25",
                        "收款单",
                        "收款结算",
                        "客户A",
                        "",
                        "100.00",
                        "100.00",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "应收款",
                    ],
                }
            ],
        }
    ]

    with pytest.raises(ReceiptPageGuardError, match="错误页面"):
        guard_receipt_result_tables(
            tables,
            {
                "result_column_indexes": {
                    "document_no": 0,
                    "document_date": 1,
                    "customer": 2,
                    "original_amount": 7,
                    "payer_name": 2,
                },
                "result_guard": {
                    "document_type_column": 16,
                    "document_type": "收款单录入",
                    "blocked_keywords": ["应收款"],
                },
            },
        )


def test_guard_receipt_result_tables_accepts_receipt_document_type():
    report = guard_receipt_result_tables(
        [
            {
                "table_index": 4,
                "row_count": 1,
                "col_count": 8,
                "rows": [
                    {
                        "row_index": 0,
                        "cells": [
                            "D1",
                            "2026-05-25",
                            "收款单",
                            "",
                            "客户A",
                            "",
                            "PAYER A",
                            "",
                            "100.00",
                        ],
                    }
                ],
            }
        ],
        {
            "result_column_indexes": {
                "document_no": 0,
                "document_date": 1,
                "customer": 2,
                "original_amount": 8,
                "payer_name": 6,
            },
            "result_guard": {"document_type_column": 2, "document_type": "收款单"},
        },
    )

    assert report["ok"] is True


def test_guard_receipt_result_tables_blocks_name_column_as_document_type():
    with pytest.raises(ReceiptPageGuardError, match="名称列"):
        guard_receipt_result_tables(
            [
                {
                    "table_index": 4,
                    "row_count": 1,
                    "col_count": 8,
                    "rows": [
                        {
                            "row_index": 0,
                            "cells": [
                                "D1",
                                "2026-05-25",
                                "收款单",
                                "",
                                "",
                                "",
                                "",
                                "100.00",
                            ],
                        }
                    ],
                }
            ],
            {
                "result_column_indexes": {
                    "document_no": 0,
                    "document_date": 1,
                    "customer": 2,
                    "original_amount": 7,
                    "payer_name": 2,
                },
                "result_guard": {
                    "document_type_column": 2,
                    "document_type": "收款单",
                    "name_column_must_not_equal_document_type": True,
                },
            },
        )
