# 生命周期：持久维护
# 覆盖的业务场景：收款单录入的 NC 多页抽取：按配置索引抽取、跨分页表收集
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_entry_nc_paged.py


from tests._receipt_entry_helpers import (
    Decimal,
    ReceiptNCResultExtractor,
    date,
    receipt_config,
)


def test_extract_receipt_nc_rows_by_configured_indexes():
    tables = [
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
                        "收款单",
                        "收款结算",
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
                        "memo",
                        "BANK PAYER",
                    ],
                }
            ],
        }
    ]
    config = receipt_config()
    config["receipt_entry"]["query"] = {
        "result_column_indexes": {
            "document_no": 0,
            "document_date": 1,
            "customer": 4,
            "original_amount": 6,
            "payer_name": 19,
        }
    }

    rows, issues = ReceiptNCResultExtractor(config).extract_by_indexes(tables, 19)

    assert issues == []
    assert len(rows) == 1
    assert rows[0].document_no == "D1"
    assert rows[0].document_date == date(2026, 3, 31)
    assert rows[0].original_amount == Decimal("37165.00")
    assert rows[0].name == "BANK PAYER"
    assert rows[0].table_index == 2


def test_extract_receipt_nc_rows_by_indexes_collects_all_paged_tables():
    config = receipt_config()
    config["receipt_entry"]["query"] = {
        "result_column_indexes": {
            "document_no": 0,
            "document_date": 1,
            "customer": 2,
            "original_amount": 8,
            "payer_name": 2,
        }
    }
    tables = [
        {
            "table_index": 4,
            "row_count": 2,
            "col_count": 32,
            "rows": [
                {
                    "row_index": 0,
                    "cells": [
                        "D1",
                        "2026-03-31",
                        "PAYER A",
                        "",
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
                        "2026-04-01",
                        "PAYER B",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "200.00",
                    ],
                },
            ],
        },
        {
            "table_index": 4,
            "row_count": 2,
            "col_count": 32,
            "rows": [
                {
                    "row_index": 0,
                    "cells": [
                        "D2",
                        "2026-04-01",
                        "PAYER B",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "200.00",
                    ],
                },
                {
                    "row_index": 1,
                    "cells": [
                        "D3",
                        "2026-04-02",
                        "PAYER C",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "300.00",
                    ],
                },
            ],
        },
    ]

    rows, issues = ReceiptNCResultExtractor(config).extract_by_indexes(tables, 2)

    assert issues == []
    assert [row.document_no for row in rows] == ["D1", "D2", "D3"]
    assert [row.original_amount for row in rows] == [
        Decimal("100.00"),
        Decimal("200.00"),
        Decimal("300.00"),
    ]
