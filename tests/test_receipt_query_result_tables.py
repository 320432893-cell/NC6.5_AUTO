# 生命周期：持久维护
# 覆盖的业务场景：收款单查询结果表读取与分页路径解析（dynamic/module-index/前缀推断）
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake JAB 替身）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_query_result_tables.py


from tests._receipt_query_helpers import (
    infer_result_area_prefix_from_table_path,
    paged_query_config,
    read_receipt_tables,
    resolve_receipt_pagination_paths_by_module_index,
    resolve_receipt_pagination_paths_dynamic,
)


def test_read_receipt_tables_keeps_single_row_main_result_and_filters_detail():
    class FakeOneRowResultJAB:
        def read_all_table_cells(self, max_rows=None, max_cols=None):
            return [
                {
                    "table_index": 1,
                    "row_count": 2,
                    "col_count": 25,
                    "rows": [
                        {
                            "row_index": 0,
                            "cells": [
                                "客户",
                                "借方",
                                "应收款",
                                "美元",
                                "FTE",
                                "1002",
                                "6.8",
                            ],
                        },
                        {
                            "row_index": 1,
                            "cells": [
                                "客户",
                                "财务费用",
                                "应收款",
                                "美元",
                                "",
                                "660305",
                                "6.8",
                            ],
                        },
                    ],
                },
                {
                    "table_index": 2,
                    "row_count": 1,
                    "col_count": 41,
                    "rows": [
                        {
                            "row_index": 0,
                            "cells": [
                                "D22026060100026949",
                                "2026-05-27",
                                "收款单",
                                "收款结算",
                                "Copeland Cold Chain LP",
                                "美元",
                                "161713.00",
                            ]
                            + [""] * 34,
                        }
                    ],
                },
            ]

    tables = read_receipt_tables(
        FakeOneRowResultJAB(),
        {
            "result_column_indexes": {
                "document_no": 0,
                "document_date": 1,
                "customer": 4,
                "original_amount": 6,
                "payer_name": 4,
            },
            "result_table_cols": 41,
        },
        max_rows=10,
        max_cols=80,
    )

    assert [table["table_index"] for table in tables] == [2]
    assert tables[0]["row_count"] == 1


def test_infer_result_area_prefix_from_table_path_strips_known_table_suffix():
    assert (
        infer_result_area_prefix_from_table_path(
            "0.0.1.0.0.0.0.4.0.0.0.1.1.0.0.0.1.1.1.0.0.0.0.0.0"
        )
        == "0.0.1.0.0.0.0.4.0.0.0.1.1.0.0.0.1.1.1.0.0.0"
    )
    assert infer_result_area_prefix_from_table_path("0.1.2.3") is None


def test_resolve_receipt_pagination_paths_dynamic_uses_path_not_fixed_column_count(
    monkeypatch,
):
    class FakePathJAB:
        max_depth = 1
        max_children = 1

        def __init__(self):
            self.visible_paths = {
                "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0.1.6",
                "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0.1.7",
                "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0.1.2",
            }

        def wait_context_by_path(
            self,
            path,
            title=None,
            class_name=None,
            name=None,
            role=None,
            require_showing=True,
            require_valid_bounds=True,
            timeout=None,
            scope_hwnd=None,
        ):
            if path in self.visible_paths:
                return {"hwnd": 330038, "class": class_name, "title": "查询结果"}
            return None

    monkeypatch.setattr(
        "tools.receipt_query_pagination_paths.enumerate_visible_table_paths",
        lambda _jab, _window_class: [
            {
                "table_index": 2,
                "path": ("0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0.0.0.0"),
                "hwnd": 330038,
                "row_count": 500,
                "col_count": 32,
            }
        ],
    )

    report = resolve_receipt_pagination_paths_dynamic(
        FakePathJAB(),
        paged_query_config(prefer_configured_paths=False),
    )

    assert report["ok"] is True
    assert report["result_area_prefix"] == (
        "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0"
    )
    assert report["result_table_path"].endswith(".0.0.0")
    assert report["page_label_path"].endswith(".1.6")
    assert report["page_size_text_path"].endswith(".1.7")
    assert report["next_page_button_path"].endswith(".1.2")


def test_resolve_receipt_pagination_paths_by_module_index_uses_shared_receipt_index():
    class FakeModuleIndexJAB:
        def __init__(self):
            self.visible_paths = {
                "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0.0.0.0": "table",
                "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0.1.6": "label",
                "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0.1.7": "text",
                "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0.1.2": "push button",
            }

        def wait_context_by_path(
            self,
            path,
            title=None,
            class_name=None,
            name=None,
            role=None,
            require_showing=True,
            require_valid_bounds=True,
            timeout=None,
            scope_hwnd=None,
        ):
            if self.visible_paths.get(path) == role:
                return {"hwnd": 330038, "class": class_name, "title": "查询结果"}
            return None

    report = resolve_receipt_pagination_paths_by_module_index(
        FakeModuleIndexJAB(),
        paged_query_config(
            prefer_configured_paths=False,
            module_index_paths_enabled=True,
        ),
    )

    assert report["ok"] is True
    assert report["resolution"] == "dynamic_module_index"
    assert report["dynamic_index"] == 5
    assert report["module_prefix"] == "0.0.1.0.0.0.0.5"
    assert report["result_table_path"].endswith(".0.0.0")
    assert report["page_size_text_path"].endswith(".1.7")


def test_read_receipt_tables_accepts_variable_result_columns_when_required_fields_fit():
    class FakeVariableColumnResultJAB:
        def __init__(self):
            self.exact_cols_args = []

        def read_all_table_cells(
            self, max_rows=None, max_cols=None, scope_hwnd=None, exact_cols=None
        ):
            self.exact_cols_args.append(exact_cols)
            return [
                {
                    "table_index": 2,
                    "row_count": 1,
                    "col_count": 32,
                    "rows": [
                        {
                            "row_index": 0,
                            "cells": [
                                "D22026060100026949",
                                "2026-05-27",
                                "收款单",
                                "收款结算",
                                "Copeland Cold Chain LP",
                                "美元",
                                "161713.00",
                            ]
                            + [""] * 25,
                        }
                    ],
                }
            ]

    jab = FakeVariableColumnResultJAB()
    tables = read_receipt_tables(
        jab,
        paged_query_config(),
        max_rows=10,
        max_cols=80,
    )

    assert jab.exact_cols_args == [None]
    assert [table["table_index"] for table in tables] == [2]
    assert tables[0]["col_count"] == 32
