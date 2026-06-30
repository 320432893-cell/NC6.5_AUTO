from datetime import date
from decimal import Decimal
from typing import Any, cast

from openpyxl import Workbook
import pytest

from core.receipt_matching import format_receipt_name_amount_mismatch_reason
from core.receipt_models import (
    ReceiptExcelRow,
    ReceiptMatchIssue,
    ReceiptNCExtractIssue,
    ReceiptNCIndexedRow,
)
from core.receipt_nc_extract import ReceiptNCResultExtractor
from core.receipt_query_report import (
    build_dry_run_match_report,
    build_dry_run_match_report_from_preview,
)
from core.receipt_query_pagination import (
    parse_page_label,
    set_receipt_page_size,
    wait_receipt_result_stable,
)
from core.receipt_query_pagination_paths import (
    resolve_receipt_pagination_paths_by_module_index,
    resolve_receipt_pagination_paths_dynamic,
    infer_result_area_prefix_from_table_path,
)
from core.receipt_query_page_reader import read_receipt_result_pages
from core.receipt_query_match_reader import read_receipt_result_pages_incremental
from core.receipt_query_result_tables import read_receipt_tables
from core.receipt_query_fill import (
    ReceiptPageGuardError,
    ensure_query_window,
    fill_receipt_query,
    guard_receipt_parent_page,
    guard_receipt_result_tables,
    wait_after_query_confirm,
)

QUERY_PREFIX = "0.0.1.0.1.0.0.1.0.0.0.0.0.1.0.1"
QUERY_FINANCE_ORG_PATH = f"{QUERY_PREFIX}.1.2.0.0.0.0"
QUERY_DATE_FROM_PATH = f"{QUERY_PREFIX}.3.2.0.0.0.0"
QUERY_DATE_TO_PATH = f"{QUERY_PREFIX}.3.2.1.0.0.0"


class FakeJAB:
    def __init__(self, existing=False):
        self.existing = existing
        self.keys = []
        self.activated = []
        self.wait_calls = 0

    def wait_window_by_title(
        self,
        title,
        class_name=None,
        timeout=None,
        include_children=False,
        visible_only=True,
        interval=0.2,
    ):
        self.wait_calls += 1
        if self.existing or self.keys:
            return 100
        return None

    def activate_window_by_title(self, title, class_name=None, timeout=None):
        self.activated.append((title, class_name, timeout))
        return True

    def press_key(self, key, wait=None):
        self.keys.append((key, wait))


class FakePagedJAB:
    def __init__(self):
        self.texts = {}
        self.actions = []
        self.keys = []
        self.reads = 0
        self.table_scopes = []

    def get_text_by_path(
        self,
        path,
        title=None,
        class_name=None,
        scope_hwnd=None,
        name=None,
        role=None,
        timeout=None,
        require_showing=True,
        require_valid_bounds=False,
    ):
        if path in self.texts:
            return self.texts[path]
        if path == "size" or path.endswith("_size"):
            return "50"
        return self.texts.get(path, "第1页 共2页 3条记录 每页显示")

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
        return {"hwnd": 330038}

    def set_text_by_path(
        self,
        path,
        text,
        title=None,
        class_name=None,
        scope_hwnd=None,
        name=None,
        role=None,
        guard_path=None,
        guard_name=None,
        guard_role=None,
        wait=None,
        timeout=None,
        require_showing=True,
        require_valid_bounds=True,
    ):
        self.texts[path] = text
        return True

    def press_key(self, key, wait=None):
        self.keys.append((key, wait))

    def read_all_table_cells(
        self, max_rows=None, max_cols=None, scope_hwnd=None, exact_cols=None
    ):
        self.reads += 1
        self.table_scopes.append(scope_hwnd)
        if self.reads == 1:
            rows = [
                {
                    "row_index": 0,
                    "cells": [
                        "D100000001",
                        "2026-03-31",
                        "收款单",
                        "",
                        "客户A",
                        "",
                        "10.00",
                    ]
                    + [""] * 34,
                    "selected": False,
                },
                {
                    "row_index": 1,
                    "cells": [
                        "D100000002",
                        "2026-03-31",
                        "收款单",
                        "",
                        "客户B",
                        "",
                        "20.00",
                    ]
                    + [""] * 34,
                    "selected": False,
                },
            ]
        else:
            rows = [
                {
                    "row_index": 0,
                    "cells": [
                        "D100000003",
                        "2026-04-01",
                        "收款单",
                        "",
                        "客户C",
                        "",
                        "30.00",
                    ]
                    + [""] * 34,
                    "selected": False,
                }
            ]
        return [
            {
                "table_index": 2,
                "row_count": len(rows),
                "col_count": 41,
                "rows": rows,
            }
        ]

    def read_table_summaries(
        self, min_rows=1, min_cols=None, scope_hwnd=None, exact_cols=None
    ):
        self.table_scopes.append(scope_hwnd)
        return [{"table_index": 2, "row_count": 3, "col_count": 41}]

    def find_context_by_path_once(self, *args, **kwargs):
        return None, None, [], None

    def release_contexts(self, vm_id, owned_contexts):
        pass

    def do_action_by_path(
        self,
        path,
        title=None,
        class_name=None,
        scope_hwnd=None,
        name=None,
        role=None,
        action_name=None,
        click_mode=None,
        wait=None,
        timeout=None,
        require_showing=True,
        require_valid_bounds=True,
    ):
        self.actions.append((path, action_name, click_mode, scope_hwnd))
        return True


class FakeFailingNextPageJAB(FakePagedJAB):
    def do_action_by_path(
        self,
        path,
        title=None,
        class_name=None,
        scope_hwnd=None,
        name=None,
        role=None,
        action_name=None,
        click_mode=None,
        wait=None,
        timeout=None,
        require_showing=True,
        require_valid_bounds=True,
    ):
        self.actions.append((path, action_name, click_mode, scope_hwnd))
        return False


class FakeNoScopePagedJAB(FakePagedJAB):
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
        return None


class FakeGuardJAB:
    def __init__(self, found):
        self.found = found
        self.released = []
        self.started = False

    def ensure_started(self):
        self.started = True

    def find_context(self, *args, **kwargs):
        if self.found:
            return 11, 22, [11]
        return None, None, []

    def release_contexts(self, vm_id, contexts):
        self.released.append((vm_id, contexts))


def paged_query_config(**pagination_overrides):
    pagination = {
        "page_size": 500,
        "page_label_path": "label",
        "page_size_text_path": "size",
        "next_page_button_path": "next",
        "window_class": "SunAwtCanvas",
        "wait_after_page_size": 0,
        "wait_after_next": 0,
    }
    pagination.update(pagination_overrides)
    return {
        "result_column_indexes": {
            "document_no": 0,
            "document_date": 1,
            "customer": 4,
            "original_amount": 6,
            "payer_name": 4,
        },
        "result_table_cols": 41,
        "pagination": pagination,
    }


class FakeReceiptQueryJAB:
    def __init__(self, config, path_ok=True):
        self.config = config
        self.path_ok = path_ok
        self.actions = []
        self.set_texts = []
        self.near_label_texts = []
        self.keys = []
        self.closed = False

    def ensure_started(self):
        return None

    def find_context(self, *args, **kwargs):
        return 11, 22, [11]

    def release_contexts(self, vm_id, contexts):
        return None

    def wait_window_by_title(self, *args, **kwargs):
        return 100

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
        expected = {
            QUERY_FINANCE_ORG_PATH,
            QUERY_DATE_FROM_PATH,
            QUERY_DATE_TO_PATH,
        }
        if self.path_ok and path in expected:
            return {"path": path}
        return None

    def set_text_by_path(
        self,
        path,
        text,
        title=None,
        class_name=None,
        scope_hwnd=None,
        name=None,
        role=None,
        guard_path=None,
        guard_name=None,
        guard_role=None,
        wait=None,
        timeout=None,
        require_showing=True,
        require_valid_bounds=True,
    ):
        self.set_texts.append(
            {
                "path": path,
                "text": text,
                "title": title,
                "class_name": class_name,
                "role": role,
                "wait": wait,
                "timeout": timeout,
                "require_showing": require_showing,
            }
        )
        if path == QUERY_FINANCE_ORG_PATH and not self.path_ok:
            return False
        return True

    def set_text_near_label(
        self,
        label,
        text,
        title=None,
        class_name=None,
        wait=None,
        timeout=None,
        require_showing=True,
    ):
        self.near_label_texts.append(
            {
                "label": label,
                "text": text,
                "title": title,
                "class_name": class_name,
                "timeout": timeout,
                "require_showing": require_showing,
            }
        )
        return True

    def do_action_by_path(
        self,
        path,
        title=None,
        class_name=None,
        scope_hwnd=None,
        name=None,
        role=None,
        action_name=None,
        click_mode=None,
        wait=None,
        timeout=None,
        require_showing=True,
        require_valid_bounds=True,
    ):
        self.actions.append(
            {
                "path": path,
                "role": role,
                "click_mode": click_mode,
                "wait": wait,
                "timeout": timeout,
            }
        )
        return True

    def close(self):
        self.closed = True


def receipt_config(path):
    return {
        "receipt_entry": {
            "state_label": "收款单录入",
            "excel": {
                "path": str(path),
                "sheet_name": "💸Payments来款通知",
                "header_row": 1,
                "start_row": 2,
                "result_sheet_name": "收款单自动化结果",
                "start_date": "2026-01-01",
                "date_column": "到款日期",
                "payer_name_column": "🟪银行来款名",
                "raw_amount_column": "🟪原始金额",
                "bank_column": "银行",
                "currency_column": "币种",
                "customer_code_column": "客户编码",
                "fee_column": "手续费",
                "organization_column": "主体名称",
            },
            "validation_policy": {
                "mode": "strict",
                "skip_invalid_rows": False,
            },
            "query": {
                "date_from": "2026-01-01",
                "date_to": "{today}",
                "open_key": "f3",
                "open_timeout": 5,
                "result_wait_timeout": 0.5,
                "result_wait_interval": 0.05,
                "result_wait_fallback": 0.0,
                "jab": {
                    "dialog_title": "查询条件",
                    "dialog_class": "SunAwtDialog",
                    "confirm_button_path": "confirm",
                    "confirm_timeout": 1.0,
                    "confirm_wait": 0.0,
                    "fields": {
                        "finance_org": {
                            "label": "收款财务组织",
                            "operator": "等于",
                            "text_path": QUERY_FINANCE_ORG_PATH,
                            "path_timeout": 0.5,
                            "timeout": 2.0,
                        },
                        "document_date": {
                            "label": "单据日期",
                            "operator": "介于",
                            "from_text_path": QUERY_DATE_FROM_PATH,
                            "to_text_path": QUERY_DATE_TO_PATH,
                        },
                    },
                },
                "result_column_indexes": {
                    "document_no": 0,
                    "document_date": 1,
                    "customer": 2,
                    "original_amount": 7,
                    "payer_name": 2,
                },
                "pagination": {
                    "page_label_path": "label",
                    "window_class": "SunAwtCanvas",
                },
            },
            "candidate_check": {
                "recent_months": 2,
                "from_date": None,
            },
            "finance_organizations": [
                {
                    "code": "A001",
                    "name": "上海移为通信技术股份有限公司",
                    "short_name": "移为",
                }
            ],
            "accounts": [
                {
                    "organization_code": "A001",
                    "organization_short_name": "移为",
                    "account_label": "Paypal",
                    "account_no": "paypal",
                }
            ],
        }
    }


def test_ensure_query_window_opens_with_f3_without_fixed_sleep():
    jab = FakeJAB(existing=False)
    ok = ensure_query_window(
        jab,
        {
            "jab_batch": {
                "open_query": {
                    "main_title": "Yonyou UClient",
                    "main_class": "YonyouUWnd",
                    "key": "f3",
                }
            }
        },
        {
            "open_timeout": 5,
            "activate_timeout": 3,
            "existing_dialog_timeout": 0.1,
            "open_wait": 0.0,
        },
        {"dialog_title": "查询条件", "dialog_class": "SunAwtDialog"},
    )

    assert ok is True
    assert jab.activated == [("Yonyou UClient", "YonyouUWnd", 3.0)]
    assert jab.keys == [("f3", 0.0)]


def test_ensure_query_window_reuses_existing_dialog():
    jab = FakeJAB(existing=True)

    ok = ensure_query_window(
        jab,
        {},
        {},
        {"dialog_title": "查询条件", "dialog_class": "SunAwtDialog"},
    )

    assert ok is True
    assert jab.activated == []
    assert jab.keys == []


def test_fill_receipt_query_sets_finance_org_by_path(monkeypatch):
    instances = []

    def make_jab(config):
        instance = FakeReceiptQueryJAB(config)
        instances.append(instance)
        return instance

    monkeypatch.setattr("core.receipt_query_fill.JABOperator", make_jab)

    result = fill_receipt_query(
        receipt_config("unused.xlsx"),
        org_code="A003",
        date_from="2026-03-31",
        date_to="2026-05-31",
        confirm=False,
    )

    jab = instances[0]
    assert result["organization_code"] == "A003"
    assert jab.actions == []
    assert jab.near_label_texts == []
    assert jab.set_texts[0] == {
        "path": QUERY_FINANCE_ORG_PATH,
        "text": "A003",
        "title": "查询条件",
        "class_name": "SunAwtDialog",
        "role": "text",
        "wait": 0.0,
        "timeout": 2,
        "require_showing": True,
    }
    assert [(item["path"], item["text"]) for item in jab.set_texts] == [
        (QUERY_FINANCE_ORG_PATH, "A003"),
        (QUERY_DATE_FROM_PATH, "2026-03-31"),
        (QUERY_DATE_TO_PATH, "2026-05-31"),
    ]
    assert jab.closed is True


def test_fill_receipt_query_confirms_without_fixed_wait(monkeypatch):
    instances = []

    def make_jab(config):
        instance = FakeReceiptQueryJAB(config)
        instances.append(instance)
        return instance

    monkeypatch.setattr("core.receipt_query_fill.JABOperator", make_jab)

    result = fill_receipt_query(
        receipt_config("unused.xlsx"),
        org_code="A003",
        date_from="2026-03-31",
        date_to="2026-05-31",
        confirm=True,
    )

    jab = instances[0]
    assert result["organization_code"] == "A003"
    assert jab.actions == [
        {
            "path": "confirm",
            "role": "push button",
            "click_mode": None,
            "wait": 0.0,
            "timeout": 1.0,
        }
    ]
    assert jab.closed is True


def test_wait_after_query_confirm_returns_when_result_table_path_is_ready(monkeypatch):
    waits = []
    monkeypatch.setattr("core.receipt_query_fill.time.sleep", waits.append)
    jab = FakePagedJAB()

    report = wait_after_query_confirm(
        jab,
        {
            "result_wait_timeout": 0.5,
            "result_wait_interval": 0.1,
            "pagination": {
                "module_index_paths_enabled": True,
                "window_class": "SunAwtCanvas",
                "page_label_path": "label",
                "page_size_text_path": "size",
                "next_page_button_path": "next",
            },
        },
    )

    assert report["ok"] is True
    assert report["method"] == "result_table_path"
    assert report["result_table_path"]
    assert waits == []


def test_fill_receipt_query_fails_without_dynamic_or_semantic_path(monkeypatch):
    instances = []

    def make_jab(config):
        instance = FakeReceiptQueryJAB(config, path_ok=False)
        instances.append(instance)
        return instance

    monkeypatch.setattr("core.receipt_query_fill.JABOperator", make_jab)

    with pytest.raises(RuntimeError, match="查询条件动态 path 定位失败"):
        fill_receipt_query(
            receipt_config("unused.xlsx"),
            org_code="A003",
            date_from="2026-03-31",
            date_to="2026-05-31",
            confirm=False,
        )

    jab = instances[0]
    assert jab.near_label_texts == []
    assert jab.set_texts == []
    assert jab.closed is True


def test_parse_page_label_reads_total_pages_and_records():
    assert parse_page_label("第1页 共66页 659条记录 每页显示") == {
        "total_pages": 66,
        "total_records": 659,
    }


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


def test_read_receipt_result_pages_sets_page_size_and_reads_next_page():
    jab = FakePagedJAB()
    jab.texts["label"] = "第1页 共2页 501条记录 每页显示"

    tables, report = read_receipt_result_pages(
        jab,
        paged_query_config(prefer_configured_paths=False),
        max_rows=500,
        max_cols=40,
    )

    assert jab.texts["size"] == "500"
    assert jab.keys == [("enter", 0.0)]
    assert jab.actions == [("next", "单击", None, 330038)]
    assert report["total_pages"] == 2
    assert report["pager_hwnd"] == 330038
    assert report["pages"][0]["next_page_method"] == "action"
    assert all(scope == 330038 for scope in jab.table_scopes)
    assert [table["row_count"] for table in tables] == [2, 1]


def test_read_receipt_result_pages_skips_page_size_change_when_already_target():
    jab = FakePagedJAB()
    jab.texts["size"] = "500"

    tables, report = read_receipt_result_pages(
        jab,
        paged_query_config(prefer_configured_paths=False),
        max_rows=500,
        max_cols=40,
    )

    assert jab.keys == []
    assert report["page_size_ok"] is True
    assert report["page_size_changed"] is False
    assert report["before_page_size_text"] == "500"
    assert report["after_page_size_text"] == "500"
    assert report["after_stability"] == {
        "ok": None,
        "skipped": True,
        "reason": "page_size_already_target",
        "label": report["before_label"],
        "tables": [],
    }
    assert report["after_stability_seconds"] == 0.0
    assert report["pagination_plan_reason"] == "total_records_within_page_size"
    assert jab.actions == []
    assert [table["row_count"] for table in tables] == [2]


def test_incremental_post_save_query_refreshes_after_page_size(monkeypatch):
    waits = []
    monkeypatch.setattr("core.receipt_query_pagination.time.sleep", waits.append)
    jab = FakePagedJAB()
    jab.texts["label"] = "第1页 共1页 16条记录 每页显示"

    tables, report, snapshot = read_receipt_result_pages_incremental(
        jab,
        paged_query_config(
            prefer_configured_paths=False,
            wait_after_page_size=0,
            wait_after_refresh=0,
            stability_timeout=1,
            stability_interval=0.1,
        ),
        ReceiptNCResultExtractor(
            {
                "receipt_entry": {
                    "query": {
                        "result_column_indexes": {
                            "document_no": 0,
                            "document_date": 1,
                            "customer": 4,
                            "original_amount": 6,
                            "payer_name": 4,
                        }
                    }
                }
            }
        ),
        [],
        max_rows=500,
        max_cols=40,
    )

    assert tables
    assert snapshot["matched"] == {}
    assert jab.keys[:2] == [("enter", 0.0), ("f5", 0.0)]
    assert report["post_page_size_refresh"]["enabled"] is True
    assert report["post_page_size_refresh"]["key"] == "f5"


def test_read_receipt_result_pages_uses_dynamic_pagination_paths(monkeypatch):
    jab = FakePagedJAB()

    def fake_dynamic(_jab, _query_cfg):
        return {
            "ok": True,
            "resolution": "dynamic",
            "window_class": "SunAwtCanvas",
            "pager_hwnd": 330038,
            "result_table_path": "0.9.0.0.0",
            "result_area_prefix": "0.9",
            "page_label_path": "dynamic_label",
            "page_size_text_path": "dynamic_size",
            "next_page_button_path": "dynamic_next",
        }

    monkeypatch.setattr(
        "core.receipt_query_pagination_paths.resolve_receipt_pagination_paths_dynamic",
        fake_dynamic,
    )
    jab.texts["dynamic_label"] = "第1页 共2页 501条记录 每页显示"

    tables, report = read_receipt_result_pages(
        jab,
        paged_query_config(prefer_configured_paths=False),
        max_rows=500,
        max_cols=40,
    )

    assert jab.texts["dynamic_size"] == "500"
    assert jab.actions == [("dynamic_next", "单击", None, 330038)]
    assert report["pager_resolution"] == "dynamic"
    assert report["page_label_path"] == "dynamic_label"
    assert report["next_page_button_path"] == "dynamic_next"
    assert [table["row_count"] for table in tables] == [2, 1]


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
        "core.receipt_query_pagination_paths.enumerate_visible_table_paths",
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


def test_read_receipt_result_pages_does_not_fall_back_to_bounds_click_for_next_page():
    jab = FakeFailingNextPageJAB()
    jab.texts["label"] = "第1页 共2页 501条记录 每页显示"

    tables, report = read_receipt_result_pages(
        jab,
        paged_query_config(),
        max_rows=500,
        max_cols=40,
    )

    assert jab.actions == [("next", "单击", None, 330038)]
    assert report["pages"][0]["next_page_ok"] is False
    assert report["pages"][0]["next_page_method"] == "failed"
    assert [table["row_count"] for table in tables] == [2]


def test_read_receipt_result_pages_blocks_next_page_without_pager_scope():
    jab = FakeNoScopePagedJAB()

    tables, report = read_receipt_result_pages(
        jab,
        paged_query_config(),
        max_rows=500,
        max_cols=40,
    )

    assert jab.actions == []
    assert report["pager_scope_ok"] is False
    assert "next_page_ok" not in report["pages"][0]
    assert [table["row_count"] for table in tables] == [2]


def test_read_receipt_result_pages_applies_stability_waits(monkeypatch):
    waits = []
    monkeypatch.setattr("core.receipt_query_pagination.time.sleep", waits.append)
    jab = FakePagedJAB()
    jab.texts["label"] = "第1页 共2页 501条记录 每页显示"

    read_receipt_result_pages(
        jab,
        paged_query_config(
            wait_before_page_size=1,
            wait_after_page_size=2,
            wait_before_read=3,
            wait_after_page_read=4,
            wait_after_next=5,
        ),
        max_rows=500,
        max_cols=40,
    )

    assert 1.0 in waits
    assert waits.count(3.0) == 2
    assert waits.count(4.0) == 2


def test_wait_receipt_result_stable_requires_repeated_label_and_tables(monkeypatch):
    waits = []
    monkeypatch.setattr("core.receipt_query_pagination.time.sleep", waits.append)

    report = wait_receipt_result_stable(
        FakePagedJAB(),
        paged_query_config(
            page_label_path="label",
            stability_timeout=5,
            stability_interval=0.25,
            stability_required=2,
        ),
    )

    assert report["ok"] is True
    assert report["samples"] == 2
    assert waits == [0.25]


def test_set_receipt_page_size_can_skip_pre_stability(monkeypatch):
    waits = []
    monkeypatch.setattr("core.receipt_query_pagination.time.sleep", waits.append)

    report = set_receipt_page_size(
        FakePagedJAB(),
        paged_query_config(
            wait_before_page_size_stable=False,
            wait_before_page_size=0,
            wait_after_page_size=0,
            stability_timeout=5,
            stability_interval=0.25,
            stability_required=2,
        ),
    )

    before_stability = report["before_stability"]
    after_stability = report["after_stability"]
    assert isinstance(before_stability, dict)
    assert isinstance(after_stability, dict)
    assert before_stability["ok"] is None
    assert before_stability["skipped"] is True
    assert before_stability["reason"] == "pre_stability_disabled"
    assert before_stability["tables"] == []
    assert after_stability["ok"] is True
    assert waits == [0.0]


def test_resolve_receipt_pagination_paths_uses_cached_report(monkeypatch):
    calls = []

    def fail_dynamic(_jab, _query_cfg):
        calls.append("dynamic")
        return {"ok": False}

    monkeypatch.setattr(
        "core.receipt_query_pagination_paths.resolve_receipt_pagination_paths_dynamic",
        fail_dynamic,
    )
    jab = FakePagedJAB()
    cast(Any, jab)._receipt_pagination_paths_cache = {
        "window_class": "SunAwtCanvas",
        "pager_hwnd": 330038,
        "page_label_path": "label",
        "page_size_text_path": "size",
        "next_page_button_path": "next",
    }

    report = set_receipt_page_size(jab, paged_query_config())

    assert report["pager_resolution"] == "cached_trusted"
    assert calls == []


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
    )

    assert "write_back" not in report
    assert report["match_summary"]["matched_rows"] == [2]
    assert report["match_summary"]["not_found_rows"] == [3]
    assert report["match_summary"]["duplicate_rows"] == [4]
    assert report["match_summary"]["exception_rows"] == [4, 5]
    assert report["match_summary"]["skipped_duplicate_rows"] == [4]
    assert report["match_summary"]["planned"] == 4
    assert report["variants"][0]["issue_samples"][2][
        "reason"
    ] == format_receipt_name_amount_mismatch_reason(
        excel_amount=Decimal("400.00"),
        excel_name="Christoff Pretorius",
        nc_amounts=["450.00"],
    )


def test_dry_run_match_report_ignores_legacy_status_column(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行", "是否NC已做过"])
    ws.append([date(2026, 5, 1), "MATCHED INC", 100, "Paypal", "未做过"])
    wb.save(path)
    wb.close()
    config = receipt_config(path)
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
    )

    assert report["org_candidates"] == 1
    assert "write_back" not in report
    assert report["match_summary"]["matched_rows"] == [2]


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
    assert report["match_summary"]["matched_rows"] == [2]


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
