# 生命周期：持久维护
# 覆盖的业务场景：收款单查询填充测试的共享 Fake JAB 替身、配置工厂、路径常量与 import
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake JAB 替身）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_query_*.py

from datetime import date
from decimal import Decimal
from typing import Any, cast

from openpyxl import Workbook, load_workbook
import pytest

from core.receipt_matching import (
    format_receipt_duplicate_reason,
    format_receipt_name_amount_mismatch_reason,
)
from core.receipt_models import (
    ReceiptExcelRow,
    ReceiptMatchIssue,
    ReceiptNCExtractIssue,
    ReceiptNCIndexedRow,
)
from core.receipt_nc_extract import ReceiptNCResultExtractor
from tools.receipt_query_report import (
    build_dry_run_match_report,
    build_dry_run_match_report_from_preview,
)
from tools.receipt_query_pagination import (
    parse_page_label,
    set_receipt_page_size,
    wait_receipt_result_stable,
)
from tools.receipt_query_pagination_paths import (
    resolve_receipt_pagination_paths_by_module_index,
    resolve_receipt_pagination_paths_dynamic,
    infer_result_area_prefix_from_table_path,
)
from tools.receipt_query_page_reader import read_receipt_result_pages
from tools.receipt_query_result_tables import read_receipt_tables
from tools.receipt_query_fill import (
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
                "nc_done_column": "是否NC已做过",
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
                "only_blank_status": True,
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


__all__ = [
    'Any',
    'Decimal',
    'FakeFailingNextPageJAB',
    'FakeGuardJAB',
    'FakeJAB',
    'FakeNoScopePagedJAB',
    'FakePagedJAB',
    'FakeReceiptQueryJAB',
    'QUERY_DATE_FROM_PATH',
    'QUERY_DATE_TO_PATH',
    'QUERY_FINANCE_ORG_PATH',
    'QUERY_PREFIX',
    'ReceiptExcelRow',
    'ReceiptMatchIssue',
    'ReceiptNCExtractIssue',
    'ReceiptNCIndexedRow',
    'ReceiptNCResultExtractor',
    'ReceiptPageGuardError',
    'Workbook',
    'build_dry_run_match_report',
    'build_dry_run_match_report_from_preview',
    'cast',
    'date',
    'ensure_query_window',
    'fill_receipt_query',
    'format_receipt_duplicate_reason',
    'format_receipt_name_amount_mismatch_reason',
    'guard_receipt_parent_page',
    'guard_receipt_result_tables',
    'infer_result_area_prefix_from_table_path',
    'load_workbook',
    'paged_query_config',
    'parse_page_label',
    'pytest',
    'read_receipt_result_pages',
    'read_receipt_tables',
    'receipt_config',
    'resolve_receipt_pagination_paths_by_module_index',
    'resolve_receipt_pagination_paths_dynamic',
    'set_receipt_page_size',
    'wait_after_query_confirm',
    'wait_receipt_result_stable',
]
