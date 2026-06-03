from tools.receipt_query_fill import (
    ensure_query_window,
    parse_page_label,
    read_receipt_result_pages,
    set_receipt_page_size,
    wait_receipt_result_stable,
)


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

    def read_all_table_cells(self, max_rows=None, max_cols=None):
        self.reads += 1
        if self.reads == 1:
            rows = [
                {"row_index": 0, "cells": ["D1", "2026-03-31"], "selected": False},
                {"row_index": 1, "cells": ["D2", "2026-03-31"], "selected": False},
            ]
        else:
            rows = [{"row_index": 0, "cells": ["D3", "2026-04-01"], "selected": False}]
        return [
            {
                "table_index": 2,
                "row_count": len(rows),
                "col_count": 10,
                "rows": rows,
            }
        ]

    def read_table_summaries(self, min_rows=1, min_cols=None):
        return [{"table_index": 2, "row_count": 3, "col_count": 10}]

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


class FakeFallbackPagedJAB(FakePagedJAB):
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
        return click_mode == "bounds"


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


def test_ensure_query_window_opens_with_f3_when_dialog_missing():
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
        {"open_timeout": 5, "activate_timeout": 3, "open_wait": 0.8},
        {"dialog_title": "查询条件", "dialog_class": "SunAwtDialog"},
    )

    assert ok is True
    assert jab.activated == [("Yonyou UClient", "YonyouUWnd", 3.0)]
    assert jab.keys == [("f3", 0.8)]


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


def test_parse_page_label_reads_total_pages_and_records():
    assert parse_page_label("第1页 共66页 659条记录 每页显示") == {
        "total_pages": 66,
        "total_records": 659,
    }


def test_read_receipt_result_pages_sets_page_size_and_reads_next_page():
    jab = FakePagedJAB()

    tables, report = read_receipt_result_pages(
        jab,
        {
            "pagination": {
                "page_size": 500,
                "page_label_path": "label",
                "page_size_text_path": "size",
                "next_page_button_path": "next",
                "window_class": "SunAwtCanvas",
                "wait_after_page_size": 0,
                "wait_after_next": 0,
            }
        },
        max_rows=500,
        max_cols=40,
    )

    assert jab.texts["size"] == "500"
    assert jab.keys == [("enter", 0.0)]
    assert jab.actions == [("next", "单击", None, 330038)]
    assert report["total_pages"] == 2
    assert report["pager_hwnd"] == 330038
    assert report["pages"][0]["next_page_method"] == "action"
    assert [table["row_count"] for table in tables] == [2, 1]


def test_read_receipt_result_pages_falls_back_to_bounds_click_for_next_page():
    jab = FakeFallbackPagedJAB()

    tables, report = read_receipt_result_pages(
        jab,
        {
            "pagination": {
                "page_size": 500,
                "page_label_path": "label",
                "page_size_text_path": "size",
                "next_page_button_path": "next",
                "window_class": "SunAwtCanvas",
                "wait_after_page_size": 0,
                "wait_after_next": 0,
            }
        },
        max_rows=500,
        max_cols=40,
    )

    assert jab.actions == [
        ("next", "单击", None, 330038),
        ("next", None, "bounds", 330038),
    ]
    assert report["pages"][0]["next_page_ok"] is True
    assert report["pages"][0]["next_page_method"] == "bounds"
    assert [table["row_count"] for table in tables] == [2, 1]


def test_read_receipt_result_pages_blocks_next_page_without_pager_scope():
    jab = FakeNoScopePagedJAB()

    tables, report = read_receipt_result_pages(
        jab,
        {
            "pagination": {
                "page_size": 500,
                "page_label_path": "label",
                "page_size_text_path": "size",
                "next_page_button_path": "next",
                "window_class": "SunAwtCanvas",
                "wait_after_page_size": 0,
                "wait_after_next": 0,
            }
        },
        max_rows=500,
        max_cols=40,
    )

    assert jab.actions == []
    assert report["pager_scope_ok"] is False
    assert report["pages"][0]["next_page_ok"] is False
    assert report["pages"][0]["next_page_method"] == "blocked_no_pager_scope"
    assert [table["row_count"] for table in tables] == [2]


def test_read_receipt_result_pages_applies_stability_waits(monkeypatch):
    waits = []
    monkeypatch.setattr("tools.receipt_query_fill.time.sleep", waits.append)

    read_receipt_result_pages(
        FakePagedJAB(),
        {
            "pagination": {
                "page_size": 500,
                "page_label_path": "label",
                "page_size_text_path": "size",
                "next_page_button_path": "next",
                "window_class": "SunAwtCanvas",
                "wait_before_page_size": 1,
                "wait_after_page_size": 2,
                "wait_before_read": 3,
                "wait_after_page_read": 4,
                "wait_after_next": 5,
            }
        },
        max_rows=500,
        max_cols=40,
    )

    assert 1.0 in waits
    assert waits.count(3.0) == 2
    assert waits.count(4.0) == 2


def test_wait_receipt_result_stable_requires_repeated_label_and_tables(monkeypatch):
    waits = []
    monkeypatch.setattr("tools.receipt_query_fill.time.sleep", waits.append)

    report = wait_receipt_result_stable(
        FakePagedJAB(),
        {
            "result_column_indexes": {"document_no": 0, "document_date": 1},
            "pagination": {
                "page_label_path": "label",
                "window_class": "SunAwtCanvas",
                "stability_timeout": 5,
                "stability_interval": 0.25,
                "stability_required": 2,
            },
        },
    )

    assert report["ok"] is True
    assert report["samples"] == 2
    assert waits == [0.25]


def test_set_receipt_page_size_can_skip_pre_stability(monkeypatch):
    waits = []
    monkeypatch.setattr("tools.receipt_query_fill.time.sleep", waits.append)

    report = set_receipt_page_size(
        FakePagedJAB(),
        {
            "result_column_indexes": {"document_no": 0, "document_date": 1},
            "pagination": {
                "page_size": 500,
                "page_label_path": "label",
                "page_size_text_path": "size",
                "window_class": "SunAwtCanvas",
                "wait_before_page_size_stable": False,
                "wait_before_page_size": 0,
                "wait_after_page_size": 0,
                "stability_timeout": 5,
                "stability_interval": 0.25,
                "stability_required": 2,
            },
        },
    )

    before_stability = report["before_stability"]
    after_stability = report["after_stability"]
    assert isinstance(before_stability, dict)
    assert isinstance(after_stability, dict)
    assert before_stability["ok"] is None
    assert after_stability["ok"] is True
    assert waits == [0.0, 0.25]
