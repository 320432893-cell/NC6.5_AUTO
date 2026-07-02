import ctypes
from decimal import Decimal
from typing import Any, cast

from core.jab_helpers import (
    context_info_has_valid_bounds,
    context_info_is_showing,
    normalize_amount,
    normalize_text,
    parse_context_path,
    text_matches,
)
from core import jab_window
from core import jab_popup
from core import jab_operator
from core.jab_context_tree import matches_control, release_contexts
from core.jab_operator import JABOperator
from core.jab_table_reader import read_table_cells_from_context
from core.jab_probe import AccessibleContextInfo


def make_info(states, x=1, y=1, width=10, height=10):
    info = AccessibleContextInfo()
    info.states_en_US = states
    info.x = x
    info.y = y
    info.width = width
    info.height = height
    return info


def test_context_info_is_showing_requires_showing_state():
    assert context_info_is_showing(make_info("enabled,visible,showing"))
    assert not context_info_is_showing(make_info("enabled,visible"))
    assert JABOperator.context_info_is_showing(make_info("enabled,visible,showing"))
    assert not JABOperator.context_info_is_showing(make_info("enabled,visible"))


def test_context_info_has_valid_bounds_rejects_mirror_bounds():
    assert context_info_has_valid_bounds(make_info("enabled", 406, 548, 70, 15))
    assert not context_info_has_valid_bounds(make_info("enabled", -1, -1, -1, -1))
    assert not context_info_has_valid_bounds(make_info("enabled", 0, 0, 0, 15))
    assert JABOperator.context_info_has_valid_bounds(
        make_info("enabled", 406, 548, 70, 15)
    )
    assert not JABOperator.context_info_has_valid_bounds(
        make_info("enabled", -1, -1, -1, -1)
    )
    assert not JABOperator.context_info_has_valid_bounds(
        make_info("enabled", 0, 0, 0, 15)
    )


def test_jab_pure_helpers_keep_operator_compatibility():
    operator = JABOperator({"jab": {}})

    assert normalize_amount("1,234.5") == Decimal("1234.50")
    assert operator.normalize_amount("1,234.5") == Decimal("1234.50")
    assert normalize_amount("bad") is None
    assert normalize_text(" 上海 \n 移为 ") == "上海移为"
    assert operator.normalize_text(" 上海 \n 移为 ") == "上海移为"
    assert text_matches("上海移为通信", "移为", "contains")
    assert operator.text_matches("上海移为通信", "移为", "contains")


def test_parse_context_path_keeps_operator_compatibility():
    assert parse_context_path("0.1.2") == [0, 1, 2]
    assert JABOperator.parse_context_path("0.1.2") == [0, 1, 2]

    try:
        parse_context_path("1.2")
    except ValueError as exc:
        assert "must start with 0" in str(exc)
    else:
        raise AssertionError("parse_context_path should reject paths not rooted at 0")


def test_window_helpers_return_safe_defaults_on_non_windows():
    assert jab_window.hide_blank_awt_windows(True) == []
    assert not jab_window.close_window_by_title("不存在")
    assert not jab_window.activate_window_by_title("不存在", timeout=0.01)
    assert jab_window.get_foreground_window_info() is None
    assert not jab_window.foreground_window_matches("不存在")
    assert jab_window.wait_window_by_title("不存在", timeout=0.01) is None
    assert not jab_window.window_exists("不存在")
    assert jab_window.find_window_handle("不存在") is None


def test_operator_window_methods_delegate_to_window_helpers(monkeypatch):
    operator = JABOperator({"jab": {"menu_wait": 0.3, "search_timeout": 0.4}})
    calls = []

    def fake_close(
        title, class_name=None, wait=None, menu_wait=0.5, clear_table_cache=None
    ):
        calls.append((title, class_name, wait, menu_wait, bool(clear_table_cache)))
        return True

    monkeypatch.setattr(jab_window, "close_window_by_title", fake_close)

    assert operator.close_window_by_title("单据", class_name="SunAwtDialog", wait=0.1)
    assert calls == [("单据", "SunAwtDialog", 0.1, 0.3, True)]


def test_operator_lifecycle_does_not_auto_cleanup_awt(monkeypatch):
    operator = JABOperator({"jab": {"startup_wait": 0}})
    calls = []

    class FakeDll:
        def initializeAccessBridge(self):
            return True

    monkeypatch.setattr("core.jab_operator.os.name", "nt")
    monkeypatch.setattr(
        "core.jab_operator.ensure_jab_setup_once",
        lambda: {"ok": True, "skipped": True, "jabswitch": "fake-jabswitch.exe"},
    )
    monkeypatch.setattr(
        "core.jab_operator.load_access_bridge",
        lambda _path: (FakeDll(), "fake.dll"),
    )
    monkeypatch.setattr("core.jab_operator.configure_jab", lambda _dll: None)
    monkeypatch.setattr(
        jab_window,
        "hide_blank_awt_windows",
        lambda enabled: calls.append(enabled) or [],
    )

    operator.ensure_started()
    operator.close()

    assert calls == []


def test_operator_explicit_awt_cleanup_still_available(monkeypatch):
    operator = JABOperator({"jab": {"hide_blank_awt_windows": True}})
    calls = []
    monkeypatch.setattr(
        jab_window,
        "hide_blank_awt_windows",
        lambda enabled: calls.append(enabled) or [{"hwnd": 1}],
    )

    assert operator.hide_blank_awt_windows() == [{"hwnd": 1}]
    assert calls == [True]


def test_generate_front_prefers_tracked_popup_and_cleans_it(monkeypatch):
    operator = JABOperator({"jab": {"menu_wait": 0.3, "search_timeout": 0.4}})
    calls = []

    monkeypatch.setattr(operator, "ensure_started", lambda: calls.append("started"))
    monkeypatch.setattr(jab_operator, "check_abort", lambda: None)
    monkeypatch.setattr(
        operator,
        "click_control",
        lambda name, **kwargs: calls.append(("click_control", name, kwargs)) or True,
    )
    monkeypatch.setattr(jab_popup, "collect_visible_popup_windows", lambda jab: ["before"])
    monkeypatch.setattr(
        jab_popup,
        "wait_for_new_popup_with_menu_item",
        lambda jab, before, menu_name, timeout, interval: {
            "ok": True,
            "popup": {"hwnd": 99},
            "windows": ["after"],
        },
    )
    monkeypatch.setattr(
        jab_popup,
        "click_menu_item_in_popup",
        lambda jab, windows, menu_name, popup_hwnd=None: calls.append(
            ("popup_click", windows, menu_name, popup_hwnd)
        )
        or {"ok": True},
    )
    monkeypatch.setattr(
        jab_popup,
        "close_popup_hwnd",
        lambda hwnd: calls.append(("popup_cleanup", hwnd)) or {"ok": True},
    )

    assert operator.do_generate_front()

    assert calls == [
        "started",
        ("click_control", "生成", {"roles": ("push button",), "timeout": 0.4}),
        ("popup_click", ["after"], "前台生成", 99),
        ("popup_cleanup", 99),
    ]


def test_generate_front_falls_back_when_tracked_popup_missing(monkeypatch):
    operator = JABOperator({"jab": {"menu_wait": 0.3, "search_timeout": 0.4}})
    clicked = []

    monkeypatch.setattr(operator, "ensure_started", lambda: None)
    monkeypatch.setattr(jab_operator, "check_abort", lambda: None)

    def fake_click(name, **kwargs):
        clicked.append((name, kwargs))
        return True

    monkeypatch.setattr(operator, "click_control", fake_click)
    monkeypatch.setattr(jab_popup, "collect_visible_popup_windows", lambda jab: [])
    monkeypatch.setattr(
        jab_popup,
        "wait_for_new_popup_with_menu_item",
        lambda jab, before, menu_name, timeout, interval: {"ok": False},
    )

    assert operator.do_generate_front()
    assert clicked == [
        ("生成", {"roles": ("push button",), "timeout": 0.4}),
        ("前台生成", {"roles": ("menu item",), "timeout": 0.4}),
    ]


def test_table_reader_keeps_row_shape_and_selection(monkeypatch):
    class TableInfo:
        rowCount = 2
        columnCount = 3

    operator = JABOperator({"jab": {}})
    values = {
        (0, 0): ("客户A", False),
        (0, 1): ("100.00", True),
        (1, 0): ("客户B", False),
        (1, 1): ("200.00", False),
    }

    monkeypatch.setattr(
        "core.jab_table_reader.get_table_cell_text_and_selection",
        lambda _jab, _vm_id, _context, row, col: values.get((row, col), ("", False)),
    )
    monkeypatch.setattr("core.jab_table_reader.check_abort", lambda: None)

    table = read_table_cells_from_context(
        operator,
        4,
        "table-context",
        100,
        TableInfo(),
        {"title": "收款单", "class_name": "SunAwtDialog", "visible": True},
        max_cols=2,
    )

    assert table == {
        "table_index": 4,
        "window_title": "收款单",
        "window_class": "SunAwtDialog",
        "window_visible": True,
        "row_count": 2,
        "col_count": 3,
        "rows": [
            {"row_index": 0, "cells": ["客户A", "100.00"], "selected": True},
            {"row_index": 1, "cells": ["客户B", "200.00"], "selected": False},
        ],
    }


def test_context_tree_matches_control_and_releases_in_reverse_order():
    assert matches_control(
        "保存(Ctrl+S)",
        "",
        "push button",
        "enabled,visible,showing",
        "保存(Ctrl+S)",
        {"push button"},
        True,
    )
    assert not matches_control(
        "保存(Ctrl+S)",
        "",
        "push button",
        "enabled,visible",
        "保存(Ctrl+S)",
        {"push button"},
        True,
    )

    class Dll:
        def __init__(self):
            self.released = []

        def releaseJavaObject(self, vm_id, context):
            self.released.append((vm_id, context))

    operator = JABOperator({"jab": {}})
    operator.dll = cast(Any, Dll())
    release_contexts(operator, 100, [1, 2, 3])

    assert operator.dll.released == [(100, 3), (100, 2), (100, 1)]


def test_find_context_by_path_rejects_visible_non_showing_mirror(monkeypatch):
    operator = JABOperator({"jab": {}})
    operator.dll = cast(
        Any,
        FakeDll(make_info("enabled,focusable,visible", -1, -1, -1, -1)),
    )

    monkeypatch.setattr(
        "core.jab_path_ops.enum_windows",
        lambda include_children=True: [(1, "", "SunAwtCanvas", 2, True)],
    )

    result = operator.find_context_by_path_once(
        "0.0",
        name="2026-06-02",
        role="push button",
        require_showing=True,
        require_valid_bounds=True,
    )

    assert result == (None, None, [], {})


class FakeDll:
    def __init__(self, info):
        self.info = info

    def isJavaWindow(self, hwnd):
        return True

    def getAccessibleContextFromHWND(self, hwnd, vm_id, root_context):
        vm_id._obj.value = 100
        root_context._obj.value = 200
        return True

    def getAccessibleChildFromContext(self, vm_id, context, index):
        return 300

    def getAccessibleContextInfo(self, vm_id, context, info_ref):
        info = ctypes.cast(info_ref, ctypes.POINTER(AccessibleContextInfo)).contents
        info.name = "2026-06-02"
        info.role_en_US = "push button"
        info.states_en_US = self.info.states_en_US
        info.x = self.info.x
        info.y = self.info.y
        info.width = self.info.width
        info.height = self.info.height
        return True
