from decimal import Decimal
from typing import cast

from core.nc_page_probe import NCPageProbe
from core.nc_state import NCStateDetector, normalize_generated_voucher


def test_normalize_generated_voucher():
    assert normalize_generated_voucher("000123", 9999) == 123
    assert normalize_generated_voucher("00000001", 9999) == 1
    assert normalize_generated_voucher("凭证号 88", 9999) == 88
    assert normalize_generated_voucher("0", 9999) is None
    assert normalize_generated_voucher("00000000", 9999) is None
    assert normalize_generated_voucher("10000", 9999) is None
    assert normalize_generated_voucher("未生成", 9999) is None


class FakeJAB:
    amount_col = 4
    partner_col = 3

    def read_window_table_cells(self, *args, **kwargs):
        return []

    def window_exists(self, *args, **kwargs):
        return False


class EmptyProbe:
    def collect_watched_controls(self):
        return []

    def detect_pending_toolbar(self, controls):
        return {"ok": False, "reason": "missing", "parent_count": 0}

    def collect_named_controls(self, *args, **kwargs):
        return []

    def collect_visible_buttons_by_desc_tokens(self, *args, **kwargs):
        return []

    def read_page_table_signatures(self, *args, **kwargs):
        return []


def test_detect_page_state_fast_fails_when_parent_and_tables_missing():
    detector = NCStateDetector(
        FakeJAB(),
        {},
        "2026-06-17",
        22,
        9999,
        lambda *args, **kwargs: None,
        lambda *args, **kwargs: None,
    )
    detector.probe = cast(NCPageProbe, EmptyProbe())

    state = detector.detect_page_state([])

    assert state.name == "error"
    assert "父页面/主表均未检测到" in state.reason


class ToolbarProbe:
    def __init__(self, toolbar, tables):
        self.toolbar = toolbar
        self.tables = tables

    def collect_watched_controls(self):
        return [
            {"name": "单据生成", "showing": True},
            {"name": "删除", "showing": True},
            {"name": "查询", "showing": True},
            {"name": "刷新", "showing": True},
            {"name": "选择", "showing": True},
            {"name": "生成", "showing": True},
        ]

    def detect_pending_toolbar(self, controls):
        return self.toolbar

    def collect_named_controls(self, *args, **kwargs):
        return []

    def collect_visible_buttons_by_desc_tokens(self, *args, **kwargs):
        return []

    def read_page_table_signatures(self, *args, **kwargs):
        return self.tables


def make_detector(probe):
    detector = NCStateDetector(
        FakeJAB(),
        {},
        "2026-06-17",
        22,
        9999,
        lambda *args, **kwargs: None,
        lambda *args, **kwargs: None,
    )
    detector.probe = cast(NCPageProbe, probe)
    return detector


def test_detect_page_state_accepts_pending_toolbar_with_41_col_table():
    detector = make_detector(
        ToolbarProbe(
            {
                "ok": True,
                "reason": "单据生成父页+待生成工具栏顺序匹配",
                "parent_count": 1,
            },
            [
                {
                    "row_count": 11,
                    "col_count": 41,
                    "voucher_values": [],
                    "rows": [],
                }
            ],
        )
    )

    state = detector.detect_page_state([])

    assert state.name == "pending"
    assert "待生成工具栏顺序匹配" in state.reason


def test_detect_page_state_treats_voucher_values_as_pending_when_toolbar_matches():
    detector = make_detector(
        ToolbarProbe(
            {
                "ok": True,
                "reason": "单据生成父页+待生成工具栏顺序匹配",
                "parent_count": 1,
            },
            [
                {
                    "row_count": 3,
                    "col_count": 41,
                    "voucher_values": ["00000000", "00000000"],
                    "rows": [],
                }
            ],
        )
    )

    state = detector.detect_page_state([])

    assert state.name == "pending"


def test_detect_page_state_rejects_parent_without_pending_toolbar():
    detector = make_detector(
        ToolbarProbe(
            {
                "ok": False,
                "reason": "待生成工具栏按钮顺序不匹配",
                "parent_count": 1,
            },
            [
                {
                    "row_count": 11,
                    "col_count": 41,
                    "voucher_values": [],
                    "rows": [],
                }
            ],
        )
    )

    state = detector.detect_page_state([])

    assert state.name == "error"
    assert "未知页面" in state.reason
