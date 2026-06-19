# 生命周期：持久维护
# 覆盖的业务场景：收款单查询的分页翻页、页大小设置与稳定性等待
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake JAB 替身）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_query_pagination.py


from tests._receipt_query_helpers import (
    FakeFailingNextPageJAB,
    FakeNoScopePagedJAB,
    FakePagedJAB,
    paged_query_config,
    read_receipt_result_pages,
    set_receipt_page_size,
    wait_receipt_result_stable,
)


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
    after_stability = report["after_stability"]
    # 业务意图：页大小已是目标值时跳过 after 稳定性等待（reason=page_size_already_target）。
    # 仅断言承载该意图的关键键值，避免新增 meta 键即碎。
    assert after_stability["skipped"] is True
    assert after_stability["reason"] == "page_size_already_target"
    assert after_stability["ok"] is None
    assert after_stability["tables"] == []
    assert report["after_stability_seconds"] == 0.0
    assert report["pagination_plan_reason"] == "total_records_within_page_size"
    assert jab.actions == []
    assert [table["row_count"] for table in tables] == [2]


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
        "tools.receipt_query_pagination_paths.resolve_receipt_pagination_paths_dynamic",
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
    monkeypatch.setattr("tools.receipt_query_pagination.time.sleep", waits.append)
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
    monkeypatch.setattr("tools.receipt_query_pagination.time.sleep", waits.append)

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
    monkeypatch.setattr("tools.receipt_query_pagination.time.sleep", waits.append)

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
