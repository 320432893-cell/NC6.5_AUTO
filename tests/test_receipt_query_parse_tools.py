# 生命周期：持久维护
# 覆盖的业务场景：收款单查询的解析/缓存工具：页签标签解析与缓存报告复用
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake JAB 替身）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_query_parse_tools.py


from tests._receipt_query_helpers import (
    Any,
    FakePagedJAB,
    cast,
    paged_query_config,
    parse_page_label,
    set_receipt_page_size,
)


def test_parse_page_label_reads_total_pages_and_records():
    assert parse_page_label("第1页 共66页 659条记录 每页显示") == {
        "total_pages": 66,
        "total_records": 659,
    }


def test_resolve_receipt_pagination_paths_uses_cached_report(monkeypatch):
    calls = []

    def fail_dynamic(_jab, _query_cfg):
        calls.append("dynamic")
        return {"ok": False}

    monkeypatch.setattr(
        "tools.receipt_query_pagination_paths.resolve_receipt_pagination_paths_dynamic",
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
