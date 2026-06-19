# 生命周期：持久维护
# 覆盖的业务场景：收款单查询窗口的 F3 打开/复用与查询条件填写、确认与结果就绪等待
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake JAB 替身）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_query_window.py


from tests._receipt_query_helpers import (
    FakeJAB,
    FakePagedJAB,
    FakeReceiptQueryJAB,
    QUERY_DATE_FROM_PATH,
    QUERY_DATE_TO_PATH,
    QUERY_FINANCE_ORG_PATH,
    ensure_query_window,
    fill_receipt_query,
    pytest,
    receipt_config,
    wait_after_query_confirm,
)


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

    monkeypatch.setattr("tools.receipt_query_fill.JABOperator", make_jab)

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

    monkeypatch.setattr("tools.receipt_query_fill.JABOperator", make_jab)

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
    monkeypatch.setattr("tools.receipt_query_fill.time.sleep", waits.append)
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

    monkeypatch.setattr("tools.receipt_query_fill.JABOperator", make_jab)

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
