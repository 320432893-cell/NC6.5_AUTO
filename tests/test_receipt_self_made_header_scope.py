# 生命周期：持久维护
# 覆盖的业务场景：自制单的表头锚点 scope 识别：上下文快照解析、canvas scope/anchor 校验、财务组织快捷键锚点与动态索引纠正
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_self_made_header_scope.py


from tests._receipt_self_made_helpers import (
    receipt_new_probe,
    trial,
)


def test_entry_context_snapshot_resolves_header_anchor(monkeypatch):
    class FakeJAB:
        pass

    windows = [
        {
            "hwnd": 24680,
            "class_name": "SunAwtCanvas",
            "visible": True,
            "is_java": True,
        }
    ]
    monkeypatch.setattr(
        receipt_new_probe,
        "collect_receipt_new_windows_compat",
        lambda _jab, **_kwargs: windows,
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "resolve_current_canvas_header_anchor",
        lambda _jab, _windows: {
            "ok": True,
            "scope_hwnd": 24680,
            "dynamic_index": 5,
            "dynamic_prefix": "0.0.1.0.0.0.0.5",
            "label_path": "0.0.1.0.0.0.0.5.0.0.0.1",
            "anchor_text": {"name": "财务组织(O)", "description": ""},
            "window": {
                "hwnd": 24680,
                "class_name": "SunAwtCanvas",
                "visible": True,
            },
        },
    )

    result = receipt_new_probe.collect_entry_context_snapshot(FakeJAB())

    assert result["confirmed"] is True
    assert result["state"]["partial_ok"] is True
    assert result["state"]["hits"][0]["control"]["dynamic_index"] == 5


def test_resolve_current_header_scope_probe_passes_jab_to_window_collector(
    monkeypatch,
):
    class FakeJAB:
        pass

    jab = FakeJAB()
    calls = []

    monkeypatch.setattr(
        trial,
        "collect_receipt_new_windows",
        lambda received_jab: calls.append(received_jab) or [],
    )
    monkeypatch.setattr(
        trial,
        "detect_self_made_entry_state",
        lambda _windows: {"ok": False, "hits": []},
    )
    monkeypatch.setattr(
        trial,
        "infer_receipt_header_scope_by_semantic",
        lambda _jab, scope_hwnd=None: {
            "ok": False,
            "reason": "not found",
            "scope_hwnd": scope_hwnd,
        },
    )

    result = trial.resolve_current_header_scope_for_probe(jab)

    assert calls == [jab]
    assert result["ok"] is False
    assert result["semantic_attempt"]["reason"] == "not found"


def test_header_fill_uses_provided_canvas_scope_when_anchor_matches(monkeypatch):
    calls = []

    monkeypatch.setattr(
        trial,
        "locate_receipt_header_scope",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("正式表头写入不应先跑 fake fast 扫描")
        ),
    )
    monkeypatch.setattr(
        trial,
        "infer_receipt_header_scope_by_semantic",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("已有 canvas hwnd + dynamic_index 时不应先全局语义扫")
        ),
    )
    monkeypatch.setattr(
        trial,
        "validate_receipt_header_scope_anchor",
        lambda _jab, scope_hwnd, dynamic_index, **_kwargs: (
            calls.append(("anchor", scope_hwnd, dynamic_index))
            or {
                "ok": True,
                "scope_hwnd": scope_hwnd,
                "mode": "provided-canvas-anchor",
                "dynamic_index": dynamic_index,
                "dynamic_prefix": f"0.0.1.0.0.0.0.{dynamic_index}",
                "matched_labels": ["财务组织"],
                "anchor_text": {"name": "财务组织(O)", "description": ""},
            }
        ),
    )

    class FakeJAB:
        def release_contexts(self, _vm_id, _contexts):
            pass

    def fake_set_header_field(jab, label, value, dynamic_index, scope_hwnd, **kwargs):
        calls.append((label, dynamic_index, scope_hwnd))
        return {"ok": True, "path": f"path-{label}"}

    monkeypatch.setattr(
        trial,
        "set_receipt_header_dynamic_field",
        fake_set_header_field,
    )

    trial.fill_header(
        FakeJAB(),
        {
            "finance_org_code": "A001",
            "document_date": "2026-04-02",
            "customer_code": "YW03200",
            "currency": "美元",
            "bank_account": "FTE123",
        },
        scope_hwnd=123,
        dynamic_index=5,
    )

    assert calls[0] == ("anchor", 123, 5)
    assert ("财务组织", 5, 123) in calls


def test_header_scope_stops_when_index_missing(monkeypatch):
    monkeypatch.setattr(
        trial,
        "infer_receipt_header_scope_by_semantic",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("正式表头缺 dynamic_index 时不能语义兜底")
        ),
    )

    scope = trial.resolve_receipt_header_scope(object(), scope_hwnd=123)

    assert scope["ok"] is False
    assert scope["scope_hwnd"] == 123
    assert scope["dynamic_index"] is None
    assert "不走语义兜底" in scope["reason"]


def test_header_scope_uses_provided_canvas_anchor_before_semantic(monkeypatch):
    calls = []

    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("表头 scope 不应再用提供的 dynamic_index 验证 path")
        ),
    )

    monkeypatch.setattr(
        trial,
        "infer_receipt_header_scope_by_semantic",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("当前 canvas 锚点通过后不应再语义扫描")
        ),
    )

    def fake_anchor(_jab, scope_hwnd, dynamic_index, **_kwargs):
        calls.append(("anchor", scope_hwnd, dynamic_index))
        return {
            "ok": True,
            "scope_hwnd": scope_hwnd,
            "mode": "provided-canvas-anchor",
            "dynamic_index": dynamic_index,
            "dynamic_prefix": f"0.0.1.0.0.0.0.{dynamic_index}",
            "matched_labels": ["财务组织"],
            "anchor_text": {"name": "财务组织(O)", "description": ""},
        }

    monkeypatch.setattr(trial, "validate_receipt_header_scope_anchor", fake_anchor)

    scope = trial.resolve_receipt_header_scope(
        object(), scope_hwnd=123, dynamic_index=2
    )

    assert scope["ok"] is True
    assert scope["dynamic_index"] == 2
    assert scope["mode"] == "provided-canvas-anchor"
    assert calls == [("anchor", 123, 2)]


def test_header_scope_validates_with_provided_anchor_path(monkeypatch):
    calls = []
    anchor_path = "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.0"

    class Info:
        name = "财务组织(O)"
        description = "财务组织(O)"

    class FakeJAB:
        def find_context_by_path_once(self, path, **kwargs):
            calls.append((path, kwargs))
            return object(), 1, [object()], {"hwnd": kwargs.get("scope_hwnd")}

        def get_context_info(self, _vm_id, _context):
            return Info()

        def release_contexts(self, _vm_id, _contexts):
            pass

    monkeypatch.setattr(
        trial,
        "infer_receipt_header_scope_by_semantic",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("已有当前 canvas 锚点失败/成功都不应掉语义")
        ),
    )

    scope = trial.resolve_receipt_header_scope(
        FakeJAB(),
        scope_hwnd=919586,
        dynamic_index=2,
        anchor_path=anchor_path,
    )

    assert scope["ok"] is True
    assert scope["semantic_label_path"] == anchor_path
    assert calls[0][0] == anchor_path
    assert calls[0][1]["scope_hwnd"] == 919586


def test_header_label_text_matches_shortcut_suffix():
    class Info:
        name = "财务组织(O)"
        description = ""

    assert trial.header_label_text_matches(Info(), "财务组织") is True


def test_header_scope_anchor_requires_exact_finance_org_shortcut():
    class ShortcutInfo:
        name = "财务组织(O)"
        description = ""

    class PlainInfo:
        name = "财务组织"
        description = ""

    class OtherInfo:
        name = "收款财务组织"
        description = ""

    assert trial.header_scope_anchor_text_matches(ShortcutInfo()) is True
    assert trial.header_scope_anchor_text_matches(PlainInfo()) is False
    assert trial.header_scope_anchor_text_matches(OtherInfo()) is False


def test_finance_org_anchor_label_path_matches_observed_current_canvas_path():
    trial.clear_receipt_header_path_template_cache()
    assert (
        trial.build_receipt_header_dynamic_label_path(2, "财务组织")
        == "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.0"
    )


def test_resolve_header_anchor_rejects_plain_finance_org_text(monkeypatch):
    class Info:
        role_en_US = "label"
        role = "label"
        states_en_US = "visible,showing"
        states = "visible,showing"
        name = "财务组织"
        description = ""
        childrenCount = 0

    class FakeDLL:
        def isJavaWindow(self, _hwnd):
            return True

        def getAccessibleContextFromHWND(self, hwnd, _vm_id_ref, _root_context):
            return bool(hwnd)

    class FakeJAB:
        dll = FakeDLL()
        max_depth = 50
        max_children = 20

        def get_scoped_windows(self, scope_hwnd, include_children=True):
            return [(scope_hwnd, "", "SunAwtCanvas", 1234, True)]

        def get_context_info(self, _vm_id, _context):
            return Info()

        def release_contexts(self, _vm_id, _contexts):
            return None

    monkeypatch.setattr(
        trial,
        "find_header_label_context_with_window",
        lambda *_args, **_kwargs: (
            object(),
            1,
            [object()],
            [0, 0, 1, 0, 0, 0, 0, 2, 0, 0],
            {"hwnd": 919586, "class_name": "SunAwtCanvas"},
        ),
    )

    result = trial.resolve_receipt_header_anchor_in_canvas(FakeJAB(), 919586)

    assert result["ok"] is False
    assert "不匹配" in result["reason"]
    assert result["anchor_text"]["name"] == "财务组织"


def test_resolve_header_anchor_corrects_dynamic_index_by_customer(monkeypatch):
    class Info:
        name = "财务组织(O)"
        description = "财务组织(O)"

    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return Info()

        def release_contexts(self, _vm_id, _contexts):
            pass

    monkeypatch.setattr(
        trial,
        "find_header_label_context_with_window",
        lambda *_args, **_kwargs: (
            object(),
            1,
            [object()],
            [0, 1, 0, 0, 0, 0, 3, 0, 0],
            {"hwnd": 197550, "class_name": "SunAwtCanvas"},
        ),
    )
    monkeypatch.setattr(
        trial,
        "correct_header_anchor_dynamic_index_by_customer",
        lambda _jab, _scope_hwnd, dynamic_index: {
            "ok": True,
            "source": "customer-semantic-correction",
            "dynamic_index": 5,
            "path": ("0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.17.0"),
            "current_attempt": {"dynamic_index": dynamic_index},
        },
    )

    result = trial.resolve_receipt_header_anchor_in_canvas(FakeJAB(), 197550)

    assert result["ok"] is True
    assert result["dynamic_index"] == 5
    assert result["initial_dynamic_index"] == 3
    assert result["mode"] == "current-canvas-anchor-corrected-by-customer"
