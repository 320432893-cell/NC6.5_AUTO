# 生命周期：持久维护
# 覆盖的业务场景：自制单的表头字段 path 路由/dynamic/semantic 兜底：填写顺序、模板学习、动态路径优先与语义兜底、后端字段状态
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_self_made_header_field.py


from tests._receipt_self_made_helpers import (
    trial,
)


def test_header_fill_writes_customer_before_date(monkeypatch):
    calls = []

    def fake_set_header_field(jab, label, value, dynamic_index, scope_hwnd, **kwargs):
        calls.append(label)
        return {
            "ok": True,
            "path": trial.build_receipt_header_dynamic_path(dynamic_index, label),
        }

    monkeypatch.setattr(
        trial,
        "set_receipt_header_dynamic_field",
        fake_set_header_field,
    )
    monkeypatch.setattr(
        trial,
        "infer_receipt_header_scope_by_semantic",
        lambda *_args, **_kwargs: {
            "ok": True,
            "scope_hwnd": 123,
            "mode": "semantic-path-inference",
            "dynamic_index": 2,
            "dynamic_prefix": "0.0.1.0.0.0.0.2",
        },
    )
    monkeypatch.setattr(
        trial,
        "validate_receipt_header_scope_anchor",
        lambda _jab, scope_hwnd, dynamic_index, **_kwargs: {
            "ok": True,
            "scope_hwnd": scope_hwnd,
            "mode": "provided-canvas-anchor",
            "dynamic_index": dynamic_index,
            "dynamic_prefix": f"0.0.1.0.0.0.0.{dynamic_index}",
            "matched_labels": ["财务组织"],
            "anchor_text": {"name": "财务组织(O)", "description": ""},
        },
    )

    class FakeJAB:
        def release_contexts(self, _vm_id, _contexts):
            pass

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
        dynamic_index=2,
    )

    assert calls == ["财务组织", "客户", "单据日期", "币种", "结算方式"]


# 注：原 test_header_fill_uses_path_flow_without_background_semantic 与
# test_header_fill_writes_customer_before_date 测同一业务意图（fill_header 按
# 财务组织→客户→单据日期→币种→结算方式 顺序写表头）。fill_header 源码不读
# jab.config、也无 background-semantic 分支，前者的 config/无 path 返回属惰性差异、
# 无独立 oracle，已合并删除（模板学习分支由
# test_header_fill_learns_header_template_from_customer 独立覆盖）。


def test_probe_header_semantic_field_speed_releases_context(monkeypatch):
    calls = []

    class FakeJAB:
        hide_blank_awt_windows_enabled = True

        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            calls.append(("ensure",))

        def release_contexts(self, vm_id, contexts):
            calls.append(("release", vm_id, tuple(contexts)))

        def close(self):
            calls.append(("close",))

    monkeypatch.setattr(trial, "JABOperator", FakeJAB)
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_semantic_label",
        lambda _jab, label, timeout=1.5, **_kwargs: {
            "ok": True,
            "label": label,
            "context": object(),
            "vm_id": 7,
            "owned_contexts": ["ctx"],
            "path": "semantic.customer.path",
            "label_path": "semantic.customer.label",
            "window": {"hwnd": 123},
        },
    )

    result = trial.probe_header_semantic_field_speed(
        {},
        "客户",
        timeout=0.35,
        repeat=1,
    )

    assert result["ok"] is True
    assert result["readonly"] is True
    assert result["attempts"][0]["path"] == "semantic.customer.path"
    assert calls == [("ensure",), ("release", 7, ("ctx",)), ("close",)]


def test_normalize_header_probe_label_accepts_ascii_keys():
    assert trial.normalize_header_probe_label("customer") == "客户"
    assert trial.normalize_header_probe_label("date") == "单据日期"
    assert trial.normalize_header_probe_label("currency") == "币种"
    assert trial.normalize_header_probe_label("settlement") == "结算方式"
    assert trial.normalize_header_probe_label("finance") == "财务组织"


def test_customer_name_candidate_rejects_java_object_string():
    assert not trial.is_valid_customer_name_candidate("[Ljava.lang.String;@75acf5a0")
    assert not trial.is_valid_customer_name_candidate("YW00178")
    assert trial.is_valid_customer_name_candidate("SERDIA ELETRONICA INDL LTDA")
    assert trial.is_valid_customer_name_candidate("上海移为通信技术股份有限公司")


def test_first_valid_customer_name_prefers_description():
    result = trial.first_valid_customer_name(
        [
            {
                "source": "path",
                "path": "customer.path",
                "description": "[Ljava.lang.String;@75acf5a0",
                "valid_values": [],
            },
            {
                "source": "path-nearby",
                "path": "customer.name.path",
                "description": "SERDIA ELETRONICA INDL LTDA",
                "valid_values": ["SERDIA ELETRONICA INDL LTDA"],
            },
        ]
    )

    assert result == {
        "value": "SERDIA ELETRONICA INDL LTDA",
        "source": "path-nearby",
        "path": "customer.name.path",
        "parent_path": None,
        "field": "description",
    }


def test_customer_name_readback_report_only_exposes_field_candidates(monkeypatch):
    class FakeJAB:
        hide_blank_awt_windows_enabled = True

        def __init__(self, _config):
            pass

        def ensure_started(self):
            pass

        def close(self):
            pass

    monkeypatch.setattr(trial, "JABOperator", FakeJAB)
    monkeypatch.setattr(
        trial,
        "resolve_current_header_scope_for_probe",
        lambda _jab, timeout=None: {
            "ok": True,
            "scope_hwnd": 123,
            "dynamic_index": 5,
        },
    )
    monkeypatch.setattr(
        trial,
        "collect_customer_field_candidates_for_scope",
        lambda *_args, **_kwargs: [
            {
                "ok": True,
                "source": "path",
                "path": "customer.path",
                "description": "SERDIA ELETRONICA INDL LTDA",
                "valid_values": ["SERDIA ELETRONICA INDL LTDA"],
            }
        ],
    )

    report = trial.probe_customer_name_readback({}, timeout=0.35)

    assert report["ok"] is True
    assert report["best"]["value"] == "SERDIA ELETRONICA INDL LTDA"
    assert "candidates" not in report
    assert report["field_candidates"][0]["source"] == "path"


def test_semantic_header_field_uses_label_window_for_text_path(monkeypatch):
    calls = []
    provided_scope_hwnd = 1001
    label_window_hwnd = 2002

    monkeypatch.setattr(
        trial,
        "find_header_label_context_with_window",
        lambda *_args, **_kwargs: (
            object(),
            1,
            [object()],
            [0, 1, 0, 0, 0, 0, 2, 0, 0, 0, 1, 1, 0, 0, 0, 1, 1, 1, 0],
            {"hwnd": label_window_hwnd, "class_name": "SunAwtCanvas"},
        ),
    )

    class FakeJAB:
        def release_contexts(self, vm_id, contexts):
            calls.append(("release", vm_id, len(contexts)))

        def find_context_by_path_once(self, path, **kwargs):
            calls.append(("find_path", path, kwargs.get("scope_hwnd")))
            return (
                object(),
                2,
                [object()],
                {"hwnd": kwargs.get("scope_hwnd"), "class_name": "SunAwtCanvas"},
            )

    result = trial.find_receipt_header_field_by_semantic_label(
        FakeJAB(),
        "财务组织",
        scope_hwnd=provided_scope_hwnd,
    )

    assert result["ok"] is True
    assert calls[1] == (
        "find_path",
        "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.2.1.0",
        label_window_hwnd,
    )
    assert result["window"]["hwnd"] == label_window_hwnd


def test_finance_org_text_path_matches_observed_current_canvas_path():
    trial.clear_receipt_header_path_template_cache()
    assert (
        trial.build_receipt_header_dynamic_path(2, "财务组织")
        == "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.2.1.0"
    )


def test_header_template_matches_observed_customer_path():
    trial.clear_receipt_header_path_template_cache()
    template = trial.infer_header_path_template_from_field(
        "0.0.1.0.0.0.0.3.0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.17.0",
        3,
        "客户",
    )

    assert template is not None
    assert template["text_suffix_template"] == (
        "0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.{index}.0"
    )
    assert (
        trial.build_receipt_header_path_from_template(3, "单据日期", template)
        == "0.0.1.0.0.0.0.3.0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.5.0"
    )


def test_find_finance_org_field_uses_observed_text_path_with_text_role():
    calls = []

    class FakeJAB:
        def find_context_by_path_once(self, path, **kwargs):
            calls.append((path, kwargs))
            return object(), 1, [object()], {"hwnd": kwargs.get("scope_hwnd")}

    result = trial.find_receipt_header_field_by_dynamic_path(
        FakeJAB(),
        "财务组织",
        2,
        scope_hwnd=919586,
        require_showing=False,
        require_valid_bounds=False,
    )

    assert result["ok"] is True
    assert calls[0][0] == "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.2.1.0"
    assert calls[0][1]["role"] == "text"
    assert calls[0][1]["scope_hwnd"] == 919586


def test_header_fill_learns_header_template_from_customer(monkeypatch):
    calls = []

    def fake_set_header_field(
        _jab,
        label,
        _value,
        dynamic_index,
        _scope_hwnd,
        **kwargs,
    ):
        calls.append((label, kwargs.get("path_template")))
        if label == "客户":
            return {
                "ok": True,
                "path": ("0.0.1.0.0.0.0.3.0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.17.0"),
            }
        return {
            "ok": True,
            "path": trial.build_receipt_header_dynamic_path(dynamic_index, label),
        }

    monkeypatch.setattr(
        trial, "set_receipt_header_dynamic_field", fake_set_header_field
    )
    monkeypatch.setattr(
        trial,
        "validate_receipt_header_scope_anchor",
        lambda _jab, scope_hwnd, dynamic_index, **_kwargs: {
            "ok": True,
            "scope_hwnd": scope_hwnd,
            "mode": "provided-canvas-anchor",
            "dynamic_index": dynamic_index,
            "dynamic_prefix": f"0.0.1.0.0.0.0.{dynamic_index}",
        },
    )

    steps = trial.fill_header(
        object(),
        {
            "finance_org_code": "A001",
            "document_date": "2026-04-02",
            "customer_code": "YW03200",
            "currency": "美元",
            "bank_account": "FTE123",
        },
        scope_hwnd=123,
        dynamic_index=3,
    )

    assert calls[1] == ("客户", None)
    assert calls[2][0] == "单据日期"
    assert calls[2][1]["source"] == "learned-from-客户"
    assert steps[1]["header_path_template_learned"]["source"] == "learned-from-客户"


def test_header_dynamic_field_prefers_dynamic_path_over_scoped_label(monkeypatch):
    class Info:
        name = "客户"
        description = ""
        role = "text"
        role_en_US = "text"
        states = "enabled,visible,showing,editable"
        states_en_US = "enabled,visible,showing,editable"

    class FakeJAB:
        dll = object()

        def get_context_info(self, _vm_id, _context):
            return Info()

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _owned_contexts):
            pass

    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 919586},
            "path": "dynamic.customer.path",
            "label_path": "dynamic.customer.label",
            "source": "path",
        },
    )
    monkeypatch.setattr(
        trial,
        "guarded_paste_header_value",
        lambda *_args: {
            "ok": True,
            "method": "guarded-clipboard-paste",
            "enter_ok": True,
        },
    )

    result = trial.set_receipt_header_dynamic_field(
        FakeJAB(),
        "客户",
        "YW00178",
        2,
        919586,
    )

    assert result["ok"] is True
    assert result["source"] == "path"
    assert result["path"] == "dynamic.customer.path"


def test_header_dynamic_field_uses_live_semantic_after_path_miss(
    monkeypatch,
):
    class Info:
        name = "客户"
        description = ""
        role = "text"
        role_en_US = "text"
        states = "enabled,visible,showing,editable"
        states_en_US = "enabled,visible,showing,editable"

    class FakeJAB:
        dll = object()

        def get_context_info(self, _vm_id, _context):
            return Info()

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _owned_contexts):
            pass

    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "dynamic path missing",
        },
    )
    semantic_calls = []

    def fake_live_semantic(*_args, **kwargs):
        semantic_calls.append(kwargs)
        return {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 919586},
            "path": "live.customer.path",
            "label_path": "live.customer.label",
            "source": "semantic-live-after-path-miss",
        }

    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_live_semantic",
        fake_live_semantic,
    )
    monkeypatch.setattr(
        trial,
        "guarded_paste_header_value",
        lambda *_args: {
            "ok": True,
            "method": "guarded-clipboard-paste",
            "enter_ok": True,
        },
    )

    result = trial.set_receipt_header_dynamic_field(
        FakeJAB(),
        "客户",
        "YW00178",
        2,
        919586,
    )

    assert result["ok"] is True
    assert result["source"] == "semantic-live-after-path-miss"
    assert result["dynamic_path_attempt"]["reason"] == "dynamic path missing"
    assert semantic_calls == [
        {
            "scope_hwnd": 919586,
            "timeout": trial.HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT,
        }
    ]


def test_header_dynamic_field_short_semantic_failure_does_not_deep_scan(
    monkeypatch,
):
    class FakeJAB:
        dll = object()

        def release_contexts(self, _vm_id, _owned_contexts):
            pass

    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "dynamic path missing",
        },
    )
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_semantic_label",
        lambda *_args, **kwargs: {
            "ok": False,
            "reason": "semantic label not found",
            "timeout": kwargs.get("timeout"),
        },
    )

    result = trial.set_receipt_header_dynamic_field(
        FakeJAB(),
        "客户",
        "YW00178",
        2,
        919586,
    )

    assert result["ok"] is False
    assert result["stage"] == "resolve"
    assert result["path_attempt"]["source"] == "semantic-live-after-path-miss"
    assert (
        result["path_attempt"]["live_semantic_timeout"]
        == trial.HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT
    )


def test_backend_field_state_accepts_description_without_foreground():
    class Info:
        name = "财务组织(O)"
        description = "上海移为通信技术股份有限公司"
        role = "text"
        role_en_US = "text"
        states = "enabled,visible,showing"
        states_en_US = "enabled,visible,showing"

    result = trial.describe_backend_field_state(
        Info(),
        text="",
        value="A001",
        accepted_text="上海移为通信技术股份有限公司",
    )

    assert result["accepted"] is True
    assert result["description"] == "上海移为通信技术股份有限公司"


def test_backend_field_state_tracks_written_value_before_business_correction():
    class Info:
        name = "财务组织(O)"
        description = "A001"
        role = "text"
        role_en_US = "text"
        states = "enabled,visible,showing"
        states_en_US = "enabled,visible,showing"

    result = trial.describe_backend_field_state(
        Info(),
        text="",
        value="A001",
        accepted_text="上海移为通信技术股份有限公司",
    )

    assert result["written"] is True
    assert result["accepted"] is False


def test_header_dynamic_field_records_snapshot_without_blocking_after_guarded_paste(
    monkeypatch,
):
    class Info:
        name = "客户"
        description = ""
        role = "text"
        role_en_US = "text"
        states = "enabled,visible,showing"
        states_en_US = "enabled,visible,showing"

    class FakeJAB:
        dll = object()

        def get_context_info(self, _vm_id, _context):
            return Info()

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _owned_contexts):
            pass

    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 123},
            "path": "0.1",
            "source": "path",
        },
    )
    monkeypatch.setattr(
        trial,
        "guarded_paste_header_value",
        lambda *_args: {
            "ok": True,
            "method": "guarded-clipboard-paste",
            "enter_ok": True,
        },
    )

    result = trial.set_receipt_header_dynamic_field(
        FakeJAB(),
        "客户",
        "YW00178",
        2,
        123,
    )

    assert result["ok"] is True
    assert result["post_write_snapshot"]["written"] is False


def test_header_dynamic_field_blocks_when_path_fails(monkeypatch):
    class FakeJAB:
        dll = object()

    class FakePreload:
        def snapshot(self, timeout=0.0):
            return {
                "status": "ready",
                "fields": {
                    "客户": {
                        "ok": True,
                        "path": "0.0.semantic.0",
                        "label_path": "0.0.semantic",
                    }
                },
            }

    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "dynamic path missing",
        },
    )
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_semantic_label",
        lambda *_args, **kwargs: {
            "ok": False,
            "reason": "semantic label not found",
            "timeout": kwargs.get("timeout"),
            "source": "semantic-live-after-path-miss",
        },
    )

    result = trial.set_receipt_header_dynamic_field(
        FakeJAB(),
        "客户",
        "YW00178",
        2,
        123,
    )

    assert result["ok"] is False
    assert result["stage"] == "resolve"
    assert result["path_attempt"]["reason"] == "semantic label not found"
    assert (
        result["path_attempt"]["live_semantic_timeout"]
        == trial.HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT
    )
    assert result["path_attempt"]["dynamic_path_attempt"]["reason"] == (
        "dynamic path missing"
    )
