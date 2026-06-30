import json
import subprocess
import ctypes
import pytest

from core import receipt_self_made_fill_trial as trial
from core import receipt_new_probe
from core.receipt_new_probe import (
    detect_self_made_entry_state,
    is_current_visible_control,
)


def test_open_self_made_requires_confirmed_entry_state(monkeypatch):
    payload = {
        "open": {"ok": True},
        "choose_self_made": {"ok": True},
        "entry_state": {"ok": False, "names": ["保存(Ctrl+S)", "暂存"]},
    }

    def fake_run(*args, **kwargs):
        return subprocess.CompletedProcess(
            args=args[0],
            returncode=0,
            stdout=json.dumps(payload, ensure_ascii=False),
            stderr="",
        )

    monkeypatch.setattr(trial.subprocess, "run", fake_run)

    result = trial.run_receipt_new_probe()

    assert result["ok"] is False


def test_open_self_made_accepts_three_entry_buttons(monkeypatch):
    payload = {
        "open": {"ok": True},
        "choose_self_made": {"ok": True},
        "entry_state": {
            "ok": True,
            "names": ["保存(Ctrl+S)", "暂存", "取消(Ctrl+Q)"],
        },
    }

    def fake_run(*args, **kwargs):
        return subprocess.CompletedProcess(
            args=args[0],
            returncode=0,
            stdout=json.dumps(payload, ensure_ascii=False),
            stderr="",
        )

    monkeypatch.setattr(trial.subprocess, "run", fake_run)

    result = trial.run_receipt_new_probe()

    assert result["ok"] is True


def test_entry_state_detects_button_descriptions():
    windows = [
        {
            "controls": [
                {"name": "保存", "description": "保存(Ctrl+S)"},
                {"name": "暂存", "description": ""},
                {"name": "取消", "description": "取消(Ctrl+Q)"},
            ]
        }
    ]

    result = detect_self_made_entry_state(windows)

    assert result["ok"] is True
    assert result["names"] == ["保存(Ctrl+S)", "取消(Ctrl+Q)", "暂存"]


def test_new_button_priority_prefers_showing_valid_button():
    mirror = {
        "window": {"hwnd": 1},
        "control": {
            "description": "新增业务数据(Ctrl+N)",
            "states": "enabled,visible",
            "bounds": [-1, -1, -1, -1],
        },
    }
    real = {
        "window": {"hwnd": 1},
        "control": {
            "description": "新增(Ctrl+N)",
            "states": "enabled,visible,showing",
            "bounds": [7, 130, 67, 30],
        },
    }

    buttons = [mirror, real]
    buttons.sort(key=receipt_new_probe.new_button_priority)

    assert buttons[0] is real


@pytest.mark.skipif(
    not hasattr(ctypes, "WinDLL"),
    reason="Windows 专有路径:依赖 ctypes.WinDLL,Linux 无法运行",
)
def test_self_made_choose_cleans_popup_before_entry_wait(monkeypatch):
    calls = []
    monkeypatch.setattr(receipt_new_probe.os, "name", "nt")
    monkeypatch.setattr(
        receipt_new_probe.ctypes,
        "WINFUNCTYPE",
        ctypes.CFUNCTYPE,
        raising=False,
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "close_popup_hwnd",
        lambda hwnd: calls.append(("cleanup", hwnd)) or {"ok": True, "hwnd": hwnd},
    )

    class FakeJAB:
        hide_blank_awt_windows_enabled = False

        def ensure_started(self):
            pass

        def close(self):
            pass

    args = type(
        "Args",
        (),
        {
            "config": "config.json",
            "method": "button",
            "class_name": None,
            "name": "新增",
            "role": None,
            "action": None,
            "return_timeout": 0.2,
            "wait": 0,
            "choose_self_made": True,
            "json": False,
            "summary": True,
        },
    )()

    sequence = iter(
        [
            [],
            [{"hwnd": 1, "is_java": True, "visible": True}],
            [{"hwnd": 2, "is_java": True, "visible": True}],
        ]
    )

    monkeypatch.setattr(receipt_new_probe, "load_config", lambda config: {})
    monkeypatch.setattr(receipt_new_probe, "JABOperator", lambda cfg: FakeJAB())
    monkeypatch.setattr(
        receipt_new_probe,
        "guard_receipt_new_parent_page",
        lambda _jab, _config: {"ok": True, "state_label": "收款单录入"},
    )
    monkeypatch.setattr(
        receipt_new_probe, "collect_receipt_new_windows", lambda jab: next(sequence)
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "open_new_menu_with_known_buttons",
        lambda *args, **kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "wait_for_self_made_popup",
        lambda jab, before, wait: {
            "ok": True,
            "popup": {"hwnd": 456},
            "windows": [{"hwnd": 1}],
        },
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "choose_self_made_menu_item",
        lambda *args, **kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "wait_self_made_entry_ready",
        lambda jab, **kwargs: (
            calls.append(("entry_wait", kwargs.get("popup_hwnd")))
            or {"ok": True, "state": {"ok": True}, "windows": [{"hwnd": 2}]}
        ),
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "detect_self_made_entry_state",
        lambda windows: {"ok": bool(windows)},
    )

    report = receipt_new_probe.run(args)

    assert "residue_cleanup" not in report
    assert calls == [("cleanup", 456), ("entry_wait", 456)]


def test_new_probe_stops_before_new_when_parent_guard_fails(monkeypatch):
    calls = {"collect": 0, "open": 0}

    class FakeJAB:
        config = {"receipt_entry": {"state_label": "收款单录入"}}
        hide_blank_awt_windows_enabled = False

        def ensure_started(self):
            pass

        def close(self):
            pass

    args = type(
        "Args",
        (),
        {
            "config": "config.json",
            "method": "button",
            "class_name": "SunAwtFrame",
            "name": "新增",
            "role": None,
            "action": None,
            "return_timeout": 0.2,
            "wait": 0,
            "choose_self_made": True,
            "json": False,
            "summary": True,
        },
    )()

    monkeypatch.setattr(
        receipt_new_probe,
        "guard_receipt_new_parent_page",
        lambda _jab, _config: {"ok": False, "reason": "不是收款单录入"},
    )
    monkeypatch.setattr(receipt_new_probe, "foreground_info", lambda: {})

    def collect_should_not_run(_jab):
        calls["collect"] += 1
        return []

    def open_should_not_run(*_args, **_kwargs):
        calls["open"] += 1
        return {"ok": True}

    monkeypatch.setattr(
        receipt_new_probe, "collect_receipt_new_windows", collect_should_not_run
    )
    monkeypatch.setattr(
        receipt_new_probe, "open_new_menu_with_known_buttons", open_should_not_run
    )

    report = receipt_new_probe.run(args, jab=FakeJAB())

    assert report["open"]["ok"] is False
    assert report["receipt_parent_guard"]["ok"] is False
    assert calls == {"collect": 0, "open": 0}


def test_new_button_path_does_not_fallback_to_ctrl_n(monkeypatch):
    assert not hasattr(receipt_new_probe, "open_new_menu_with_ctrl_n")

    report = receipt_new_probe.open_new_menu_with_known_buttons(
        jab=object(),
        args=type(
            "Args", (), {"method": "button", "action": None, "return_timeout": 0}
        )(),
        buttons=[],
        all_buttons=[],
        foreground={"root_class_name": "YonyouUWnd"},
    )

    assert report["ok"] is False
    assert "不回退 Ctrl+N" in report["reason"]


def test_header_fill_writes_customer_before_date(monkeypatch):
    calls = []

    def fake_set_header_field(jab, label, value, dynamic_index, scope_hwnd, **kwargs):
        calls.append((label, value, kwargs.get("accepted_text")))
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
            "header_currency_code": "USD",
            "bank_account": "FTE123",
        },
        scope_hwnd=123,
        dynamic_index=2,
    )

    assert [item[0] for item in calls] == ["财务组织", "客户", "单据日期", "币种", "结算方式"]
    assert calls[3] == ("币种", "USD", ["USD", "美元"])


def test_currency_acceptance_texts_accepts_code_and_chinese_name():
    assert trial.currency_acceptance_texts("USD", "美元") == ["USD", "美元"]
    assert trial.currency_acceptance_texts("CNY", "人民币") == ["CNY", "人民币"]
    assert trial.currency_acceptance_texts("人民币") == ["人民币", "CNY"]


def test_header_fill_uses_path_flow_without_background_semantic(monkeypatch):
    calls = []

    monkeypatch.setattr(
        trial,
        "set_receipt_header_dynamic_field",
        lambda _jab, label, *_args, **_kwargs: calls.append(label) or {"ok": True},
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
        config = {"jab": {}}

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


def test_customer_name_candidate_rejects_java_object_string():
    assert not trial.is_valid_customer_name_candidate("[Ljava.lang.String;@75acf5a0")
    assert not trial.is_valid_customer_name_candidate("YW00178")
    assert trial.is_valid_customer_name_candidate("SERDIA ELETRONICA INDL LTDA")
    assert trial.is_valid_customer_name_candidate("上海移为通信技术股份有限公司")


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
            AssertionError("正式表头缺 dynamic_index 时不能走旧全局语义兜底")
        ),
    )

    scope = trial.resolve_receipt_header_scope(object(), scope_hwnd=123)

    assert scope["ok"] is False
    assert scope["scope_hwnd"] == 123
    assert scope["dynamic_index"] is None
    assert "deep scan 未找到财务组织" in scope["reason"]


def test_header_scope_uses_provided_canvas_anchor_before_semantic(monkeypatch):
    calls = []

    monkeypatch.setattr(
        trial,
        "validate_receipt_header_scope_anchor",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("财务组织 path 命中后不应再用旧锚点验证")
        ),
    )

    monkeypatch.setattr(
        trial,
        "infer_receipt_header_scope_by_semantic",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("当前 canvas 锚点通过后不应再语义扫描")
        ),
    )

    def fake_finance_scope(_jab, scope_hwnd, **kwargs):
        calls.append(("finance-path", scope_hwnd, kwargs))
        return {
            "ok": True,
            "scope_hwnd": scope_hwnd,
            "mode": "finance-org-shallow-semantic",
            "dynamic_index": kwargs["preferred_dynamic_index"],
            "dynamic_prefix": f"0.0.1.0.0.0.0.{kwargs['preferred_dynamic_index']}",
            "matched_labels": ["财务组织"],
            "anchor_text": {"name": "财务组织(O)", "description": ""},
        }

    monkeypatch.setattr(
        trial,
        "find_finance_org_header_scope_by_paths",
        fake_finance_scope,
    )

    scope = trial.resolve_receipt_header_scope(
        object(), scope_hwnd=123, dynamic_index=2
    )

    assert scope["ok"] is True
    assert scope["dynamic_index"] == 2
    assert scope["mode"] == "finance-org-shallow-semantic"
    assert calls == [
        (
            "finance-path",
            123,
            {"preferred_dynamic_index": 2},
        )
    ]


def test_header_scope_can_trust_open_step_scope_without_revalidating(monkeypatch):
    monkeypatch.setattr(
        trial,
        "validate_receipt_header_scope_anchor",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("完整流程已在开单阶段解析 scope，不应重复验证")
        ),
    )

    scope = trial.resolve_receipt_header_scope(
        object(),
        scope_hwnd=123,
        dynamic_index=2,
        anchor_path="0.anchor",
        trust_provided_scope=True,
    )

    assert scope["ok"] is True
    assert scope["dynamic_index"] == 2
    assert scope["semantic_label_path"] == "0.anchor"
    assert scope["mode"] == "provided-canvas-anchor-trusted"


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
    monkeypatch.setattr(
        trial,
        "find_finance_org_header_scope_by_paths",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "fast path miss",
        },
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
        == "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.0.1.1.0"
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
        == "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.0.1.1.2.1.0"
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
    assert calls[0][0] == "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.0.1.1.2.1.0"
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

    assert calls[1][0] == "客户"
    assert calls[1][1] is None
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
        "find_receipt_header_field_by_scoped_label",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("dynamic path 命中后不能再逐字段语义搜索")
        ),
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
        path_template={"text_suffix_template": "x.{index}.0"},
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
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_finance_org_container",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "container miss",
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
        "find_receipt_header_field_by_scoped_label",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "scoped fast miss",
        },
    )
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
    assert result["dynamic_path_attempt"] is None
    assert result["path_attempt"]["container_attempt"]["reason"] == "container miss"
    assert semantic_calls == [
        {
            "scope_hwnd": 919586,
            "timeout": trial.HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT,
            "include_scoped": False,
        }
    ]


def test_header_dynamic_field_uses_finance_org_container_for_first_customer(
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
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("客户首定位无学习模板时不应先试默认 path")
        ),
    )
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_finance_org_container",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 919586},
            "path": "container.customer.path",
            "label_path": "container.customer.label",
            "source": "finance-org-container-label-following-text",
        },
    )
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_live_semantic",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("容器命中后不应继续语义 DFS")
        ),
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
        header_scope={"label_path": "0.anchor"},
    )

    assert result["ok"] is True
    assert result["source"] == "finance-org-container-label-following-text"
    assert result["path"] == "container.customer.path"


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
        "find_receipt_header_field_by_finance_org_container",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "container miss",
        },
    )
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_scoped_label",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "scoped fast miss",
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
        path_template={"text_suffix_template": "x.{index}.0"},
    )

    assert result["ok"] is False
    assert result["stage"] == "resolve"
    assert result["path_attempt"]["reason"] == "semantic label not found"
    assert (
        result["path_attempt"]["live_semantic_timeout"]
        == trial.HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT
    )
    assert result["path_attempt"]["dynamic_path_attempt"]["reason"] == "dynamic path missing"


def test_header_live_semantic_prefers_scoped_fast_before_deep_semantic(monkeypatch):
    class FakeJAB:
        pass

    deep_calls = []

    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_scoped_label",
        lambda *_args, **_kwargs: {
            "ok": True,
            "label": "客户",
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "path": "scoped.customer.path",
            "label_path": "scoped.customer.label",
            "window": {"hwnd": 123},
        },
    )
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_semantic_label",
        lambda *_args, **_kwargs: deep_calls.append(True)
        or {"ok": False, "reason": "should not run"},
    )

    result = trial.find_receipt_header_field_by_live_semantic(
        FakeJAB(),
        "客户",
        scope_hwnd=123,
    )

    assert result["ok"] is True
    assert result["source"] == "scoped-label-fast-before-semantic"
    assert result["path"] == "scoped.customer.path"
    assert deep_calls == []


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


def test_header_dynamic_field_requires_accepted_text_before_success(monkeypatch):
    class Info:
        name = "币种"
        description = "人民币"
        role = "text"
        role_en_US = "text"
        states = "enabled,visible,showing,editable"
        states_en_US = "enabled,visible,showing,editable"

    class FakeJAB:
        dll = object()

        def get_context_info(self, _vm_id, _context):
            return Info()

        def get_text_context_value(self, _vm_id, _context):
            return "人民币"

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
            "path": "currency.path",
            "label_path": "currency.label",
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
    monkeypatch.setattr(trial.time, "sleep", lambda *_args, **_kwargs: None)

    result = trial.set_receipt_header_dynamic_field(
        FakeJAB(),
        "币种",
        "美元",
        2,
        919586,
        accepted_text="美元",
    )

    assert result["ok"] is False
    assert result["set_ok"] is True
    assert result["acceptance_probe"]["accepted"] is False
    assert result["post_write_snapshot"]["description"] == "人民币"


def test_visible_control_does_not_require_positive_coordinates():
    control = {
        "states": "enabled,visible,showing",
        "bounds": [-31992, -31862, 54, 30],
    }

    assert is_current_visible_control(control) is True


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
        path_template={"text_suffix_template": "x.{index}.0"},
    )

    assert result["ok"] is True
    assert result["post_write_snapshot"]["written"] is False


def test_finance_org_uses_control_name_without_write_path_in_formal_flow(monkeypatch):
    class FakeJAB:
        pass

    monkeypatch.setattr(
        trial,
        "set_finance_org_by_legacy_control_name",
        lambda *_args, **_kwargs: {
            "ok": False,
            "method": "legacy-control-name-guarded-paste-enter",
            "reason": "foreground mismatch",
        },
    )

    result = trial.set_receipt_header_dynamic_field(
        FakeJAB(),
        "财务组织",
        "A001",
        2,
        919586,
    )

    assert result["ok"] is False
    assert result["method"] == "legacy-control-name-guarded-paste-enter"
    assert result["reason"] == "foreground mismatch"
    assert "dynamic_write_attempt" not in result


def test_finance_org_formal_flow_prefers_control_name_even_if_write_path_exists(monkeypatch):
    legacy_result = {
        "ok": True,
        "method": "legacy-control-name-guarded-paste-enter",
        "source": "legacy-control-name",
        "path": "0.legacy",
        "set_ok": True,
        "enter_ok": True,
    }
    monkeypatch.setattr(
        trial,
        "set_finance_org_by_legacy_control_name",
        lambda *_args, **_kwargs: legacy_result,
    )

    result = trial.set_receipt_header_dynamic_field(
        object(),
        "财务组织",
        "A001",
        2,
        919586,
    )

    assert result["ok"] is True
    assert result["method"] == "legacy-control-name-guarded-paste-enter"
    assert result["enter_ok"] is True
    assert result["dynamic_index"] == 2
    assert result["dynamic_prefix"].endswith(".2")


def test_finance_org_header_scope_prefers_shallow_semantic(monkeypatch):
    def fake_shallow(*_args, **_kwargs):
        return {
            "ok": True,
            "scope_hwnd": 919586,
            "mode": "finance-org-shallow-semantic",
            "dynamic_index": 5,
            "dynamic_prefix": "0.0.1.0.0.0.0.5",
            "label_path": "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.0.1.1.0",
        }

    class FakeJAB:
        def find_context_by_path_once(self, *_args, **_kwargs):
            raise AssertionError("shallow semantic hit should skip path scan")

    monkeypatch.setattr(
        trial,
        "_find_finance_org_header_scope_by_shallow_semantic",
        fake_shallow,
    )

    result = trial.find_finance_org_header_scope_by_paths(FakeJAB(), 919586)

    assert result["ok"] is True
    assert result["dynamic_index"] == 5
    assert result["deep_scan_skipped"] is True
    assert result["preferred_dynamic_index"] is None


def test_finance_org_header_scope_falls_back_to_deep_semantic_without_path_probe(monkeypatch):
    calls = []

    def fake_semantic(_jab, _scope_hwnd, **kwargs):
        calls.append(kwargs["scan_mode"])
        if kwargs["scan_mode"] == "finance-org-shallow-semantic":
            return {
                "ok": False,
                "mode": kwargs["scan_mode"],
                "reason": "finance org anchor not found in shallow scope",
            }
        return {
            "ok": True,
            "scope_hwnd": 919586,
            "mode": kwargs["scan_mode"],
            "dynamic_index": 6,
            "dynamic_prefix": "0.0.1.0.0.0.0.6",
            "label_path": "0.0.1.0.0.0.0.6.deep",
        }

    class FakeJAB:
        def find_context_by_path_once(self, *_args, **_kwargs):
            raise AssertionError("deep fallback must not probe guessed dynamic paths")

    monkeypatch.setattr(
        trial,
        "_find_finance_org_header_scope_by_shallow_semantic",
        fake_semantic,
    )

    result = trial.find_finance_org_header_scope_by_paths(FakeJAB(), 919586)

    assert result["ok"] is True
    assert result["mode"] == "finance-org-deep-semantic"
    assert result["dynamic_index"] == 6
    assert result["preferred_dynamic_index"] is None
    assert calls == ["finance-org-shallow-semantic", "finance-org-deep-semantic"]
    assert result["shallow_semantic_attempt"]["ok"] is False


def test_finance_org_shallow_semantic_canonical_label_path_uses_dynamic_index(
    monkeypatch,
):
    class Dll:
        def getAccessibleContextFromHWND(self, _hwnd, vm_id_ref, root_ref):
            vm_id_ref._obj.value = 1
            root_ref._obj.value = 100
            return True

        def getAccessibleChildFromContext(self, _vm_id, context, index):
            children = {
                100: [200],
                200: [300],
                300: [],
            }
            try:
                return children[context][index]
            except IndexError:
                return 0

    class Info:
        def __init__(self, role, name="", children=0):
            self.role = role
            self.role_en_US = role
            self.name = name
            self.description = ""
            self.states = "visible,showing"
            self.states_en_US = "visible,showing"
            self.childrenCount = children

    class FakeJAB:
        dll = Dll()
        max_children = 100

        def get_context_info(self, _vm_id, context):
            return {
                100: Info("panel", children=1),
                200: Info("panel", children=1),
                300: Info("label", name="财务组织(O)", children=0),
            }[context]

        def release_contexts(self, *_args):
            return None

    monkeypatch.setattr(
        trial,
        "extract_receipt_header_dynamic_index",
        lambda _path: 2,
    )

    result = trial._find_finance_org_header_scope_by_shallow_semantic(
        FakeJAB(),
        919586,
        order="bfs",
        max_depth=5,
        max_nodes=10,
    )

    assert result["ok"] is True
    assert result["dynamic_index"] == 2
    assert result["label_path"] == (
        "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.0.1.1.0"
    )


def test_finance_org_formal_flow_only_uses_control_name(monkeypatch):
    legacy_result = {
        "ok": True,
        "method": "legacy-control-name-guarded-paste-enter",
        "source": "legacy-control-name",
        "path": "0.legacy",
        "set_ok": True,
        "enter_ok": True,
    }
    monkeypatch.setattr(
        trial,
        "set_finance_org_by_legacy_control_name",
        lambda *_args, **_kwargs: legacy_result,
    )

    result = trial.set_receipt_header_dynamic_field(
        object(),
        "财务组织",
        "A001",
        2,
        919586,
    )

    assert result["ok"] is True
    assert result["method"] == "legacy-control-name-guarded-paste-enter"
    assert "dynamic_write_attempt" not in result
    assert result["enter_ok"] is True


def test_finance_org_control_name_write_uses_guarded_paste_enter_first(monkeypatch):
    class Info:
        def __init__(self, description=""):
            self.description = description
            self.name = "财务组织(O)"
            self.role = "text"
            self.role_en_US = "text"
            self.states = "enabled,visible,showing,editable"
            self.states_en_US = "enabled,visible,showing,editable"

    class FakeJAB:
        def __init__(self):
            self.read_count = 0

        def get_context_info(self, _vm_id, _context):
            self.read_count += 1
            if self.read_count >= 3:
                return Info("上海移为通信技术股份有限公司")
            return Info()

        def get_text_context_value(self, _vm_id, _context):
            return "A001"

        def set_text_context(self, *_args):
            raise AssertionError(
                "guarded paste success should not call setTextContents"
            )

        def release_contexts(self, _vm_id, _owned_contexts):
            pass

    monkeypatch.setattr(trial.time, "sleep", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        trial,
        "find_context_with_window",
        lambda *_args, **_kwargs: (
            object(),
            1,
            [],
            [2, 1, 0],
            {"hwnd": 919586},
        ),
    )
    monkeypatch.setattr(
        trial,
        "guarded_paste_header_value",
        lambda _jab, _vm_id, _context, window_info, value: {
            "ok": True,
            "method": "guarded-clipboard-paste",
            "enter_ok": True,
            "window": window_info,
            "value": value,
        },
    )

    result = trial.set_finance_org_by_legacy_control_name(
        FakeJAB(),
        "A001",
        scope_hwnd=919586,
    )

    assert result["ok"] is True
    assert result["method"] == "legacy-control-name-guarded-paste-enter"
    assert result["guarded_paste"]["enter_ok"] is True
    assert result["set_text_ok"] is False
    assert result["acceptance_probe"]["accepted"] is True
    assert result["accepted_text"] == "上海移为通信技术股份有限公司"


def test_finance_org_control_name_write_blocks_until_chinese_acceptance(monkeypatch):
    class Info:
        name = "财务组织(O)"
        description = ""
        role = "text"
        role_en_US = "text"
        states = "enabled,visible,showing,editable"
        states_en_US = "enabled,visible,showing,editable"

    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return Info()

        def get_text_context_value(self, _vm_id, _context):
            return "A001"

        def set_text_context(self, *_args):
            raise AssertionError(
                "guarded paste success should not call setTextContents"
            )

        def release_contexts(self, _vm_id, _owned_contexts):
            pass

    monkeypatch.setattr(trial.time, "sleep", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        trial,
        "find_context_with_window",
        lambda *_args, **_kwargs: (
            object(),
            1,
            [],
            [2, 1, 0],
            {"hwnd": 919586},
        ),
    )
    monkeypatch.setattr(
        trial,
        "guarded_paste_header_value",
        lambda _jab, _vm_id, _context, window_info, value: {
            "ok": True,
            "method": "guarded-clipboard-paste",
            "enter_ok": True,
            "window": window_info,
            "value": value,
        },
    )

    result = trial.set_finance_org_by_legacy_control_name(
        FakeJAB(),
        "A001",
        scope_hwnd=919586,
        accepted_text="上海移为通信技术股份有限公司",
    )

    assert result["ok"] is False
    assert result["method"] == "legacy-control-name-guarded-paste-enter"
    assert result["guarded_paste"]["enter_ok"] is True
    assert result["set_text_ok"] is False
    assert result["acceptance_probe"]["accepted"] is False
    assert result["acceptance_probe"]["reason"] == "财务组织未确认解析为中文"


def test_finance_org_acceptance_can_use_scoped_chinese_probe(monkeypatch):
    class Info:
        name = "财务组织(O)"
        description = ""
        role = "text"
        role_en_US = "text"
        states = "enabled,visible,showing,editable"
        states_en_US = "enabled,visible,showing,editable"

    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return Info()

        def get_text_context_value(self, _vm_id, _context):
            return "A001"

    monkeypatch.setattr(trial.time, "sleep", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        trial,
        "probe_finance_org_accepted_text_in_scope",
        lambda _jab, expected_text, scope_hwnd=None: {
            "ok": True,
            "accepted": True,
            "expected_text": expected_text,
            "path": "0.scope.hit",
            "scope_hwnd": scope_hwnd,
            "snapshot": {
                "accepted": True,
                "written": False,
                "text": "",
                "name": expected_text,
                "description": "",
            },
        },
    )

    result = trial.confirm_finance_org_accepted(
        FakeJAB(),
        1,
        object(),
        expected_text="上海移为通信技术股份有限公司",
        value="A001",
        scope_hwnd=919586,
    )

    assert result["accepted"] is True
    assert result["source"] == "scope-text-probe"
    assert result["scope_probe"]["path"] == "0.scope.hit"


def test_finance_org_acceptance_uses_scope_probe_only_after_current_context_poll(
    monkeypatch,
):
    class Info:
        name = "财务组织(O)"
        description = "A001"
        role = "text"
        role_en_US = "text"
        states = "enabled,visible,showing,editable"
        states_en_US = "enabled,visible,showing,editable"

    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return Info()

        def get_text_context_value(self, _vm_id, _context):
            return "A001"

    calls = {"scope_probe": 0}
    monkeypatch.setattr(trial.time, "sleep", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        trial,
        "probe_finance_org_accepted_text_in_scope",
        lambda *_args, **_kwargs: calls.__setitem__(
            "scope_probe",
            calls["scope_probe"] + 1,
        )
        or {"ok": False, "accepted": False},
    )

    result = trial.confirm_finance_org_accepted(
        FakeJAB(),
        1,
        object(),
        expected_text="上海移为通信技术股份有限公司",
        value="A001",
        scope_hwnd=919586,
        timeout=0.01,
        interval=0.001,
    )

    assert result["accepted"] is False
    assert len(result["attempts"]) > 1
    assert calls["scope_probe"] == 1


def test_guarded_paste_header_value_uses_jab_press_key_enter(monkeypatch):
    calls = []
    result_ref = {}

    monkeypatch.setattr(
        trial,
        "foreground_matches_window",
        lambda _window: {"ok": True, "target_window": {"hwnd": 919586}},
    )
    monkeypatch.setattr(trial, "get_clipboard_text", lambda: "old")
    monkeypatch.setattr(
        trial,
        "set_clipboard_text",
        lambda text: calls.append(("set_clipboard_text", text)),
    )
    monkeypatch.setattr(
        trial,
        "restore_clipboard_text",
        lambda text: calls.append(("restore_clipboard_text", text)) or True,
    )
    monkeypatch.setattr(
        trial,
        "send_hotkey_ctrl_a",
        lambda: calls.append(("hotkey", "ctrl+a")),
    )
    monkeypatch.setattr(
        trial,
        "send_hotkey_ctrl_v",
        lambda: calls.append(("hotkey", "ctrl+v")),
    )

    class DLL:
        def requestFocus(self, _vm_id, _context):
            return True

    class FakeJAB:
        dll = DLL()

        def press_key(self, key, wait=None):
            calls.append(("jab.press_key", key, wait))

    result = trial.guarded_paste_header_value(
        FakeJAB(),
        1,
        object(),
        {"hwnd": 919586},
        "A001",
    )
    result_ref.update(result)

    assert result_ref["ok"] is True
    assert result_ref["enter_method"] == "jab.press_key"
    assert ("jab.press_key", "enter", 0) in calls
    assert ("restore_clipboard_text", "old") in calls


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
        "find_receipt_header_field_by_scoped_label",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "scoped label missing",
        },
    )
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
        path_template={"text_suffix_template": "x.{index}.0"},
    )

    assert result["ok"] is False
    assert result["stage"] == "resolve"
    assert result["path_attempt"]["reason"] == "semantic label not found"
    assert (
        result["path_attempt"]["live_semantic_timeout"]
        == trial.HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT
    )
    assert result["path_attempt"]["dynamic_path_attempt"]["reason"] == "dynamic path missing"
