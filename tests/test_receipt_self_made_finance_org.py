# 生命周期：持久维护
# 覆盖的业务场景：自制单的财务组织写入与中文验收：control-name guarded paste+Enter、阻塞至中文接受、scoped 中文探针
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_self_made_finance_org.py


from tests._receipt_self_made_helpers import (
    trial,
)


def test_finance_org_does_not_fallback_to_dynamic_path_when_control_write_fails(
    monkeypatch,
):
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
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("财务组织控件名写入失败后不应回退动态 path")
        ),
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


def test_finance_org_prefers_control_name_guarded_paste_enter(monkeypatch):
    legacy_result = {
        "ok": True,
        "method": "legacy-control-name-guarded-paste-enter",
        "source": "legacy-control-name",
        "path": "0.legacy",
        "set_ok": True,
        "enter_ok": True,
        "guarded_paste": {
            "ok": True,
            "method": "guarded-clipboard-paste",
            "enter_ok": True,
        },
        "post_write_snapshot": {
            "accepted": True,
            "written": True,
            "text": "A001",
            "description": "",
        },
    }
    monkeypatch.setattr(
        trial,
        "set_finance_org_by_legacy_control_name",
        lambda *_args, **_kwargs: legacy_result,
    )
    monkeypatch.setattr(
        trial,
        "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("财务组织控件名写入成功后不应跑 dynamic path")
        ),
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
