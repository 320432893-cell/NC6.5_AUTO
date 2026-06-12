import json
import subprocess
import ctypes

from tools import receipt_self_made_fill_trial as trial
from tools import receipt_new_probe
from tools.receipt_new_probe import (
    detect_self_made_entry_state,
    is_current_visible_control,
)


def test_open_self_made_requires_entry_state(monkeypatch):
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


def test_self_made_choose_cleanup_runs_after_success(monkeypatch):
    cleaned = {"called": False}

    monkeypatch.setattr(receipt_new_probe.os, "name", "nt")
    monkeypatch.setattr(
        receipt_new_probe.ctypes,
        "WINFUNCTYPE",
        ctypes.CFUNCTYPE,
        raising=False,
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "cleanup_awt_popup_residue",
        lambda: cleaned.update(called=True) or {"ok": True, "targets": []},
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "close_popup_hwnd",
        lambda hwnd: {"ok": True, "hwnd": hwnd},
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
            "path": None,
            "title": None,
            "class_name": None,
            "name": "新增",
            "role": None,
            "action": None,
            "return_timeout": 0.2,
            "wait": 0,
            "choose_self_made": True,
            "self_made_index": 0,
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
        receipt_new_probe, "collect_receipt_new_windows", lambda jab: next(sequence)
    )
    monkeypatch.setattr(receipt_new_probe, "find_named_controls", lambda *a, **k: [])
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
        "quick_check_self_made_entry_state",
        lambda jab: {"ok": True, "partial_ok": False, "windows": [{"hwnd": 2}]},
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "detect_self_made_entry_state",
        lambda windows: {"ok": bool(windows)},
    )

    assert receipt_new_probe.run(args)["residue_cleanup"]["ok"] is True
    assert cleaned["called"] is True


def test_cleanup_awt_popup_residue_moves_small_awt_windows(monkeypatch):
    calls = []

    def hwnd_value(hwnd):
        return int(getattr(hwnd, "value", hwnd))

    class FakeUser32:
        def EnumWindows(self, callback, lparam):
            callback(111, lparam)
            return True

        def EnableWindow(self, hwnd, enabled):
            calls.append(("EnableWindow", hwnd_value(hwnd), enabled))

        def ShowWindow(self, hwnd, cmd):
            calls.append(("ShowWindow", hwnd_value(hwnd), cmd))

        def SetWindowPos(self, hwnd, *args):
            calls.append(("SetWindowPos", hwnd_value(hwnd)))

        def PostMessageW(self, hwnd, *args):
            calls.append(("PostMessageW", hwnd_value(hwnd)))

        def GetDesktopWindow(self):
            return 999

        def RedrawWindow(self, hwnd, rect, region, flags):
            calls.append(("RedrawWindow", hwnd, flags))
            return True

    monkeypatch.setattr(receipt_new_probe.os, "name", "nt")
    monkeypatch.setattr(
        receipt_new_probe.ctypes,
        "WINFUNCTYPE",
        ctypes.CFUNCTYPE,
        raising=False,
    )
    monkeypatch.setattr(
        receipt_new_probe.ctypes,
        "windll",
        type("W", (), {"user32": FakeUser32()})(),
        raising=False,
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "describe_hwnd",
        lambda user32, hwnd: {
            "exists": True,
            "hwnd": int(hwnd.value),
            "visible": False,
            "class_name": "SunAwtWindow",
            "title": "",
            "width": 139,
            "height": 76,
        },
    )

    result = receipt_new_probe.cleanup_awt_popup_residue()

    assert result["targets"][0]["hwnd"] == 111
    assert ("ShowWindow", 111, 0) in calls
    assert ("SetWindowPos", 111) in calls
    assert ("PostMessageW", 111) in calls
    assert any(call[0] == "RedrawWindow" for call in calls)


def test_account_reference_stops_after_open_by_default(monkeypatch):
    class FakeJAB:
        def __init__(self):
            self.actions = []

        def do_action_by_path(self, *args, **kwargs):
            self.actions.append((args, kwargs))
            return True

        def get_scoped_windows(self, *args, **kwargs):
            return [(1234, "使用权参照", "SunAwtDialog", 99, True)]

        class Dll:
            @staticmethod
            def isJavaWindow(hwnd):
                return True

        dll = Dll()

    searched = {"called": False}

    def fail_if_search(*args, **kwargs):
        searched["called"] = True
        raise AssertionError("search must not run by default")

    monkeypatch.setattr(trial, "set_reference_search_text", fail_if_search)

    jab = FakeJAB()
    result = trial.set_header_account_by_reference(jab, "FTE123")

    assert result["ok"] is False
    assert result["blocked"] is True
    assert result["dialog"]["hwnd"] == 1234
    assert result["next_required"] == "foreground_check_account_reference"
    assert searched["called"] is False


def test_account_reference_existing_dialog_blocks_without_click(monkeypatch):
    class FakeJAB:
        def do_action_by_path(self, *args, **kwargs):
            raise AssertionError("must not click account button when dialog exists")

    monkeypatch.setattr(
        trial,
        "wait_reference_dialog",
        lambda jab, timeout=6.0: {
            "hwnd": 1234,
            "title": "使用权参照",
            "class_name": "SunAwtDialog",
            "pid": 99,
            "visible": True,
        },
    )

    result = trial.set_header_account_by_reference(FakeJAB(), "FTE123")

    assert result["ok"] is False
    assert result["blocked"] is True
    assert result["dialog"]["hwnd"] == 1234


def test_header_fill_writes_customer_before_date(monkeypatch):
    calls = []

    monkeypatch.setattr(
        trial,
        "set_text_by_control_name",
        lambda jab, control_name, value, **kwargs: (
            calls.append("财务组织") or {"ok": True}
        ),
    )

    def fake_set_header_field(jab, label, value, **kwargs):
        calls.append(label)
        return {"ok": True}

    monkeypatch.setattr(trial, "set_receipt_header_form_field", fake_set_header_field)
    monkeypatch.setattr(
        trial,
        "set_header_account_by_reference",
        lambda jab, account, continue_after_open=False: (
            calls.append("收款银行账户") or {"ok": False, "blocked": True}
        ),
    )

    trial.fill_header(
        object(),
        {
            "finance_org_code": "A001",
            "document_date": "2026-04-02",
            "customer_code": "YW03200",
            "currency": "美元",
            "bank_account": "FTE123",
        },
        continue_account_reference=True,
    )

    assert calls[:4] == ["财务组织", "客户", "单据日期", "币种"]


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


def test_visible_control_does_not_require_positive_coordinates():
    control = {
        "states": "enabled,visible,showing",
        "bounds": [-31992, -31862, 54, 30],
    }

    assert is_current_visible_control(control) is True


def test_context_commit_action_uses_jab_action():
    class FakeJAB:
        def get_action_names(self, vm_id, context):
            return ["单击"]

        def do_action(self, vm_id, context, action_name=None):
            self.called = (vm_id, context, action_name)
            return True

    jab = FakeJAB()

    result = trial.do_context_commit_action(jab, 1, 2)

    assert result["ok"] is True
    assert result["action"] == "单击"
    assert jab.called == (1, 2, "单击")
