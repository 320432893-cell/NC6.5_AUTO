# 生命周期：持久维护
# 覆盖的业务场景：自制单的新增开单与按钮识别：自制按钮信任、三态入口按钮、new 按钮优先级、父页守护与 Ctrl+N 兜底
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_self_made_new_entry.py


from tests._receipt_self_made_helpers import (
    ctypes,
    detect_self_made_entry_state,
    is_current_visible_control,
    json,
    receipt_new_probe,
    subprocess,
    trial,
)


def test_open_self_made_trusts_successful_self_made_action(monkeypatch):
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

    assert result["ok"] is True


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


def test_self_made_choose_does_not_run_residue_cleanup(monkeypatch):
    snapshots = {"called": False}
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
        receipt_new_probe,
        "guard_receipt_new_parent_page",
        lambda _jab, _config: {"ok": True, "state_label": "收款单录入"},
    )
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
        "collect_entry_context_snapshot",
        lambda jab: (
            snapshots.update(called=True)
            or {"ok": True, "state": {"ok": False}, "windows": [{"hwnd": 2}]}
        ),
    )
    monkeypatch.setattr(
        receipt_new_probe,
        "detect_self_made_entry_state",
        lambda windows: {"ok": bool(windows)},
    )

    report = receipt_new_probe.run(args)

    assert "residue_cleanup" not in report
    assert snapshots["called"] is True


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
            "path": None,
            "title": None,
            "class_name": "SunAwtFrame",
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
    monkeypatch.setattr(
        receipt_new_probe,
        "open_new_menu_with_ctrl_n",
        lambda _foreground: (_ for _ in ()).throw(
            AssertionError("正式收款开单不能回退 Ctrl+N")
        ),
    )

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


def test_visible_control_does_not_require_positive_coordinates():
    control = {
        "states": "enabled,visible,showing",
        "bounds": [-31992, -31862, 54, 30],
    }

    assert is_current_visible_control(control) is True
