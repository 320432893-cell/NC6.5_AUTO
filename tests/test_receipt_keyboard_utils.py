# 生命周期：持久维护
# 覆盖的业务场景：收款单明细/保存受保护快捷键不引入固定空等
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_keyboard_utils.py

from core import receipt_keyboard_utils as keyboard


def test_receipt_ctrl_i_ctrl_d_do_not_use_fixed_settle(monkeypatch):
    calls = []

    def fake_guard(window, key_name, sender, settle_seconds=0.8):
        calls.append((window, key_name, sender.__name__, settle_seconds))
        return {"ok": True, "key": key_name}

    monkeypatch.setattr(keyboard, "guarded_send_table_hotkey", fake_guard)

    assert keyboard.guarded_send_ctrl_i({"hwnd": 1})["ok"] is True
    assert keyboard.guarded_send_ctrl_d({"hwnd": 2})["ok"] is True

    assert calls == [
        ({"hwnd": 1}, "Ctrl+I", "send_hotkey_ctrl_i", 0.0),
        ({"hwnd": 2}, "Ctrl+D", "send_hotkey_ctrl_d", 0.0),
    ]


def test_guarded_send_table_hotkey_skips_sleep_when_settle_is_zero(monkeypatch):
    calls = {"sender": 0, "sleep": []}

    monkeypatch.setattr(
        keyboard,
        "foreground_matches_window",
        lambda _window: {"ok": True, "target_window": {}, "foreground": {}},
    )
    monkeypatch.setattr(
        keyboard.time,
        "sleep",
        lambda seconds: calls["sleep"].append(seconds),
    )

    def sender():
        calls["sender"] += 1

    result = keyboard.guarded_send_table_hotkey(
        {"hwnd": 1},
        "Ctrl+S",
        sender,
        settle_seconds=0.0,
    )

    assert result["ok"] is True
    assert calls == {"sender": 1, "sleep": []}
