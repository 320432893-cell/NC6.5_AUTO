# 生命周期：持久维护
# 覆盖的业务场景：收款单明细 writer 只负责写入，不保留同步读回/重试旧逻辑
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_detail_writer.py

from tools import receipt_detail_writer as writer


def test_write_field_once_defers_readback_to_verifier(monkeypatch):
    field = {"col": 4, "name": "收款银行账户", "value_key": "bank_account"}

    monkeypatch.setattr(
        writer,
        "focus_detail_cell",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        writer,
        "keyboard_write_selected_cell",
        lambda *_args, **_kwargs: {"ok": True},
    )

    result = writer.write_field_once(
        jab=object(),
        located={"best": {"window": {"hwnd": 1}, "path": "0.1"}},
        table_window={"hwnd": 1},
        row_index=0,
        row_count=1,
        field=field,
        business={"bank_account": "FTE1219165931831"},
        attempt_no=1,
        current_col=None,
    )

    assert result["ok"] is True
    assert "selected_before_write" not in result
    assert result["commit_target"]["skipped"] is True
    assert "后台 verifier" in result["commit_target"]["reason"]


def test_write_field_once_passes_failure_recovery_hook_without_calling_it(monkeypatch):
    field = {"col": 4, "name": "收款银行账户", "value_key": "bank_account"}
    hooks = []

    monkeypatch.setattr(
        writer,
        "focus_detail_cell",
        lambda *_args, **_kwargs: {"ok": True},
    )

    def fake_keyboard_write(*_args, **kwargs):
        hooks.append(kwargs.get("recover_after_failure"))
        return {"ok": True}

    monkeypatch.setattr(writer, "keyboard_write_selected_cell", fake_keyboard_write)

    def recover_after_failure():
        raise AssertionError("正常成功路径不应检查弹窗")

    result = writer.write_field_once(
        jab=object(),
        located={"best": {"window": {"hwnd": 1}, "path": "0.1"}},
        table_window={"hwnd": 1},
        row_index=0,
        row_count=1,
        field=field,
        business={"bank_account": "FTE1219165931831"},
        attempt_no=1,
        current_col=None,
        recover_after_failure=recover_after_failure,
    )

    assert result["ok"] is True
    assert hooks == [recover_after_failure]


def test_keyboard_write_recovers_after_clipboard_open_failure(monkeypatch):
    from tools import receipt_detail_screen_writer as screen

    calls = {"clipboard": 0, "recover": 0}

    monkeypatch.setattr(
        screen,
        "foreground_matches_table",
        lambda _window: {"ok": True, "foreground": {"hwnd": 1}},
    )
    monkeypatch.setattr(screen, "send_virtual_key", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(screen, "send_hotkey_ctrl_a", lambda: None)
    monkeypatch.setattr(screen, "send_hotkey_ctrl_v", lambda: None)
    monkeypatch.setattr(
        screen, "guarded_press_virtual_key", lambda *_args: {"ok": True}
    )
    monkeypatch.setattr(screen, "safe_clipboard_read", lambda: None)
    monkeypatch.setattr(screen, "restore_clipboard_text", lambda _text: True)

    def fake_set_clipboard(_text):
        calls["clipboard"] += 1
        if calls["clipboard"] == 1:
            raise RuntimeError("OpenClipboard failed")

    monkeypatch.setattr(screen, "set_clipboard_text", fake_set_clipboard)

    def recover_after_failure():
        calls["recover"] += 1
        return {"ok": True, "attempted": True}

    result = screen.keyboard_write_selected_cell(
        {"hwnd": 1},
        "YW00178",
        recover_after_failure=recover_after_failure,
    )

    assert result["ok"] is True
    assert calls == {"clipboard": 2, "recover": 1}
