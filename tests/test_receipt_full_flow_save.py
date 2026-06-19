# 生命周期：持久维护
# 覆盖的业务场景：收款单完整流程的保存机制：Ctrl+S 键盘保存、前台/Oracle 守护、模态恢复
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_save.py


from tests._receipt_full_flow_helpers import (
    open_report_with_header_anchor,
    plan_row,
    run_one_row,
    save_receipt_by_ctrl_s,
)


def test_save_receipt_uses_keyboard_hotkey_not_jab_button_or_sendinput(monkeypatch):
    calls = {"hotkey": [], "states": 0}

    class FakeJAB:
        def click_save(self, timeout=None):
            raise AssertionError("收款单保存不能调用凭证/制单保存按钮查找")

        def wait_save_success(self, timeout=None):
            raise AssertionError("收款单保存不等保存成功提示作为触发闭包")

        def press_hotkey(self, *keys, wait=None):
            calls["hotkey"].append((keys, wait))

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.probe_receipt_entry_page",
        lambda _jab: {
            "ok": True,
            "scope": {"scope_hwnd": 12345},
            "method": "header-scope",
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.foreground_matches_window",
        lambda window: {"ok": True, "target_window": window, "foreground": {}},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.collect_receipt_new_windows",
        lambda _jab: [{"hwnd": 12345}],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_receipt_parent_new_ready",
        lambda _windows: {
            "ok": True,
            "usable_new_button_count": 1,
            "usable_new_buttons": [{"control": {"description": "新增(Ctrl+N)"}}],
        },
    )

    def fake_detect(_windows):
        calls["states"] += 1
        return {"ok": False, "reason": "已回到新增态"}

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_self_made_entry_state",
        fake_detect,
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), timeout=0.5)

    assert result["ok"] is True
    assert result["triggered"] is True
    assert calls["hotkey"] == [(("ctrl", "s"), 0)]
    assert result["hotkey"]["mode"] == "jab.press_hotkey"
    assert result["oracle"]["name"] == "receipt_parent_new_ready_after_save"
    assert result["oracle"]["parent_new_state"]["ok"] is True
    assert calls["states"] == 1


def test_save_receipt_stops_before_oracle_when_foreground_guard_fails(monkeypatch):
    class FakeJAB:
        def press_hotkey(self, *keys, wait=None):
            raise AssertionError("前台保护失败时不应触发 Ctrl+S")

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.probe_receipt_entry_page",
        lambda _jab: {"ok": True, "scope": {"scope_hwnd": 12345}},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.foreground_matches_window",
        lambda _window: {"ok": False, "reason": "当前前台窗口不是目标 NC 窗口"},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.collect_receipt_new_windows",
        lambda _jab: (_ for _ in ()).throw(
            AssertionError("Ctrl+S 未发出时不应继续等保存 oracle")
        ),
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), timeout=0.5)

    assert result["ok"] is False
    assert result["triggered"] is False
    assert "当前前台窗口不是目标 NC 窗口" in result["reason"]


def test_save_receipt_uses_entry_state_hwnd_without_header_scope_probe(monkeypatch):
    calls = {"guard": [], "hotkey": []}

    class FakeJAB:
        def press_hotkey(self, *keys, wait=None):
            calls["hotkey"].append((keys, wait))

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.collect_receipt_new_windows",
        lambda _jab: [{"hwnd": 24680, "controls": []}],
    )

    state_calls = {"count": 0}

    def fake_detect(_windows):
        state_calls["count"] += 1
        if state_calls["count"] == 1:
            return {
                "ok": True,
                "hits": [{"window": {"hwnd": 24680, "class_name": "SunAwtCanvas"}}],
            }
        return {"ok": False, "reason": "已回到新增态"}

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_self_made_entry_state",
        fake_detect,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.foreground_matches_window",
        lambda window: calls["guard"].append(window) or {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_receipt_parent_new_ready",
        lambda _windows: {"ok": True, "usable_new_button_count": 1},
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), timeout=0.5)

    assert result["ok"] is True
    assert calls["guard"] == [{"hwnd": 24680}]
    assert calls["hotkey"] == [(("ctrl", "s"), 0)]


def test_save_receipt_does_not_treat_missing_entry_buttons_as_success_without_new_button(
    monkeypatch,
):
    class FakeJAB:
        def press_hotkey(self, *keys, wait=None):
            return None

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.probe_receipt_entry_page",
        lambda _jab: {"ok": True, "scope": {"scope_hwnd": 12345}},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.foreground_matches_window",
        lambda _window: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.collect_receipt_new_windows",
        lambda _jab: [{"hwnd": 12345}],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_self_made_entry_state",
        lambda _windows: {"ok": False, "reason": "三按钮读不到"},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_receipt_parent_new_ready",
        lambda _windows: {"ok": False, "usable_new_button_count": 0},
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), timeout=0.01)

    assert result["ok"] is False
    assert result["oracle"]["ok"] is False
    assert result["oracle"]["parent_new_state"]["ok"] is False
    assert "新增" in result["reason"]


def test_run_one_row_recovers_modal_only_after_save_failure(monkeypatch):
    calls = {"save": 0, "recover": 0}

    class LocalFakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class LocalFakeVerifier:
        def __init__(self, *_args, **_kwargs):
            pass

        def start(self):
            pass

        def submit_snapshot(self, *args, **kwargs):
            return "snapshot-0"

        def submit_row_count(self, *args, **kwargs):
            return "rows-0"

        def wait(self, *args, **kwargs):
            return {"ok": True}

        def close(self, timeout=1.0):
            pass

    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", LocalFakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda *_args, **_kwargs: [
            {"ok": True, "label": "客户", "accepted_text": "ACME LTD"}
        ],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda *_args, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", LocalFakeVerifier
    )

    def fake_save(*_args, **_kwargs):
        calls["save"] += 1
        if calls["save"] == 1:
            return {"ok": False, "reason": "前台窗口不是目标 NC 窗口"}
        return {"ok": True, "triggered": True}

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.save_receipt_by_ctrl_s", fake_save
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
        lambda *_args, **_kwargs: (
            calls.__setitem__("recover", calls["recover"] + 1)
            or {"ok": True, "attempted": True}
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=True)

    assert report["ok"] is True
    assert calls == {"save": 2, "recover": 1}
    assert report["save"]["retried_after_modal_recovery"] is True
