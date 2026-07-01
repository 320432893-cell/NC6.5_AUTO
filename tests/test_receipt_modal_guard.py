# 生命周期：持久维护
# 覆盖的业务场景：收款单动作失败后 Java dialog 聚焦并 Alt+C 恢复
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB

from core.receipt_modal_guard import recover_cancelable_modal_now


def test_recover_cancelable_modal_now_sends_alt_c(monkeypatch):
    calls = {"alt_c": 0}
    snapshots = [
        [
            {
                "hwnd": 100,
                "title": "提示",
                "class_name": "SunAwtDialog",
                "root_hwnd": 1,
                "cancel_controls": [{"name": "取消", "description": "Alt+C"}],
            }
        ],
        [],
    ]

    monkeypatch.setattr(
        "core.receipt_modal_guard.collect_visible_java_dialogs",
        lambda _jab: snapshots.pop(0) if snapshots else [],
    )
    monkeypatch.setattr(
        "core.receipt_modal_guard.send_hotkey_alt_c",
        lambda: calls.__setitem__("alt_c", calls["alt_c"] + 1),
    )
    monkeypatch.setattr(
        "core.receipt_modal_guard.focus_window",
        lambda hwnd: {"ok": True, "hwnd": hwnd},
    )

    result = recover_cancelable_modal_now(object(), stage="明细主行", settle_timeout=0)

    assert result["ok"] is True
    assert result["attempted"] is True
    assert result["stage"] == "明细主行"
    assert calls["alt_c"] == 1
