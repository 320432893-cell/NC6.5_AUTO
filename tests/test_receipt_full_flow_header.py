# 生命周期：持久维护
# 覆盖的业务场景：收款单完整流程的表头锚点/路由：accepted-text、客户名回读、anchor path、dynamic index、scope hwnd
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_header.py


from tests._receipt_full_flow_helpers import (
    FakeInfo,
    extract_entry_anchor_path,
    extract_entry_dynamic_index,
    extract_entry_scope_hwnd,
    extract_header_accepted_text,
    open_self_made_entry,
    read_customer_name_after_header,
    wait_receipt_header_anchor_in_current_canvas,
)


def test_extract_header_accepted_text_rejects_java_object_string():
    assert (
        extract_header_accepted_text(
            [
                {
                    "label": "客户",
                    "value": "YW00178",
                    "post_write_snapshot": {
                        "description": "[Ljava.lang.String;@75acf5a0",
                    },
                }
            ],
            "客户",
        )
        == ""
    )


def test_read_customer_name_after_header_uses_customer_description(monkeypatch):
    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return FakeInfo(description="INDUSTRIAS METALURGICAS PESCARMONA")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "path": "customer.path",
            "label_path": "customer.label",
        },
    )

    result = read_customer_name_after_header(
        FakeJAB(),
        [
            {
                "ok": True,
                "label": "客户",
                "value": "YW00178",
                "dynamic_index": 4,
                "path": "customer.path",
            }
        ],
        4,
        197550,
    )

    assert result["ok"] is True
    assert result["value"] == "INDUSTRIAS METALURGICAS PESCARMONA"
    assert result["source"] == "path-readback"


def test_extract_entry_anchor_path_uses_exact_finance_org_anchor():
    report = {
        "entry_state": {
            "hits": [
                {
                    "control": {
                        "path": "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.0",
                        "name": "财务组织(O)",
                        "description": "财务组织(O)",
                    }
                }
            ]
        }
    }

    assert (
        extract_entry_anchor_path(report) == "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.0"
    )


def test_header_anchor_wait_polls_current_canvas_every_point_two(monkeypatch):
    calls = {"anchor": [], "sleep": []}

    def fake_anchor(_jab, hwnd, timeout=0.05):
        calls["anchor"].append((hwnd, timeout))
        if len(calls["anchor"]) == 1:
            return {"ok": False, "reason": "not ready"}
        return {
            "ok": True,
            "scope_hwnd": hwnd,
            "dynamic_index": 5,
            "dynamic_prefix": "0.0.1.0.0.0.0.5",
        }

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.resolve_receipt_header_anchor_in_canvas",
        fake_anchor,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.time.sleep",
        lambda seconds: calls["sleep"].append(seconds),
    )

    result = wait_receipt_header_anchor_in_current_canvas(
        object(),
        919586,
        timeout=1.2,
        interval=0.2,
    )

    assert result["ok"] is True
    assert result["dynamic_index"] == 5
    assert len(calls["anchor"]) == 2
    assert calls["anchor"][0][0] == 919586
    assert calls["sleep"] == [0.2]
    assert result["poll_interval"] == 0.2


def test_open_self_made_entry_always_runs_new_probe(monkeypatch):
    calls = []

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.run_receipt_new_probe",
        lambda: calls.append("new-probe") or {"ok": True, "mode": "new-self-made"},
    )

    assert open_self_made_entry({"receipt_entry": {}}) == {
        "ok": True,
        "mode": "new-self-made",
    }
    assert calls == ["new-probe"]


def test_open_self_made_entry_reuses_existing_jab(monkeypatch):
    calls = []
    jab = object()

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.run_receipt_new_probe",
        lambda: (_ for _ in ()).throw(AssertionError("不应起子进程开单")),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.run_receipt_new_probe_with_jab",
        lambda actual_jab: (
            calls.append(actual_jab) or {"ok": True, "mode": "in-process"}
        ),
    )

    assert open_self_made_entry({"receipt_entry": {}}, jab) == {
        "ok": True,
        "mode": "in-process",
    }
    assert calls == [jab]


def test_extract_entry_dynamic_index_from_entry_button_path():
    report = {
        "entry_state": {
            "hits": [
                {
                    "control": {
                        "path": "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0",
                    }
                }
            ]
        }
    }

    assert extract_entry_dynamic_index(report) == 5


def test_extract_entry_dynamic_index_from_anchor_hit():
    report = {
        "entry_state": {
            "hits": [
                {
                    "control": {
                        "path": "0.0.1.0.0.0.0.5.0.0.0.1",
                        "dynamic_index": 7,
                    }
                }
            ]
        }
    }

    assert extract_entry_dynamic_index(report) == 7


def test_extract_entry_dynamic_index_prefers_customer_corrected_anchor_index():
    report = {
        "entry_state": {
            "hits": [
                {
                    "control": {
                        "path": "0.0.1.0.0.0.0.3.0.0.0.1",
                        "dynamic_index": 5,
                        "dynamic_prefix": "0.0.1.0.0.0.0.5",
                    }
                }
            ]
        },
        "anchor": {
            "mode": "current-canvas-anchor-corrected-by-customer",
            "initial_dynamic_index": 3,
            "dynamic_index": 5,
        },
    }

    assert extract_entry_dynamic_index(report) == 5


def test_extract_entry_scope_hwnd_from_windows_after_choose():
    report = {
        "parsed": {
            "windows_after_choose": [
                {
                    "is_java": True,
                    "visible": True,
                    "hwnd": 24680,
                    "class_name": "SunAwtCanvas",
                },
            ]
        }
    }

    assert extract_entry_scope_hwnd(report) == 24680


def test_extract_entry_scope_hwnd_prefers_self_made_canvas():
    frame_hwnd = 1001
    canvas_hwnd = 2002
    report = {
        "entry_state": {
            "hits": [
                {
                    "window": {
                        "hwnd": frame_hwnd,
                        "class_name": "SunAwtFrame",
                        "visible": True,
                    },
                    "control": {"path": "0.0.0.0.1.0.0.0.0.3.0.0"},
                },
                {
                    "window": {
                        "hwnd": canvas_hwnd,
                        "class_name": "SunAwtCanvas",
                        "visible": True,
                    },
                    "control": {"path": "0.0.1.0.0.0.0.3.0.0"},
                },
            ]
        }
    }

    assert extract_entry_scope_hwnd(report) == canvas_hwnd
