# 生命周期：持久维护
# 覆盖的业务场景：收款单完整流程的单行明细 pipeline：明细写入校验/修复、表头锚点重试与回读诊断停机
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_detail_pipeline.py


from tests._receipt_full_flow_helpers import (
    open_report_with_header_anchor,
    plan_row,
    run_one_row,
)


def test_run_one_row_uses_detail_pipeline_verifier(monkeypatch):
    calls = {
        "start": 0,
        "field": [],
        "snapshot": [],
        "rows": [],
        "wait": [],
        "fill_header_kwargs": [],
        "body_locate_kwargs": [],
        "account_scope": [],
    }

    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class FakeVerifier:
        def __init__(self, config, located, flow_started_at=None, **kwargs):
            self.config = config
            self.located = located
            self.flow_started_at = flow_started_at
            self.kwargs = kwargs

        def start(self):
            calls["start"] += 1

        def submit_field(self, row_index, field, business):
            task_id = f"field-{len(calls['field'])}"
            calls["field"].append(
                (row_index, field["name"], business[field["value_key"]])
            )
            return task_id

        def submit_snapshot(
            self, label, max_rows=5, timeout=1.2, interval=0.08, min_matches=0
        ):
            task_id = f"snapshot-{len(calls['snapshot'])}"
            calls["snapshot"].append((label, max_rows, min_matches))
            return task_id

        def submit_row_count(self, expected_rows, timeout=1.1, interval=0.06):
            task_id = f"rows-{len(calls['rows'])}"
            calls["rows"].append(expected_rows)
            return task_id

        def wait(self, task_ids, timeout=2.0):
            calls["wait"].append((list(task_ids), timeout))
            return {"ok": True, "submitted": list(task_ids), "done": len(task_ids)}

        def close(self, timeout=1.0):
            calls["closed"] = timeout

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: {
            "ok": True,
            "entry_state": {
                "hits": [
                    {
                        "window": {
                            "hwnd": 2002,
                            "class_name": "SunAwtCanvas",
                            "visible": True,
                        },
                        "control": {
                            "path": "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0",
                            "name": "财务组织(O)",
                            "description": "财务组织(O)",
                        },
                    }
                ]
            },
        },
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", FakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda _jab, _business, **kwargs: (
            calls["fill_header_kwargs"].append(kwargs)
            or [
                {
                    "ok": True,
                    "label": "客户",
                    "value": "YW03574",
                    "accepted_text": "ACME LTD",
                }
            ]
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **kwargs: (
            calls["body_locate_kwargs"].append(kwargs)
            or {
                "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
                "candidates": [],
            }
        ),
    )

    def fail_sync_read_before(*_args, **_kwargs):
        raise AssertionError("完整流程不应在明细写入前同步读整表")

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.read_body_table",
        fail_sync_read_before,
    )

    def fake_write_detail(_jab, business, _located, after_field=None, **_kwargs):
        field = {"col": 4, "name": "收款银行账户", "value_key": "bank_account"}
        step = {"ok": True, "input_ok": True, "name": field["name"]}
        assert after_field is not None
        step["async_verify_task"] = after_field(0, field, business, step)
        return [step]

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        fake_write_detail,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_header_account_description",
        lambda _jab, _timeout=0.0, scope=None: (
            calls["account_scope"].append(scope) or {"accepted": True}
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", FakeVerifier
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("no-save 正常路径不应主动检查弹窗")
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True, (
        report.get("failed_step"),
        report.get("reason"),
        report.get("exception"),
        report.get("detail_pipeline_repair"),
        report.get("detail_pipeline_verify_after_repair"),
    )
    assert calls["start"] == 1
    assert calls["field"] == [(0, "收款银行账户", "FTE1219165931831")]
    assert calls["snapshot"] == [("after-main-line", 3, 1)]
    assert calls["rows"] == [1]
    assert calls["wait"] == [(["field-0", "rows-0"], 2.0)]
    assert calls["fill_header_kwargs"][0]["scope_hwnd"] == 2002
    assert calls["fill_header_kwargs"][0]["dynamic_index"] == 5
    # 业务意图：anchor_path 走 dynamic_index=5 的模块前缀。只断言前缀(锁模块)+锚点段，
    # 不再逐字符锁死整条 NC 树路径，避免树结构微调即碎。
    anchor_path = calls["fill_header_kwargs"][0]["anchor_path"]
    assert anchor_path.startswith("0.0.1.0.0.0.0.5.")
    assert anchor_path.endswith(".1.1.1.0")
    assert calls["body_locate_kwargs"][0]["scope_hwnd"] == 2002
    body_path = calls["body_locate_kwargs"][0]["cached"]["best"]["path"]
    assert body_path.startswith("0.0.1.0.0.0.0.5.")
    assert calls["body_locate_kwargs"][0]["cached"]["best"]["window"] == {
        "hwnd": 2002,
        "class_name": "SunAwtCanvas",
    }
    assert calls["account_scope"][0]["scope_hwnd"] == 2002
    assert calls["account_scope"][0]["dynamic_index"] == 5
    assert report["before_table"]["skipped"] is True
    assert report["after_table"]["skipped"] is True
    assert calls["closed"] == 0.2


def test_run_one_row_retries_current_canvas_header_anchor(monkeypatch):
    calls = {
        "anchor_retry": [],
        "fill_header_kwargs": [],
        "body_locate_kwargs": [],
    }

    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class FakeVerifier:
        def __init__(self, *_args, **_kwargs):
            pass

        def start(self):
            pass

        def submit_field(self, *_args, **_kwargs):
            return "field-0"

        def submit_snapshot(self, *_args, **_kwargs):
            return "snapshot-0"

        def submit_row_count(self, *_args, **_kwargs):
            return "rows-0"

        def wait(self, *_args, **_kwargs):
            return {"ok": True}

        def close(self, timeout=1.0):
            pass

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: {
            "ok": True,
            "entry_state": {
                "hits": [
                    {
                        "window": {
                            "hwnd": 919586,
                            "class_name": "SunAwtCanvas",
                            "visible": True,
                        },
                        "control": {"path": "not-a-header-path"},
                    }
                ]
            },
        },
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", FakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_receipt_header_anchor_in_current_canvas",
        lambda _jab, hwnd, timeout=1.2, interval=0.2: (
            calls["anchor_retry"].append((hwnd, timeout, interval))
            or {
                "ok": True,
                "scope_hwnd": hwnd,
                "dynamic_index": 5,
                "dynamic_prefix": "0.0.1.0.0.0.0.5",
            }
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda _jab, _business, **kwargs: (
            calls["fill_header_kwargs"].append(kwargs)
            or [{"ok": True, "label": "客户", "accepted_text": "ACME LTD"}]
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **kwargs: (
            calls["body_locate_kwargs"].append(kwargs)
            or {
                "best": {
                    "path": "0.1",
                    "row_count": 1,
                    "col_count": 25,
                    "window": {"hwnd": 919586},
                },
                "candidates": [],
            }
        ),
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
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", FakeVerifier
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True, report
    assert calls["anchor_retry"] == [(919586, 1.2, 0.2)]
    assert report["entry_dynamic_index"] == 5
    assert report["entry_header_anchor_retry"]["ok"] is True
    assert calls["fill_header_kwargs"][0]["scope_hwnd"] == 919586
    assert calls["fill_header_kwargs"][0]["dynamic_index"] == 5
    assert calls["body_locate_kwargs"][0]["scope_hwnd"] == 919586
    # 业务意图：缓存的 body 路径走 dynamic_index=5 模块前缀。只锁前缀即可验证模块路由，
    # 不再逐字符锁死整条 NC 树路径。
    assert (
        calls["body_locate_kwargs"][0]["cached"]["best"]["path"].startswith(
            "0.0.1.0.0.0.0.5."
        )
    )


def test_run_one_row_stops_when_current_canvas_header_anchor_missing(monkeypatch):
    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: {
            "ok": True,
            "entry_state": {
                "hits": [
                    {
                        "window": {
                            "hwnd": 919586,
                            "class_name": "SunAwtCanvas",
                            "visible": True,
                        },
                        "control": {"path": "not-a-header-path"},
                    }
                ]
            },
        },
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", FakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_receipt_header_anchor_in_current_canvas",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "当前 canvas 未找到财务组织(O) 锚点",
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("锚点失败后不应进入表头写入")
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("锚点失败后不应定位明细表")
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is False
    assert report["failed_step"] == "header-anchor"
    assert "不走语义兜底" in report["reason"]


def test_run_one_row_repairs_pending_detail_field_with_cached_path(monkeypatch):
    calls = {
        "field": [],
        "snapshot": [],
        "rows": [],
        "wait": [],
        "repair": [],
        "locate": 0,
    }
    located = {
        "best": {
            "path": "0.1",
            "row_count": 1,
            "col_count": 25,
            "window": {"hwnd": 2002},
        },
        "candidates": [],
    }

    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class RepairingVerifier:
        def __init__(self, config, actual_located, flow_started_at=None, **kwargs):
            self.config = config
            self.located = actual_located
            self.flow_started_at = flow_started_at
            self.kwargs = kwargs

        def start(self):
            pass

        def submit_field(self, row_index, field, business):
            task_id = f"field-{len(calls['field'])}"
            calls["field"].append(
                {
                    "task_id": task_id,
                    "row_index": row_index,
                    "name": field["name"],
                    "value": business[field["value_key"]],
                }
            )
            return task_id

        def submit_snapshot(
            self, label, max_rows=5, timeout=1.2, interval=0.08, min_matches=0
        ):
            task_id = f"snapshot-{len(calls['snapshot'])}"
            calls["snapshot"].append(
                {
                    "task_id": task_id,
                    "label": label,
                    "max_rows": max_rows,
                    "min_matches": min_matches,
                }
            )
            return task_id

        def submit_row_count(self, expected_rows, timeout=1.1, interval=0.06):
            task_id = f"rows-{len(calls['rows'])}"
            calls["rows"].append(expected_rows)
            return task_id

        def wait(self, task_ids, timeout=2.0):
            ids = list(task_ids)
            calls["wait"].append((ids, timeout))
            if len(calls["wait"]) == 1:
                return {
                    "ok": False,
                    "submitted": ["field-0", "rows-0"],
                    "done": 1,
                    "pending": 1,
                    "failed": [],
                    "results": {
                        "rows-0": {
                            "ok": True,
                            "type": "row_count",
                            "expected_rows": 1,
                            "actual_rows": 1,
                        }
                    },
                }
            return {
                "ok": True,
                "submitted": ["field-0", "rows-0", "field-1"],
                "done": 3,
                "pending": 0,
                "failed": [],
                "results": {
                    "rows-0": {"ok": True, "type": "row_count"},
                    "field-1": {
                        "ok": True,
                        "type": "field",
                        "name": "收款银行账户",
                    },
                },
            }

        def snapshot(self):
            return {"ok": len(calls["wait"]) >= 2}

        def close(self, timeout=1.0):
            calls["closed"] = timeout

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: {
            "ok": True,
            "entry_state": {
                "hits": [
                    {
                        "window": {
                            "hwnd": 2002,
                            "class_name": "SunAwtCanvas",
                            "visible": True,
                        },
                        "control": {
                            "path": "0.0.1.0.0.0.0.5.0.0",
                            "dynamic_index": 5,
                        },
                    }
                ]
            },
        },
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", FakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda *_args, **_kwargs: [
            {"ok": True, "label": "客户", "accepted_text": "ACME LTD"}
        ],
    )

    def fake_locate(*_args, **_kwargs):
        calls["locate"] += 1
        return located

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        fake_locate,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.read_body_table",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("修复成功后不应整表读 fallback")
        ),
    )

    def fake_write_detail(_jab, business, actual_located, after_field=None, **_kwargs):
        assert actual_located["best"]["path"] == located["best"]["path"]
        assert actual_located["best"]["window"] == located["best"]["window"]
        field = {"col": 4, "name": "收款银行账户", "value_key": "bank_account"}
        assert after_field is not None
        return [
            {
                "ok": True,
                "input_ok": True,
                "name": field["name"],
                "async_verify_task": after_field(
                    0,
                    field,
                    business,
                    {"ok": True, "name": field["name"]},
                ),
            }
        ]

    def fake_write_field_once(
        _jab,
        actual_located,
        table_window,
        row_index,
        row_count,
        field,
        business,
        attempt_no,
        current_col=None,
        recover_after_failure=None,
    ):
        calls["repair"].append(
            {
                "path": actual_located["best"]["path"],
                "table_window": table_window,
                "row_index": row_index,
                "row_count": row_count,
                "field": field["name"],
                "value": business[field["value_key"]],
                "attempt_no": attempt_no,
                "current_col": current_col,
                "has_recover_hook": recover_after_failure is not None,
            }
        )
        return {
            "ok": True,
            "input_ok": True,
            "commit_ok": True,
            "commit_col": field["col"],
            "target": {"row": row_index, "col": field["col"]},
        }

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        fake_write_detail,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_field_once",
        fake_write_field_once,
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
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", RepairingVerifier
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
        lambda *_args, **_kwargs: {"ok": False, "attempted": False},
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True, (
        report.get("failed_step"),
        report.get("reason"),
        report.get("exception"),
        report.get("detail_pipeline_repair"),
        report.get("detail_pipeline_verify_after_repair"),
    )
    assert calls["locate"] == 1
    assert calls["field"] == [
        {
            "task_id": "field-0",
            "row_index": 0,
            "name": "收款银行账户",
            "value": "FTE1219165931831",
        },
        {
            "task_id": "field-1",
            "row_index": 0,
            "name": "收款银行账户",
            "value": "FTE1219165931831",
        },
    ]
    assert calls["snapshot"] == [
        {
            "task_id": "snapshot-0",
            "label": "after-main-line",
            "max_rows": 3,
            "min_matches": 1,
        },
        {
            "task_id": "snapshot-1",
            "label": "after-detail-repair",
            "max_rows": 3,
            "min_matches": 1,
        },
    ]
    assert calls["rows"] == [1]
    assert calls["wait"] == [
        (["field-0", "rows-0"], 2.0),
        (["field-1"], 2.0),
    ]
    assert calls["repair"] == [
        {
            "path": "0.1",
            "table_window": {"hwnd": 2002},
            "row_index": 0,
            "row_count": 1,
            "field": "收款银行账户",
            "value": "FTE1219165931831",
            "attempt_no": 2,
            "current_col": None,
            "has_recover_hook": True,
        }
    ]
    assert report["detail_pipeline_repair"]["ok"] is True
    assert report["detail_pipeline_verify_after_repair"]["ok"] is True
    assert report["after_table"]["skipped"] is True


def test_run_one_row_stops_when_customer_name_readback_is_empty(monkeypatch):
    class LocalFakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class LocalFakeVerifier:
        def __init__(self, config, located, flow_started_at=None, **_kwargs):
            self.config = config
            self.located = located
            self.flow_started_at = flow_started_at

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

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", LocalFakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda _jab, _business, **_kwargs: [
            {"ok": True, "label": "客户", "value": "YW03574"}
        ],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.read_body_table",
        lambda _jab, step: {"ok": True, "step": step, "rows": []},
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
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("no-save 正常路径不应主动检查弹窗")
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is False
    assert report["failed_step"] == "header-customer-name"
    assert report["nc_customer_name"] == ""
    assert "客户名称未确认" in report["reason"]


def test_run_one_row_continues_when_header_account_readback_is_empty(monkeypatch):
    account_readback_timeouts = []

    class LocalFakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class LocalFakeVerifier:
        def __init__(self, config, located, flow_started_at=None, **_kwargs):
            self.config = config
            self.located = located
            self.flow_started_at = flow_started_at

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

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", LocalFakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda _jab, _business, **_kwargs: [
            {
                "ok": True,
                "label": "客户",
                "value": "YW03574",
                "accepted_text": "ACME LTD",
            }
        ],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.read_body_table",
        lambda _jab, step: {"ok": True, "step": step, "rows": []},
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
        lambda _jab, timeout=5.0, **_kwargs: (
            account_readback_timeouts.append(timeout)
            or {"accepted": False, "description": "", "text": ""}
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", LocalFakeVerifier
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("no-save 正常路径不应主动检查弹窗")
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True
    assert "header_account_readback_warning" in report
    assert account_readback_timeouts == [0.0]


def test_pause_after_customer_diagnoses_cleared_header_and_stops(monkeypatch):
    class LocalFakeInfo:
        name = ""
        description = ""

    class LocalFakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

        def get_context_info(self, vm_id, context):
            return LocalFakeInfo()

        def get_text_context_value(self, vm_id, context):
            return ""

        def release_contexts(self, vm_id, contexts):
            return None

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", LocalFakeJAB)
    monkeypatch.setattr("builtins.input", lambda _prompt: "")
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.find_receipt_header_field_by_dynamic_path",
        lambda _jab, label, dynamic_index, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [object()],
            "path": f"path-{label}",
            "dynamic_prefix": "0.0.1.0.0.0.0.5",
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("表头诊断失败后不应继续定位明细表")
        ),
    )

    def fake_fill_header(_jab, _business, after_field=None, **_kwargs):
        assert after_field is not None
        steps = [
            {
                "ok": True,
                "label": "财务组织",
                "value": "A001",
                "dynamic_index": 5,
                "dynamic_prefix": "0.0.1.0.0.0.0.5",
                "path": "path-finance",
            },
            {
                "ok": True,
                "label": "客户",
                "value": "YW03574",
                "dynamic_index": 5,
                "dynamic_prefix": "0.0.1.0.0.0.0.5",
                "path": "path-customer",
            },
        ]
        after_field("财务组织", "A001", steps[0])
        callback = after_field("客户", "YW03574", steps[1])
        steps[1]["after_field_callback"] = callback
        if callback and not callback.get("ok", True):
            steps.append(
                {
                    "step": "blocked",
                    "ok": False,
                    "label": "客户",
                    "reason": callback["reason"],
                }
            )
        return steps

    monkeypatch.setattr("tools.receipt_full_flow_entry.fill_header", fake_fill_header)

    report = run_one_row(
        {},
        plan_row(10),
        save_enabled=False,
        pause_after_header_field="客户",
        diagnose_header_after_pause=True,
    )

    assert report["ok"] is False
    assert report["failed_step"] == "header-fill"
    diagnostics = report["header_pause_diagnostics"][0]
    assert diagnostics["ok"] is False
    assert [item["label"] for item in diagnostics["header_readback"]] == [
        "财务组织",
        "客户",
    ]
    assert all(item["present"] is False for item in diagnostics["header_readback"])
