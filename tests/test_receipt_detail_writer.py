# 生命周期：持久维护
# 覆盖的业务场景：收款单明细 writer 只负责写入，不保留同步读回/重试旧逻辑
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_detail_writer.py

from core import receipt_detail_writer as writer
from core.receipt_detail_fields import DETAIL_FIELDS, FEE_FIELDS


def test_detail_settlement_fields_keep_keyboard_write_with_immediate_verify():
    settlement_fields = [
        field for field in DETAIL_FIELDS + FEE_FIELDS if field["name"] == "结算方式"
    ]

    assert len(settlement_fields) == 2
    for field in settlement_fields:
        assert field["commit_key"] == "Enter"
        assert field["input_mode"] == "paste"
        assert "edit_mode" not in field
        assert "activate_by_row_bounds_click" not in field
        assert "activate_by_double_click" not in field
        assert field["pre_write_stabilize"] is True
        assert field["pre_write_stabilize_wait"] == 0.08
        assert field["immediate_verify"] is True
        assert field["immediate_verify_attempts"] == 2
        assert field["immediate_verify_wait"] == 0.05


def test_amount_fields_keep_sensitive_neighbor_guard():
    amount_fields = [
        field
        for field in DETAIL_FIELDS + FEE_FIELDS
        if field["name"] == "贷方原币金额"
    ]

    assert len(amount_fields) == 2
    for field in amount_fields:
        assert field["kind"] == "amount"
        assert field["sensitive_neighbor_cols"] == [6]
        assert field["immediate_verify"] is True
        assert field["immediate_verify_attempts"] == 2
        assert field["immediate_verify_wait"] == 0.15


def test_subject_fields_verify_immediately_with_neighbor_guard():
    subject_fields = [
        field for field in DETAIL_FIELDS + FEE_FIELDS if field["name"] == "科目"
    ]

    assert len(subject_fields) == 2
    for field in subject_fields:
        assert field["kind"] == "code_prefix"
        assert field["sensitive_neighbor_cols"] == [4, 6, 7, 8]
        assert field["immediate_verify"] is True
        assert field["immediate_verify_attempts"] == 3
        assert field["immediate_verify_wait"] == 0.2


def test_business_type_fields_verify_immediately():
    business_type_fields = [
        field
        for field in DETAIL_FIELDS + FEE_FIELDS
        if field["name"] == "收款业务类型"
    ]

    assert len(business_type_fields) == 2
    for field in business_type_fields:
        assert field["immediate_verify"] is True
        assert field["immediate_verify_attempts"] == 2
        assert field["immediate_verify_wait"] == 0.05


def test_fee_account_clear_verifies_immediately():
    main_account = next(
        field
        for field in DETAIL_FIELDS
        if field["name"] == "收款银行账户"
        and field["value_key"] == "bank_account"
    )
    fee_account = next(
        field
        for field in FEE_FIELDS
        if field["name"] == "收款银行账户"
        and field["value_key"] == "fee_account"
    )

    assert main_account["immediate_verify"] is True
    assert main_account["immediate_verify_attempts"] == 3
    assert main_account["immediate_verify_wait"] == 0.2
    assert fee_account["kind"] == "blank"
    assert fee_account["immediate_verify"] is True
    assert fee_account["immediate_verify_attempts"] == 3
    assert fee_account["immediate_verify_wait"] == 0.2


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
        field=field,
        business={"bank_account": "FTE1219165931831"},
        attempt_no=1,
        current_col=None,
    )

    assert result["ok"] is True
    assert "selected_before_write" not in result
    assert result["commit_target"]["skipped"] is True
    assert "后台 verifier" in result["commit_target"]["reason"]
    assert set(result["stage_timing"]) >= {
        "focus_entry",
        "activation",
        "pre_write_stabilize",
        "neighbor_guard_before",
        "screen_write",
        "unaccounted",
    }


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
        field=field,
        business={"bank_account": "FTE1219165931831"},
        attempt_no=1,
        current_col=None,
        recover_after_failure=recover_after_failure,
    )

    assert result["ok"] is True
    assert hooks == [recover_after_failure]


def test_write_field_once_refocuses_target_cell_even_when_current_col_exists(monkeypatch):
    field = {"col": 7, "name": "贷方原币金额", "value_key": "amount"}
    focused = []

    def fake_focus(_jab, _located, row_index, col_index):
        focused.append((row_index, col_index))
        return {"ok": True, "target": {"row": row_index, "col": col_index}}

    def fail_if_arrow_navigation_is_used(*_args, **_kwargs):
        raise AssertionError("后续字段不应再通过方向键从上一列移动")

    monkeypatch.setattr(writer, "focus_detail_cell", fake_focus)
    monkeypatch.setattr(
        writer,
        "keyboard_write_selected_cell",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        writer,
        "move_selected_cell_by_arrows",
        fail_if_arrow_navigation_is_used,
        raising=False,
    )

    result = writer.write_field_once(
        jab=object(),
        located={"best": {"window": {"hwnd": 1}, "path": "0.1"}},
        table_window={"hwnd": 1},
        row_index=0,
        field=field,
        business={"amount": "375000.00"},
        attempt_no=1,
        current_col=6,
    )

    assert result["ok"] is True
    assert focused == [(0, 7)]
    assert result["navigation"]["skipped"] is True
    assert "path" in result["navigation"]["reason"]


def test_write_field_once_enters_bank_account_from_neighbor_col_without_mouse(monkeypatch):
    field = {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "bank_account",
        "focus_via_col": 5,
    }
    focused = []
    moves = []
    write_kwargs = []

    def fake_focus(_jab, _located, row_index, col_index):
        focused.append((row_index, col_index))
        return {"ok": True, "target": {"row": row_index, "col": col_index}}

    def fake_move(_table_window, from_col, to_col):
        moves.append((from_col, to_col))
        return {"ok": True, "from_col": from_col, "to_col": to_col, "key": "Left"}

    monkeypatch.setattr(writer, "focus_detail_cell", fake_focus)
    monkeypatch.setattr(writer, "move_selected_cell_by_arrows", fake_move)

    def fake_keyboard_write(*_args, **kwargs):
        write_kwargs.append(kwargs)
        return {"ok": True}

    monkeypatch.setattr(writer, "keyboard_write_selected_cell", fake_keyboard_write)

    result = writer.write_field_once(
        jab=object(),
        located={"best": {"window": {"hwnd": 1}, "path": "0.1"}},
        table_window={"hwnd": 1},
        row_index=0,
        field=field,
        business={"bank_account": "1783854003"},
        attempt_no=1,
        current_col=2,
    )

    assert result["ok"] is True
    assert focused == [(0, 5)]
    assert moves == [(5, 4)]
    assert result["navigation"]["via_col"] == 5
    assert result["target"] == {"row": 0, "col": 4}
    assert result["activation"]["method"] == "keyboard-only"


def test_write_field_once_uses_keyboard_only_activation(monkeypatch):
    field = {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "bank_account",
        "focus_via_col": 5,
    }
    monkeypatch.setattr(
        writer,
        "focus_detail_cell",
        lambda _jab, _located, row, col: {"ok": True, "target": {"row": row, "col": col}},
    )
    monkeypatch.setattr(
        writer,
        "move_selected_cell_by_arrows",
        lambda _window, from_col, to_col: {
            "ok": True,
            "from_col": from_col,
            "to_col": to_col,
        },
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
        field=field,
        business={"bank_account": "1783854003"},
        attempt_no=1,
    )

    assert result["ok"] is True
    assert result["activation"]["method"] == "keyboard-only"


def test_write_field_once_stabilizes_by_reusing_current_focus(monkeypatch):
    field = {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "bank_account",
        "focus_via_col": 5,
        "pre_write_stabilize": True,
        "pre_write_stabilize_wait": 0.0,
    }
    focused = []
    moves = []
    writes = []

    def fake_focus(_jab, _located, row_index, col_index):
        focused.append((row_index, col_index))
        return {"ok": True, "target": {"row": row_index, "col": col_index}}

    def fake_move(_table_window, from_col, to_col):
        moves.append((from_col, to_col))
        return {"ok": True, "from_col": from_col, "to_col": to_col, "key": "Left"}

    def fake_keyboard_write(_window, value, **_kwargs):
        writes.append(value)
        return {"ok": True, "commit": {"ok": True, "key": "Enter"}}

    monkeypatch.setattr(writer, "focus_detail_cell", fake_focus)
    monkeypatch.setattr(writer, "move_selected_cell_by_arrows", fake_move)
    monkeypatch.setattr(writer, "keyboard_write_selected_cell", fake_keyboard_write)

    result = writer.write_field_once(
        jab=object(),
        located={"best": {"window": {"hwnd": 1}, "path": "0.1"}},
        table_window={"hwnd": 1},
        row_index=0,
        field=field,
        business={"bank_account": "1783854003"},
        attempt_no=1,
    )

    assert result["ok"] is True
    assert result["pre_write_stabilize"]["ok"] is True
    assert result["pre_write_stabilize"]["method"] == "reuse-current-focus"
    assert "before" not in result["pre_write_stabilize"]
    assert "snapshot" not in result["pre_write_stabilize"]
    assert focused == [(0, 5)]
    assert moves == [(5, 4)]
    assert writes == ["1783854003"]


def test_write_field_once_can_force_refocus_during_stabilize(monkeypatch):
    field = {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "bank_account",
        "focus_via_col": 5,
        "pre_write_stabilize": True,
        "pre_write_stabilize_refocus": True,
        "pre_write_stabilize_wait": 0.0,
    }
    focused = []
    moves = []

    def fake_focus(_jab, _located, row_index, col_index):
        focused.append((row_index, col_index))
        return {"ok": True, "target": {"row": row_index, "col": col_index}}

    def fake_move(_table_window, from_col, to_col):
        moves.append((from_col, to_col))
        return {"ok": True, "from_col": from_col, "to_col": to_col, "key": "Left"}

    monkeypatch.setattr(writer, "focus_detail_cell", fake_focus)
    monkeypatch.setattr(writer, "move_selected_cell_by_arrows", fake_move)
    monkeypatch.setattr(
        writer,
        "keyboard_write_selected_cell",
        lambda *_args, **_kwargs: {"ok": True, "commit": {"ok": True, "key": "Enter"}},
    )

    result = writer.write_field_once(
        jab=object(),
        located={"best": {"window": {"hwnd": 1}, "path": "0.1"}},
        table_window={"hwnd": 1},
        row_index=0,
        field=field,
        business={"bank_account": "1783854003"},
        attempt_no=1,
    )

    assert result["ok"] is True
    assert result["pre_write_stabilize"]["method"] == "refocus-only"
    assert focused == [(0, 5), (0, 5)]
    assert moves == [(5, 4), (5, 4)]


def test_write_detail_line_immediately_verifies_bank_account(monkeypatch):
    field = {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "bank_account",
        "commit_key": "Enter",
        "edit_mode": "selected",
        "input_mode": "paste",
        "pre_commit_wait": 0.1,
        "immediate_verify": True,
        "immediate_verify_attempts": 2,
        "immediate_verify_wait": 0.0,
    }
    writes = []

    monkeypatch.setattr(
        writer,
        "focus_detail_cell",
        lambda _jab, _located, row, col: {
            "ok": True,
            "target": {"row": row, "col": col},
        },
    )

    def fake_keyboard_write(_window, value, **_kwargs):
        writes.append(value)
        return {"ok": True, "commit": {"ok": True, "key": "Enter"}}

    monkeypatch.setattr(writer, "keyboard_write_selected_cell", fake_keyboard_write)
    monkeypatch.setattr(
        writer,
        "read_row_cells",
        lambda *_args, **_kwargs: (
            {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
            {"4": "1783854003"},
        ),
    )

    steps = writer.write_detail_line_by_screen(
        jab=object(),
        business={"bank_account": "1783854003"},
        located={"best": {"window": {"hwnd": 1}, "row_count": 1, "col_count": 25}},
        fields=[field],
    )

    assert len(steps) == 1
    assert steps[0]["ok"] is True
    assert steps[0]["immediate_verify"]["ok"] is True
    assert steps[0]["immediate_verify"]["rewrites"] == []
    assert writes == ["1783854003"]


def test_write_detail_line_rewrites_bank_account_before_next_field(monkeypatch):
    bank_field = {
        "col": 4,
        "name": "收款银行账户",
        "value_key": "bank_account",
        "commit_key": "Enter",
        "edit_mode": "selected",
        "input_mode": "paste",
        "pre_commit_wait": 0.1,
        "focus_via_col": 5,
        "immediate_verify": True,
        "immediate_verify_attempts": 2,
        "immediate_verify_wait": 0.0,
    }
    subject_field = {
        "col": 5,
        "name": "科目",
        "value_key": "main_subject",
        "kind": "code_prefix",
        "input_mode": "paste",
    }
    focus_calls = []
    move_calls = []
    writes = []
    snapshots = iter(
        [
            (
                {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
                {"4": ""},
            ),
            (
                {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
                {"4": "1783854003"},
            ),
        ]
    )

    def fake_focus(_jab, _located, row, col):
        focus_calls.append((row, col))
        return {"ok": True, "target": {"row": row, "col": col}}

    def fake_keyboard_write(_window, value, **_kwargs):
        writes.append(value)
        return {"ok": True, "commit": {"ok": True, "key": "Enter"}}

    def fake_move(_window, from_col, to_col):
        move_calls.append((from_col, to_col))
        return {"ok": True, "from_col": from_col, "to_col": to_col, "key": "Left"}

    monkeypatch.setattr(writer, "focus_detail_cell", fake_focus)
    monkeypatch.setattr(writer, "move_selected_cell_by_arrows", fake_move)
    monkeypatch.setattr(writer, "keyboard_write_selected_cell", fake_keyboard_write)
    monkeypatch.setattr(writer, "read_row_cells", lambda *_args, **_kwargs: next(snapshots))

    steps = writer.write_detail_line_by_screen(
        jab=object(),
        business={"bank_account": "1783854003", "main_subject": "1002"},
        located={"best": {"window": {"hwnd": 1}, "row_count": 1, "col_count": 25}},
        fields=[bank_field, subject_field],
    )

    assert [step["name"] for step in steps] == ["收款银行账户", "科目"]
    assert all(step["ok"] for step in steps)
    assert steps[0]["immediate_verify"]["ok"] is True
    assert len(steps[0]["immediate_verify"]["rewrites"]) == 1
    assert writes == ["1783854003", "1783854003", "1002"]
    assert focus_calls == [(0, 5), (0, 5), (0, 5)]
    assert move_calls == [(5, 4), (5, 4)]


def test_immediate_verify_failure_keeps_diagnostic_cells(monkeypatch):
    field = {
        "col": 5,
        "name": "科目",
        "value_key": "main_subject",
        "kind": "code_prefix",
        "input_mode": "paste",
        "immediate_verify": True,
        "immediate_verify_attempts": 1,
        "immediate_verify_wait": 0.0,
    }

    monkeypatch.setattr(
        writer,
        "focus_detail_cell",
        lambda _jab, _located, row, col: {
            "ok": True,
            "target": {"row": row, "col": col},
        },
    )
    monkeypatch.setattr(
        writer,
        "keyboard_write_selected_cell",
        lambda *_args, **_kwargs: {"ok": True, "commit": {"ok": True}},
    )
    monkeypatch.setattr(
        writer,
        "read_row_cells",
        lambda *_args, **_kwargs: (
            {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
            {
                "4": "1783854003",
                "5": "",
                "6": "1.000000",
                "7": "",
                "8": "",
                "11": "",
            },
        ),
    )

    steps = writer.write_detail_line_by_screen(
        jab=object(),
        business={"main_subject": "1002"},
        located={"best": {"window": {"hwnd": 1}, "row_count": 1, "col_count": 25}},
        fields=[field],
    )

    assert steps[0]["ok"] is False
    verify = steps[0]["immediate_verify"]["verifications"][0]
    assert verify["diagnostic_cells"] == {
        "4": "1783854003",
        "5": "",
        "6": "1.000000",
        "7": "",
        "8": "",
        "11": "",
    }
    assert steps[0]["reason"] == steps[0]["immediate_verify"]["reason"]


def test_write_detail_line_rewrites_settlement_immediately(monkeypatch):
    settlement_field = {
        "col": 11,
        "name": "结算方式",
        "value_key": "settlement",
        "commit_key": "Enter",
        "input_mode": "paste",
        "immediate_verify": True,
        "immediate_verify_attempts": 2,
        "immediate_verify_wait": 0.0,
    }
    focus_calls = []
    writes = []
    snapshots = iter(
        [
            (
                {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
                {"11": ""},
            ),
            (
                {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
                {"11": "网银"},
            ),
        ]
    )

    def fake_focus(_jab, _located, row, col):
        focus_calls.append((row, col))
        return {"ok": True, "target": {"row": row, "col": col}}

    def fake_keyboard_write(_window, value, **_kwargs):
        writes.append(value)
        return {"ok": True, "commit": {"ok": True, "key": "Enter"}}

    monkeypatch.setattr(writer, "focus_detail_cell", fake_focus)
    monkeypatch.setattr(writer, "keyboard_write_selected_cell", fake_keyboard_write)
    monkeypatch.setattr(writer, "read_row_cells", lambda *_args, **_kwargs: next(snapshots))

    steps = writer.write_detail_line_by_screen(
        jab=object(),
        business={"settlement": "网银"},
        located={"best": {"window": {"hwnd": 1}, "row_count": 1, "col_count": 25}},
        fields=[settlement_field],
    )

    assert len(steps) == 1
    assert steps[0]["ok"] is True
    assert steps[0]["immediate_verify"]["ok"] is True
    assert len(steps[0]["immediate_verify"]["rewrites"]) == 1
    assert writes == ["网银", "网银"]
    assert focus_calls == [(0, 11), (0, 11)]


def test_write_detail_line_blocks_when_amount_neighbor_changes(monkeypatch):
    amount_field = {
        "col": 7,
        "name": "贷方原币金额",
        "value_key": "amount",
        "kind": "amount",
        "input_mode": "paste",
        "sensitive_neighbor_cols": [6],
        "immediate_verify": True,
        "immediate_verify_attempts": 1,
        "immediate_verify_wait": 0.0,
    }
    reads = iter(
        [
            (
                {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
                {"6": "1.000000", "7": "", "8": ""},
            ),
            (
                {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
                {"6": "375,000.00", "7": "375,000.00", "8": ""},
            ),
        ]
    )

    monkeypatch.setattr(
        writer,
        "focus_detail_cell",
        lambda _jab, _located, row, col: {
            "ok": True,
            "target": {"row": row, "col": col},
        },
    )
    monkeypatch.setattr(
        writer,
        "keyboard_write_selected_cell",
        lambda *_args, **_kwargs: {"ok": True, "commit": {"ok": True}},
    )
    monkeypatch.setattr(writer, "read_row_cells", lambda *_args, **_kwargs: next(reads))

    steps = writer.write_detail_line_by_screen(
        jab=object(),
        business={"amount": "375000.00"},
        located={"best": {"window": {"hwnd": 1}, "row_count": 1, "col_count": 25}},
        fields=[amount_field],
    )

    assert len(steps) == 1
    assert steps[0]["ok"] is False
    immediate = steps[0]["immediate_verify"]
    assert immediate["ok"] is False
    assert "敏感邻列" in immediate["reason"]
    verify = immediate["verifications"][0]
    assert verify["neighbor_guard_after"]["ok"] is False
    assert verify["neighbor_guard_after"]["changes"][0]["col"] == 6
    assert verify["diagnostic_cells"]["6"] == "375,000.00"
    assert verify["diagnostic_cells"]["7"] == "375,000.00"


def test_write_detail_line_allows_amount_when_only_local_amount_changes(monkeypatch):
    amount_field = {
        "col": 7,
        "name": "贷方原币金额",
        "value_key": "amount",
        "kind": "amount",
        "input_mode": "paste",
        "sensitive_neighbor_cols": [6],
        "immediate_verify": True,
        "immediate_verify_attempts": 1,
        "immediate_verify_wait": 0.0,
    }
    reads = iter(
        [
            (
                {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
                {"6": "1.000000", "7": "", "8": "0.00"},
            ),
            (
                {"ok": True, "fast_path": True, "row_count": 1, "col_count": 25},
                {"6": "1.000000", "7": "375,000.00", "8": "375,000.00"},
            ),
        ]
    )

    monkeypatch.setattr(
        writer,
        "focus_detail_cell",
        lambda _jab, _located, row, col: {
            "ok": True,
            "target": {"row": row, "col": col},
        },
    )
    monkeypatch.setattr(
        writer,
        "keyboard_write_selected_cell",
        lambda *_args, **_kwargs: {"ok": True, "commit": {"ok": True}},
    )
    monkeypatch.setattr(writer, "read_row_cells", lambda *_args, **_kwargs: next(reads))

    steps = writer.write_detail_line_by_screen(
        jab=object(),
        business={"amount": "375000.00"},
        located={"best": {"window": {"hwnd": 1}, "row_count": 1, "col_count": 25}},
        fields=[amount_field],
    )

    assert len(steps) == 1
    assert steps[0]["ok"] is True
    verify = steps[0]["immediate_verify"]["verifications"][0]
    assert verify["neighbor_guard_after"]["ok"] is True
    assert verify["neighbor_guard_after"]["changes"] == []


def test_keyboard_write_recovers_after_clipboard_open_failure(monkeypatch):
    from core import receipt_detail_screen_writer as screen

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
