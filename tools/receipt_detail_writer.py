# 职责：按字段映射写入收款单明细行，并通过表格读回校验
# 不做什么：不增删明细行，不读取 Excel，不处理 CLI/打印
# 允许依赖层：tools.receipt_detail_fields/reader/screen_writer、tools.receipt_body_table_locator
# 谁不应该 import：配置校验、Sheet 写入、收款匹配模块不应 import

import time

from tools.receipt_body_table_locator import locate_receipt_body_table_cached
from tools.receipt_detail_fields import (
    DETAIL_FIELDS,
    apply_readback_to_steps,
    field_mismatch_reason,
    field_matches,
    make_detail_step,
)
from tools.receipt_detail_reader import read_row_cells
from tools.receipt_detail_screen_writer import (
    KEYBOARD_INPUT_COMMIT_KEY,
    focus_detail_cell,
    keyboard_write_selected_cell,
    move_selected_cell_by_arrows,
    read_selected_cell,
)
from tools.receipt_keyboard_utils import STOP_HOTKEY, is_stop_hotkey_pressed

MAX_FIELD_RETRIES = 3


def write_field_once(
    jab,
    located,
    table_window,
    row_index,
    row_count,
    field,
    next_col,
    business,
    attempt_no,
    current_col=None,
):
    value = str(business[field["value_key"]])
    attempt_start = time.perf_counter()
    target_col = int(field["col"])
    if current_col is None:
        focus = focus_detail_cell(
            jab,
            located,
            row_index,
            target_col,
        )
        navigation = {
            "ok": True,
            "skipped": True,
            "reason": "首个字段直接定位",
            "from_col": target_col,
            "to_col": target_col,
        }
    else:
        focus = {"ok": True, "skipped": True, "reason": "沿用当前选中单元格"}
        navigation = move_selected_cell_by_arrows(table_window, current_col, target_col)
    selected_before = None
    if focus.get("ok") and navigation.get("ok"):
        selected_before = read_selected_cell(jab, located)
        selected = selected_before.get("single") if selected_before.get("ok") else None
        selected_row = selected.get("row") if selected else None
        selected_col = selected.get("col") if selected else None
        if not (
            selected
            and selected_row is not None
            and selected_col is not None
            and int(selected_row) == int(row_index)
            and int(selected_col) == target_col
        ):
            navigation = {
                **navigation,
                "ok": False,
                "selected_before_write": selected_before,
                "reason": (
                    f"方向键导航后当前格不匹配：期望第 {row_index + 1} 行第 {target_col} 列，"
                    f"实际 {selected!r}"
                ),
            }
    if focus.get("ok") and navigation.get("ok"):
        commit_key = field.get("commit_key") or KEYBOARD_INPUT_COMMIT_KEY
        screen = keyboard_write_selected_cell(
            table_window,
            value,
            commit_key=commit_key,
            clear_only=field.get("kind") == "blank",
            accept_key=field.get("accept_key"),
            typing_interval=field.get("typing_interval", 0.0),
            edit_mode=field.get("edit_mode", "editor"),
            input_mode=field.get("input_mode", "paste"),
            pre_commit_wait=field.get("pre_commit_wait", 0.025),
        )
    else:
        screen = {
            "ok": False,
            "reason": focus.get("reason") or navigation.get("reason"),
            "focus": focus,
            "navigation": navigation,
        }
    selected_after = None
    if screen.get("ok"):
        selected_after = {"ok": True, "skipped": True, "reason": "最终统一读回校验"}
    return {
        "attempt": attempt_no,
        "seconds": round(time.perf_counter() - attempt_start, 3),
        "mode": "keyboard",
        "input_ok": bool(screen.get("ok")),
        "input_reason": screen.get("reason"),
        "target": {"row": row_index, "col": target_col},
        "table_bounds": None,
        "cell_width": None,
        "cell_height": None,
        "focus": focus,
        "navigation": navigation,
        "selected_before_write": selected_before,
        "commit_ok": bool(screen.get("ok")),
        "commit_key": field.get("commit_key") or KEYBOARD_INPUT_COMMIT_KEY,
        "accept_key": field.get("accept_key"),
        "typing_interval": field.get("typing_interval", 0.0),
        "edit_mode": field.get("edit_mode", "editor"),
        "input_mode": field.get("input_mode", "paste"),
        "pre_commit_wait": field.get("pre_commit_wait", 0.025),
        "commit_col": current_col_after_commit(target_col, field.get("commit_key")),
        "commit_target": selected_after,
        "commit_reason": screen.get("reason"),
        "ok": bool(screen.get("ok")),
    }


def current_col_after_commit(target_col, commit_key):
    key = commit_key or KEYBOARD_INPUT_COMMIT_KEY
    if str(key).lower() in {"enter", "right"}:
        return int(target_col) + 1
    if str(key).lower() == "left":
        return max(int(target_col) - 1, 0)
    return int(target_col)


def refresh_unmatched_settlement_steps(
    jab, steps, row_index, timeout=0.35, interval=0.07, located=None
):
    if not any(step.get("name") == "结算方式" and not step.get("ok") for step in steps):
        return None
    deadline = time.perf_counter() + timeout
    last_snapshot = None
    last_cells = None
    while True:
        snapshot, cells = read_row_cells(jab, row_index, located)
        last_snapshot = snapshot
        last_cells = cells
        if snapshot.get("ok"):
            apply_readback_to_steps(steps, cells)
            if all(step.get("ok") for step in steps if step.get("name") == "结算方式"):
                return {
                    "ok": True,
                    "seconds": round(
                        timeout - max(deadline - time.perf_counter(), 0), 3
                    ),
                    "snapshot": snapshot,
                }
        if time.perf_counter() >= deadline:
            return {
                "ok": False,
                "seconds": timeout,
                "snapshot": last_snapshot,
                "cells": last_cells,
            }
        time.sleep(interval)


def write_detail_line_by_screen(
    jab,
    business,
    located,
    fields=None,
    row_index=0,
    after_field=None,
):
    fields = fields or DETAIL_FIELDS
    best = located.get("best") or {}
    table_window = best.get("window") or {}
    table_bounds = best.get("bounds")
    col_count = int(best.get("col_count") or 0)
    row_count = int(best.get("row_count") or 0)
    if row_count <= row_index or col_count < 25:
        return [
            {
                "ok": False,
                "changed": False,
                "partial_success": False,
                "name": "明细表",
                "reason": f"明细表尺寸异常：{row_count} 行 x {col_count} 列，目标第 {row_index + 1} 行",
            }
        ]

    steps = []
    current_col = None
    for index, field in enumerate(fields):
        if is_stop_hotkey_pressed():
            steps.append(
                {
                    "ok": False,
                    "changed": False,
                    "partial_success": bool(steps),
                    "name": field["name"],
                    "value": business[field["value_key"]],
                    "reason": f"检测到紧急停止键 {STOP_HOTKEY}",
                }
            )
            break

        step = make_detail_step(field, business, row_index, row_count, col_count)
        next_field = fields[index + 1] if index + 1 < len(fields) else fields[0]
        attempt = write_field_once(
            jab,
            located,
            table_window,
            row_index,
            row_count,
            field,
            next_field["col"],
            business,
            attempt_no=1,
            current_col=current_col,
        )
        if attempt.get("ok"):
            commit_col = attempt.get("commit_col")
            if commit_col is not None:
                current_col = int(commit_col)
        step["attempts"].append(attempt)
        step["input_ok"] = bool(attempt.get("input_ok"))
        step["target"] = attempt.get("target")
        step["commit_click"] = {
            "ok": attempt.get("commit_ok"),
            "target": attempt.get("commit_target"),
            "reason": attempt.get("commit_reason"),
        }
        step["geometry"].update(
            {
                "table_bounds": attempt.get("table_bounds") or table_bounds,
                "cell_width": attempt.get("cell_width"),
                "cell_height": attempt.get("cell_height"),
            }
        )
        if after_field and attempt.get("ok"):
            step["async_verify_task"] = after_field(row_index, field, business, step)
        if not attempt.get("ok"):
            step["reason"] = attempt.get("input_reason") or attempt.get("commit_reason")
        steps.append(step)
        if not attempt.get("ok"):
            break
    else:
        if after_field:
            for step in steps:
                step["ok"] = bool(step.get("input_ok"))
                step["blocked"] = not step["ok"]
                step["reason"] = None if step["ok"] else step.get("reason")
                step["actual"] = None
                step["deferred_readback"] = {
                    "ok": True,
                    "reason": "后台 pipeline verifier 批量读回校验",
                }
        else:
            _snapshot, cells = read_row_cells(jab, row_index, located)
            apply_readback_to_steps(steps, cells)
            settle_refresh = refresh_unmatched_settlement_steps(
                jab,
                steps,
                row_index,
                located=located,
            )
            if settle_refresh:
                for step in steps:
                    if step.get("name") == "结算方式":
                        step["settlement_stability_check"] = settle_refresh

        for step in steps:
            while (
                not step.get("ok")
                and len(step.get("attempts") or []) < MAX_FIELD_RETRIES
            ):
                field = next(item for item in fields if item["col"] == step["col"])
                field_index = fields.index(field)
                next_field = (
                    fields[field_index + 1]
                    if field_index + 1 < len(fields)
                    else fields[0]
                )
                refreshed = locate_receipt_body_table_cached(
                    jab,
                    cached=located,
                    max_rows=max(5, row_count),
                    scope_hwnd=(table_window or {}).get("hwnd"),
                )
                attempt = write_field_once(
                    jab,
                    refreshed,
                    table_window,
                    row_index,
                    row_count,
                    field,
                    next_field["col"],
                    business,
                    attempt_no=len(step["attempts"]) + 1,
                    current_col=None,
                )
                step["attempts"].append(attempt)
                step["input_ok"] = bool(attempt.get("input_ok"))
                step["target"] = attempt.get("target")
                step["commit_click"] = {
                    "ok": attempt.get("commit_ok"),
                    "target": attempt.get("commit_target"),
                    "reason": attempt.get("commit_reason"),
                }
                _snapshot, cells = read_row_cells(jab, row_index, refreshed)
                actual = cells.get(str(step["col"]))
                step["actual"] = actual
                ok = bool(attempt.get("ok")) and field_matches(
                    actual, step.get("raw_value") or step["value"], step.get("kind")
                )
                step["ok"] = ok
                step["blocked"] = not ok
                step["reason"] = (
                    None
                    if ok
                    else (
                        attempt.get("input_reason")
                        or attempt.get("commit_reason")
                        or field_mismatch_reason(step, actual, "修复后校验失败")
                    )
                )
            if not step.get("ok"):
                break
    any_changed = any(
        attempt.get("input_ok")
        for step in steps
        for attempt in step.get("attempts", [])
    )
    all_ok = bool(steps) and all(bool(step.get("ok")) for step in steps)
    for step in steps:
        step["changed"] = any(
            bool(attempt.get("input_ok")) for attempt in step.get("attempts", [])
        )
        step["partial_success"] = bool(any_changed and not all_ok)
        step["effect_policy"] = (
            "字段级保留已成功写入；失败后返回 steps，调用方按 partial_success 判断是否人工核对"
        )
    return steps
