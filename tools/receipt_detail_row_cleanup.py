# 职责：执行收款单明细多余行删除和手续费账户清空规则
# 不做什么：不新增手续费行，不定义字段映射，不负责 CLI/打印
# 允许依赖层：tools.receipt_detail_fields/reader/screen_writer、tools.receipt_body_table_locator
# 谁不应该 import：配置校验、Sheet 写入、收款匹配模块不应 import

import os
import time

from tools.receipt_body_table_locator import locate_receipt_body_table_cached
from tools.receipt_detail_fields import (
    ACCOUNT_COL,
    AMOUNT_COL,
    SUBJECT_COL,
    normalize_amount_text,
    normalize_text,
)
from tools.receipt_detail_reader import (
    read_located_body_table,
    read_row_cells,
    read_table_row_count_by_path,
    wait_body_row_count,
    wait_table_row_count_by_path,
)
from tools.receipt_detail_screen_writer import (
    focus_detail_cell,
    keyboard_write_selected_cell,
)
from tools.receipt_keyboard_utils import guarded_send_ctrl_d


def skip_fee_extra_row_delete_enabled():
    return os.environ.get("RECEIPT_SKIP_FEE_EXTRA_ROW_DELETE", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "y",
    }


def read_fee_prepare_row_count(jab, located, scope_hwnd=None):
    fast = read_table_row_count_by_path(jab, located)
    if fast.get("ok"):
        return {
            "ok": True,
            "fast_path": True,
            "row_count": fast.get("row_count"),
            "source": "read_table_row_count_by_path",
        }
    fallback = read_located_body_table(
        jab,
        located,
        "before_fee_row_prepare",
        scope_hwnd=scope_hwnd,
    )
    fallback["fast_path"] = False
    fallback["fast_reason"] = fast.get("reason")
    return fallback


def _with_delete_effect(result, expected_rows=None):
    steps = result.get("steps") or []
    before_rows = int(result.get("before_rows") or result.get("after_rows") or 0)
    after_rows = int(result.get("after_rows") or before_rows)
    changed = after_rows < before_rows or any(step.get("ok") for step in steps)
    result["changed"] = changed
    result["partial_success"] = bool(changed and not result.get("ok"))
    if result.get("skipped"):
        result["effect_policy"] = "未删行；行数已满足目标"
    elif result.get("ok"):
        result["effect_policy"] = (
            "已按最后一行向前删除；删除不可回滚，成功后保留当前行数"
        )
    else:
        result["effect_policy"] = (
            "删行失败时保留已删部分；调用方必须用 before_rows/after_rows/steps 判定人工修复范围"
        )
    if expected_rows is not None:
        result["expected_rows"] = expected_rows
    return result


def guard_extra_row_deletable(jab, located, row_index):
    snapshot, cells = read_row_cells(jab, row_index, located)
    if not snapshot.get("ok"):
        return {
            "ok": False,
            "reason": f"删除前无法读取第 {row_index + 1} 行：{snapshot.get('reason')}",
            "snapshot": snapshot,
        }
    subject = normalize_text(cells.get(str(SUBJECT_COL)))
    amount = normalize_amount_text(cells.get(str(AMOUNT_COL)))
    safe = not subject and amount in ("", "0.00")
    return {
        "ok": safe,
        "row_index": row_index,
        "cells": cells,
        "snapshot": snapshot,
        "subject": subject,
        "amount": amount,
        "reason": None
        if safe
        else (
            f"第 {row_index + 1} 行已有业务内容，禁止删除："
            f"科目={subject!r}，金额={amount!r}"
        ),
    }


def cleanup_rows_after_first(jab, located, scope_hwnd=None):
    started_at = time.perf_counter()
    before = read_located_body_table(
        jab,
        located,
        "before_cleanup_rows_after_first",
        scope_hwnd,
    )
    if not before.get("ok"):
        return {
            "ok": False,
            "reason": f"清理多余行前无法读取明细表：{before.get('reason')}",
            "before": before,
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    before_rows = int(before.get("row_count") or 0)
    if before_rows <= 1:
        return _with_delete_effect(
            {
                "ok": True,
                "skipped": True,
                "reason": "当前只有主行，无需清理多余行",
                "before_rows": before_rows,
                "after_rows": before_rows,
                "steps": [],
                "seconds": round(time.perf_counter() - started_at, 3),
            },
            expected_rows=1,
        )

    steps = []
    current_rows = before_rows
    while current_rows > 1:
        step_started_at = time.perf_counter()
        refreshed = locate_receipt_body_table_cached(
            jab,
            cached=located,
            max_rows=5,
            scope_hwnd=scope_hwnd
            or ((located.get("best") or {}).get("window") or {}).get("hwnd"),
        )
        best = refreshed.get("best") or {}
        table_window = best.get("window") or {}
        target_row = current_rows - 1
        focused = focus_detail_cell(jab, refreshed, target_row, 1)
        if not focused.get("ok"):
            return _with_delete_effect(
                {
                    "ok": False,
                    "reason": focused.get("reason"),
                    "before_rows": before_rows,
                    "after_rows": current_rows,
                    "steps": steps,
                    "focused": focused,
                    "seconds": round(time.perf_counter() - started_at, 3),
                },
                expected_rows=1,
            )
        guard = guard_extra_row_deletable(jab, refreshed, target_row)
        if not guard.get("ok"):
            return _with_delete_effect(
                {
                    "ok": False,
                    "reason": guard.get("reason"),
                    "before_rows": before_rows,
                    "after_rows": current_rows,
                    "steps": steps,
                    "guard": guard,
                    "seconds": round(time.perf_counter() - started_at, 3),
                },
                expected_rows=1,
            )
        sent = guarded_send_ctrl_d(table_window)
        waited = wait_body_row_count(
            jab,
            refreshed,
            expected_rows=current_rows - 1,
            label="after_cleanup_one_extra_row",
            scope_hwnd=scope_hwnd,
        )
        after = waited.get("snapshot") or {}
        after_rows = int(after.get("row_count") or 0)
        step = _delete_row_step(
            target_row,
            current_rows,
            after_rows,
            step_started_at,
            focused,
            sent,
            waited,
        )
        steps.append(step)
        if not step["ok"]:
            return _with_delete_effect(
                {
                    "ok": False,
                    "reason": step["reason"],
                    "before_rows": before_rows,
                    "after_rows": after_rows,
                    "steps": steps,
                    "seconds": round(time.perf_counter() - started_at, 3),
                },
                expected_rows=1,
            )
        current_rows = after_rows

    return _with_delete_effect(
        {
            "ok": True,
            "skipped": False,
            "reason": f"已删除第 1 行以外的多余行：{before_rows} -> 1",
            "before_rows": before_rows,
            "after_rows": current_rows,
            "steps": steps,
            "seconds": round(time.perf_counter() - started_at, 3),
        },
        expected_rows=1,
    )


def clear_fee_account_if_filled(jab, located, row_index, known_cells=None):
    snapshot = None
    cells = known_cells or {}
    before = normalize_text(cells.get(str(ACCOUNT_COL)))
    if known_cells is None or str(ACCOUNT_COL) not in cells:
        snapshot, cells = read_row_cells(jab, row_index, located)
        before = normalize_text(cells.get(str(ACCOUNT_COL)))
        if not snapshot.get("ok"):
            return {
                "ok": False,
                "reason": f"清空前无法读取手续费行：{snapshot.get('reason')}",
            }
    if not before:
        return {
            "ok": True,
            "skipped": True,
            "before": before,
            "after": before,
            "source": "known_cells" if known_cells is not None else "read_row_cells",
            "reason": "手续费账户列为空，不切换焦点",
        }

    best = located.get("best") or {}
    table_window = best.get("window") or {}
    focused = focus_detail_cell(jab, located, row_index, ACCOUNT_COL)
    if not focused.get("ok"):
        return {
            "ok": False,
            "before": before,
            "reason": focused.get("reason"),
            "focused": focused,
        }

    sent = keyboard_write_selected_cell(table_window, "", clear_only=True)
    _after_snapshot, after_cells = read_row_cells(jab, row_index, located)
    after = normalize_text(after_cells.get(str(ACCOUNT_COL)))
    return {
        "ok": bool(sent.get("ok")) and not after,
        "before": before,
        "after": after,
        "focused": focused,
        "sent": sent,
        "reason": None
        if bool(sent.get("ok")) and not after
        else sent.get("reason") or "Delete 后账户列仍非空",
    }


def delete_extra_row_if_present(
    jab,
    located,
    expected_rows,
    scope_hwnd=None,
    known_row_count=None,
    defer_wait=False,
):
    if known_row_count is None:
        fast_count = read_table_row_count_by_path(jab, located)
        if (
            fast_count.get("ok")
            and int(fast_count.get("row_count") or 0) <= expected_rows
        ):
            row_count = int(fast_count.get("row_count") or 0)
            return _with_delete_effect(
                {
                    "ok": True,
                    "skipped": True,
                    "fast_path": True,
                    "before_rows": row_count,
                    "after_rows": row_count,
                    "steps": [],
                    "seconds": 0.0,
                    "reason": f"当前 {row_count} 行，无需删行，不切换焦点",
                },
                expected_rows=expected_rows,
            )
    fast = fast_delete_extra_rows_by_row_count(
        jab,
        located,
        expected_rows,
        known_row_count=known_row_count,
        defer_wait=defer_wait,
    )
    if fast.get("ok") or fast.get("skipped"):
        return fast

    started_at = time.perf_counter()
    before = read_located_body_table(
        jab,
        located,
        "before_extra_row_delete",
        scope_hwnd,
    )
    if not before.get("ok"):
        return {
            "ok": False,
            "reason": f"删行前无法读取明细表：{before.get('reason')}",
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    before_rows = int(before.get("row_count") or 0)
    if before_rows <= expected_rows:
        return _with_delete_effect(
            {
                "ok": True,
                "skipped": True,
                "before_rows": before_rows,
                "after_rows": before_rows,
                "seconds": round(time.perf_counter() - started_at, 3),
            },
            expected_rows=expected_rows,
        )

    steps = []
    current_rows = before_rows
    while current_rows > expected_rows:
        step_started_at = time.perf_counter()
        refreshed = locate_receipt_body_table_cached(
            jab,
            cached=located,
            max_rows=max(5, current_rows),
            scope_hwnd=scope_hwnd
            or ((located.get("best") or {}).get("window") or {}).get("hwnd"),
        )
        best = refreshed.get("best") or {}
        table_window = best.get("window") or {}
        target_row = current_rows - 1
        focused = focus_detail_cell(jab, refreshed, target_row, 1)
        if not focused.get("ok"):
            return _with_delete_effect(
                {
                    "ok": False,
                    "before_rows": before_rows,
                    "after_rows": current_rows,
                    "reason": focused.get("reason"),
                    "steps": steps,
                    "focused": focused,
                    "seconds": round(time.perf_counter() - started_at, 3),
                },
                expected_rows=expected_rows,
            )
        guard = guard_extra_row_deletable(jab, refreshed, target_row)
        if not guard.get("ok"):
            return _with_delete_effect(
                {
                    "ok": False,
                    "before_rows": before_rows,
                    "after_rows": current_rows,
                    "reason": guard.get("reason"),
                    "steps": steps,
                    "guard": guard,
                    "seconds": round(time.perf_counter() - started_at, 3),
                },
                expected_rows=expected_rows,
            )
        sent = guarded_send_ctrl_d(table_window)
        waited = wait_body_row_count(
            jab,
            refreshed,
            expected_rows=current_rows - 1,
            label="after_extra_row_delete",
            scope_hwnd=scope_hwnd,
        )
        after = waited.get("snapshot") or {}
        after_rows = int(after.get("row_count") or 0)
        step = _delete_row_step(
            target_row,
            current_rows,
            after_rows,
            step_started_at,
            focused,
            sent,
            waited,
        )
        steps.append(step)
        if not step["ok"]:
            return _with_delete_effect(
                {
                    "ok": False,
                    "before_rows": before_rows,
                    "after_rows": after_rows,
                    "reason": step["reason"],
                    "steps": steps,
                    "seconds": round(time.perf_counter() - started_at, 3),
                },
                expected_rows=expected_rows,
            )
        current_rows = after_rows

    return _with_delete_effect(
        {
            "ok": current_rows == expected_rows,
            "before_rows": before_rows,
            "after_rows": current_rows,
            "steps": steps,
            "seconds": round(time.perf_counter() - started_at, 3),
            "reason": None
            if current_rows == expected_rows
            else f"删行后行数未回到 {expected_rows}，实际 {current_rows}",
        },
        expected_rows=expected_rows,
    )


def fast_delete_extra_rows_by_row_count(
    jab,
    located,
    expected_rows,
    known_row_count=None,
    defer_wait=False,
):
    started_at = time.perf_counter()
    if known_row_count is None:
        before = read_table_row_count_by_path(jab, located)
        if not before.get("ok"):
            return {
                "ok": False,
                "fast_path": True,
                "fallback_required": True,
                "reason": before.get("reason"),
                "seconds": round(time.perf_counter() - started_at, 3),
            }
        before_rows = int(before.get("row_count") or 0)
    else:
        before_rows = int(known_row_count or 0)
    if before_rows <= expected_rows:
        return _with_delete_effect(
            {
                "ok": True,
                "skipped": True,
                "fast_path": True,
                "before_rows": before_rows,
                "after_rows": before_rows,
                "steps": [],
                "seconds": round(time.perf_counter() - started_at, 3),
            },
            expected_rows=expected_rows,
        )

    best = located.get("best") or {}
    table_window = best.get("window") or {}
    steps = []
    current_rows = before_rows
    while current_rows > expected_rows:
        step_started_at = time.perf_counter()
        target_row = current_rows - 1
        focused = focus_detail_cell(jab, located, target_row, 1)
        if not focused.get("ok"):
            return _with_delete_effect(
                {
                    "ok": False,
                    "fast_path": True,
                    "fallback_required": True,
                    "before_rows": before_rows,
                    "after_rows": current_rows,
                    "steps": steps,
                    "focused": focused,
                    "reason": focused.get("reason"),
                    "seconds": round(time.perf_counter() - started_at, 3),
                },
                expected_rows=expected_rows,
            )
        guard = guard_extra_row_deletable(jab, located, target_row)
        if not guard.get("ok"):
            return _with_delete_effect(
                {
                    "ok": False,
                    "fast_path": True,
                    "fallback_required": False,
                    "before_rows": before_rows,
                    "after_rows": current_rows,
                    "steps": steps,
                    "guard": guard,
                    "reason": guard.get("reason"),
                    "seconds": round(time.perf_counter() - started_at, 3),
                },
                expected_rows=expected_rows,
            )
        sent = guarded_send_ctrl_d(table_window)
        if defer_wait:
            waited = {
                "ok": True,
                "skipped": True,
                "expected_rows": current_rows - 1,
                "actual_rows": current_rows - 1 if sent.get("ok") else current_rows,
                "reason": "行数读回交给最终 pipeline verifier",
            }
        else:
            waited = wait_table_row_count_by_path(
                jab,
                located,
                expected_rows=current_rows - 1,
                label="after_fast_extra_row_delete",
            )
        after_rows = int(waited.get("actual_rows") or 0)
        step = {
            "target_row": target_row,
            "before_rows": current_rows,
            "after_rows": after_rows,
            "seconds": round(time.perf_counter() - step_started_at, 3),
            "focused": focused,
            "sent": sent,
            "waited": waited,
            "ok": bool(sent.get("ok")) and waited.get("ok"),
            "reason": None
            if bool(sent.get("ok")) and waited.get("ok")
            else sent.get("reason") or waited.get("reason"),
        }
        steps.append(step)
        if not step["ok"]:
            return _with_delete_effect(
                {
                    "ok": False,
                    "fast_path": True,
                    "fallback_required": True,
                    "before_rows": before_rows,
                    "after_rows": after_rows,
                    "steps": steps,
                    "reason": step.get("reason"),
                    "seconds": round(time.perf_counter() - started_at, 3),
                },
                expected_rows=expected_rows,
            )
        current_rows = after_rows

    return _with_delete_effect(
        {
            "ok": True,
            "fast_path": True,
            "before_rows": before_rows,
            "after_rows": current_rows,
            "steps": steps,
            "seconds": round(time.perf_counter() - started_at, 3),
            "reason": None,
        },
        expected_rows=expected_rows,
    )


def _delete_row_step(
    target_row,
    current_rows,
    after_rows,
    step_started_at,
    focused,
    sent,
    waited,
):
    ok = bool(sent.get("ok")) and waited.get("ok") and after_rows == current_rows - 1
    return {
        "target_row": target_row,
        "before_rows": current_rows,
        "after_rows": after_rows,
        "seconds": round(time.perf_counter() - step_started_at, 3),
        "focused": focused,
        "sent": sent,
        "waited": waited,
        "ok": ok,
        "reason": None
        if ok
        else sent.get("reason")
        or waited.get("reason")
        or f"Ctrl+D 后行数未从 {current_rows} 变为 {current_rows - 1}，实际 {after_rows}",
    }
