# 职责：按字段映射写入收款单明细行，并把字段校验交给调用方的后台 verifier
# 不做什么：不增删明细行，不读取 Excel，不处理 CLI/打印
# 允许依赖层：core.receipt_detail_fields/reader/screen_writer、core.receipt_body_table_locator
# 谁不应该 import：配置校验、Sheet 写入、收款匹配模块不应 import

import time
from decimal import Decimal, InvalidOperation

from core.receipt_detail_fields import (
    DETAIL_FIELDS,
    field_expected_value,
    field_matches,
    make_detail_step,
)
from core.receipt_detail_reader import read_row_cells
from core.receipt_detail_screen_writer import (
    KEYBOARD_INPUT_COMMIT_KEY,
    focus_detail_cell,
    keyboard_write_selected_cell,
    move_selected_cell_by_arrows,
)
from core.receipt_keyboard_utils import STOP_HOTKEY, is_stop_hotkey_pressed

DETAIL_DIAGNOSTIC_COLS = (4, 5, 6, 7, 8, 11)


def write_field_once(
    jab,
    located,
    table_window,
    row_index,
    field,
    business,
    attempt_no,
    current_col=None,
    recover_after_failure=None,
    readback_context=None,
):
    value = str(business[field["value_key"]])
    attempt_start = time.perf_counter()
    stage_timing = {}
    target_col = int(field["col"])
    stage_started = time.perf_counter()
    focus, navigation = focus_entry_for_field(
        jab,
        located,
        table_window,
        row_index,
        field,
        current_col=current_col,
    )
    stage_timing["focus_entry"] = round(time.perf_counter() - stage_started, 4)
    if focus.get("ok") and navigation.get("ok"):
        stage_started = time.perf_counter()
        activation = activate_field_before_write(jab, located, row_index, field)
        stage_timing["activation"] = round(time.perf_counter() - stage_started, 4)
        if activation.get("ok"):
            stage_started = time.perf_counter()
            pre_write_stabilize = stabilize_field_before_write(
                jab,
                located,
                table_window,
                row_index,
                field,
                previous_focus=focus,
                previous_navigation=navigation,
            )
            stage_timing["pre_write_stabilize"] = round(
                time.perf_counter() - stage_started,
                4,
            )
        else:
            pre_write_stabilize = {
                "ok": True,
                "skipped": True,
                "reason": "字段激活失败，未执行写前稳定",
            }
            stage_timing["pre_write_stabilize"] = 0.0
        if activation.get("ok") and pre_write_stabilize.get("ok"):
            stage_started = time.perf_counter()
            neighbor_guard_before = capture_sensitive_neighbors(
                jab,
                located,
                row_index,
                field,
                readback_context=readback_context,
            )
            stage_timing["neighbor_guard_before"] = round(
                time.perf_counter() - stage_started,
                4,
            )
        else:
            neighbor_guard_before = {
                "ok": True,
                "skipped": True,
                "reason": "字段激活或写前稳定失败，未执行敏感邻列哨兵",
            }
            stage_timing["neighbor_guard_before"] = 0.0
        if (
            activation.get("ok")
            and pre_write_stabilize.get("ok")
            and neighbor_guard_before.get("ok")
        ):
            stage_started = time.perf_counter()
            screen = write_field_value_to_focused_cell(
                table_window,
                value,
                field,
                recover_after_failure=recover_after_failure,
            )
            stage_timing["screen_write"] = round(
                time.perf_counter() - stage_started,
                4,
            )
        else:
            screen = {
                "ok": False,
                "reason": activation.get("reason")
                or pre_write_stabilize.get("reason")
                or neighbor_guard_before.get("reason"),
                "activation": activation,
                "pre_write_stabilize": pre_write_stabilize,
                "neighbor_guard_before": neighbor_guard_before,
            }
            stage_timing["screen_write"] = 0.0
    else:
        activation = {"ok": False, "skipped": True, "reason": "定位或导航失败，未激活"}
        pre_write_stabilize = {
            "ok": True,
            "skipped": True,
            "reason": "定位或导航失败，未执行写前稳定",
        }
        neighbor_guard_before = {
            "ok": True,
            "skipped": True,
            "reason": "定位或导航失败，未执行敏感邻列哨兵",
        }
        screen = {
            "ok": False,
            "reason": focus.get("reason") or navigation.get("reason"),
            "focus": focus,
            "navigation": navigation,
        }
        stage_timing.update(
            {
                "activation": 0.0,
                "pre_write_stabilize": 0.0,
                "neighbor_guard_before": 0.0,
                "screen_write": 0.0,
            }
        )
    attempt_seconds = round(time.perf_counter() - attempt_start, 3)
    known_seconds = sum(float(value or 0) for value in stage_timing.values())
    stage_timing["unaccounted"] = round(max(attempt_seconds - known_seconds, 0.0), 4)
    return {
        "attempt": attempt_no,
        "seconds": attempt_seconds,
        "stage_timing": stage_timing,
        "mode": "keyboard",
        "input_ok": bool(screen.get("ok")),
        "input_reason": screen.get("reason"),
        "target": {"row": row_index, "col": target_col},
        "table_bounds": None,
        "cell_width": None,
        "cell_height": None,
        "focus": focus,
        "navigation": navigation,
        "activation": activation,
        "pre_write_stabilize": pre_write_stabilize,
        "neighbor_guard_before": neighbor_guard_before,
        "commit_ok": bool(screen.get("ok")),
        "commit_key": field.get("commit_key") or KEYBOARD_INPUT_COMMIT_KEY,
        "accept_key": field.get("accept_key"),
        "typing_interval": field.get("typing_interval", 0.0),
        "edit_mode": field.get("edit_mode", "editor"),
        "input_mode": field.get("input_mode", "paste"),
        "pre_commit_wait": field.get("pre_commit_wait", 0.025),
        "commit_col": current_col_after_commit(target_col, field.get("commit_key")),
        "commit_target": {
            "ok": True,
            "skipped": True,
            "reason": "字段写入后不做同步读回；由后台 verifier 或调用方统一闭包",
        },
        "commit_reason": screen.get("reason"),
        "screen_timing": screen.get("screen_timing"),
        "screen_commit": screen.get("commit"),
        "screen_accept": screen.get("accept"),
        "ok": bool(screen.get("ok")),
        "modal_recovery": screen.get("modal_recovery"),
        "retry_modal_recovery": screen.get("retry_modal_recovery"),
    }


def focus_entry_for_field(
    jab,
    located,
    table_window,
    row_index,
    field,
    current_col=None,
):
    target_col = int(field["col"])
    via_col = field.get("focus_via_col")
    if via_col is None:
        focus = focus_detail_cell(jab, located, row_index, target_col)
        return focus, {
            "ok": True,
            "skipped": True,
            "reason": "每个字段按明细表 path 重新定位到目标单元格",
            "from_col": current_col,
            "to_col": target_col,
        }

    entry_col = int(via_col)
    focus = focus_detail_cell(jab, located, row_index, entry_col)
    if not focus.get("ok"):
        return focus, {
            "ok": False,
            "from_col": entry_col,
            "to_col": target_col,
            "reason": focus.get("reason") or "字段入口列定位失败",
        }
    navigation = move_selected_cell_by_arrows(table_window, entry_col, target_col)
    navigation["via_col"] = entry_col
    if navigation.get("ok") and not navigation.get("reason"):
        navigation["reason"] = (
            f"{field['name']} 先按 path 定位入口列 {entry_col}，再方向键进入目标列 {target_col}"
        )
    return focus, navigation


def activate_field_before_write(jab, located, row_index, field):
    return {
        "ok": True,
        "skipped": True,
        "method": "keyboard-only",
        "target": {"row": row_index, "col": int(field["col"])},
        "reason": "正式明细写入不使用鼠标 bounds；已按 path/键盘定位目标单元格",
    }


def stabilize_field_before_write(
    jab,
    located,
    table_window,
    row_index,
    field,
    previous_focus=None,
    previous_navigation=None,
):
    if not field.get("pre_write_stabilize"):
        return {
            "ok": True,
            "skipped": True,
            "reason": "字段未启用写前稳定",
        }

    started_at = time.perf_counter()
    wait_seconds = max(float(field.get("pre_write_stabilize_wait") or 0.0), 0.0)
    if wait_seconds:
        time.sleep(wait_seconds)
    if (
        not field.get("pre_write_stabilize_refocus")
        and (previous_focus or {}).get("ok")
        and (previous_navigation or {}).get("ok")
    ):
        return {
            "ok": True,
            "skipped": False,
            "method": "reuse-current-focus",
            "seconds": round(time.perf_counter() - started_at, 3),
            "wait_seconds": wait_seconds,
            "row": int(row_index),
            "col": int(field["col"]),
            "focus": previous_focus,
            "navigation": previous_navigation,
            "reason": "已复用刚完成的目标单元格定位，未重复请求 JAB 选中",
        }
    focus, navigation = focus_entry_for_field(
        jab,
        located,
        table_window,
        row_index,
        field,
    )
    ok = bool(focus.get("ok")) and bool(navigation.get("ok"))
    return {
        "ok": ok,
        "skipped": False,
        "method": "refocus-only",
        "seconds": round(time.perf_counter() - started_at, 3),
        "wait_seconds": wait_seconds,
        "row": int(row_index),
        "col": int(field["col"]),
        "focus": focus,
        "navigation": navigation,
        "reason": None
        if ok
        else (
            focus.get("reason")
            or navigation.get("reason")
            or "写前稳定后重新定位目标单元格失败"
        ),
    }


def write_field_value_to_focused_cell(
    table_window,
    value,
    field,
    recover_after_failure=None,
):
    commit_key = field.get("commit_key") or KEYBOARD_INPUT_COMMIT_KEY
    return keyboard_write_selected_cell(
        table_window,
        value,
        commit_key=commit_key,
        clear_only=field.get("kind") == "blank",
        accept_key=field.get("accept_key"),
        typing_interval=field.get("typing_interval", 0.0),
        edit_mode=field.get("edit_mode", "editor"),
        input_mode=field.get("input_mode", "paste"),
        pre_commit_wait=field.get("pre_commit_wait", 0.025),
        recover_after_failure=recover_after_failure,
    )


def sensitive_neighbor_cols(field):
    cols = []
    for col in field.get("sensitive_neighbor_cols") or []:
        try:
            cols.append(int(col))
        except (TypeError, ValueError):
            continue
    return cols


def open_detail_readback_context(jab, located):
    best = (located or {}).get("best") or {}
    path = best.get("path")
    cached_window = best.get("window") or {}
    if not path or not hasattr(jab, "find_context_by_path_once"):
        return {
            "ok": False,
            "reason": "缺少明细表 path 或 JAB context path 能力",
        }
    try:
        context, vm_id, owned, window = jab.find_context_by_path_once(
            path,
            class_name=cached_window.get("class_name") or cached_window.get("class"),
            scope_hwnd=cached_window.get("hwnd"),
            role="table",
            require_showing=False,
            require_valid_bounds=False,
        )
        if not context:
            return {"ok": False, "path": path, "reason": "明细表 context 未命中"}
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            release_detail_readback_context(jab, {"vm_id": vm_id, "owned": owned})
            return {"ok": False, "path": path, "reason": "明细表 table_info 不可读"}
        return {
            "ok": True,
            "path": path,
            "context": context,
            "vm_id": vm_id,
            "owned": owned,
            "window": window or cached_window,
            "row_count": int(table_info.rowCount),
            "col_count": int(table_info.columnCount),
        }
    except Exception as exc:
        return {
            "ok": False,
            "path": path,
            "reason": f"{type(exc).__name__}: {exc}",
        }


def release_detail_readback_context(jab, context):
    if not context or not context.get("vm_id") or not context.get("owned"):
        return
    try:
        jab.release_contexts(context["vm_id"], context["owned"])
    except Exception:
        pass


def read_row_cells_by_context(jab, readback_context, row_index, cols, step):
    if not readback_context or not readback_context.get("ok"):
        return {
            "step": step,
            "ok": False,
            "fast_context": False,
            "reason": (readback_context or {}).get("reason")
            or "明细表快读 context 不可用",
        }, {}
    row_count = int(readback_context.get("row_count") or 0)
    col_count = int(readback_context.get("col_count") or 0)
    if row_index < 0 or row_index >= row_count:
        return {
            "step": step,
            "ok": False,
            "fast_context": True,
            "path": readback_context.get("path"),
            "row_count": row_count,
            "col_count": col_count,
            "reason": f"明细表快读行号越界：row={row_index}, row_count={row_count}",
        }, {}
    read_cols = sorted({int(col) for col in cols if 0 <= int(col) < col_count})
    try:
        cells = {}
        for col in read_cols:
            text, _selected = jab.get_table_cell_text_and_selection(
                readback_context["vm_id"],
                readback_context["context"],
                int(row_index),
                int(col),
            )
            cells[str(col)] = text
        return {
            "step": step,
            "ok": True,
            "fast_context": True,
            "fast_path": True,
            "semantic_fallback_used": False,
            "method": "table-context-cells",
            "path": readback_context.get("path"),
            "row_count": row_count,
            "col_count": col_count,
            "read_columns": read_cols,
            "reason": None,
        }, cells
    except Exception as exc:
        return {
            "step": step,
            "ok": False,
            "fast_context": True,
            "path": readback_context.get("path"),
            "row_count": row_count,
            "col_count": col_count,
            "read_columns": read_cols,
            "reason": f"{type(exc).__name__}: {exc}",
        }, {}


def read_row_cells_for_verify(
    jab,
    located,
    row_index,
    cols,
    step,
    readback_context=None,
    fallback=True,
):
    fast_snapshot, fast_cells = read_row_cells_by_context(
        jab,
        readback_context,
        row_index,
        cols,
        step,
    )
    if fast_snapshot.get("ok"):
        return fast_snapshot, fast_cells
    if not fallback:
        return fast_snapshot, fast_cells
    snapshot, cells = read_row_cells(jab, row_index, located)
    if isinstance(snapshot, dict):
        snapshot = {
            **snapshot,
            "fast_context_fallback_used": True,
            "fast_context_failure": fast_snapshot,
        }
    return snapshot, cells


def capture_sensitive_neighbors(jab, located, row_index, field, readback_context=None):
    cols = sensitive_neighbor_cols(field)
    if not cols:
        return {
            "ok": True,
            "skipped": True,
            "reason": "字段未配置敏感邻列哨兵",
        }
    started_at = time.perf_counter()
    read_cols = list(cols) + [int(field["col"])]
    snapshot, cells = read_row_cells_for_verify(
        jab,
        located,
        row_index,
        read_cols,
        "sensitive_neighbor_before",
        readback_context=readback_context,
    )
    if not snapshot.get("ok"):
        return {
            "ok": False,
            "skipped": False,
            "seconds": round(time.perf_counter() - started_at, 3),
            "cols": cols,
            "reason": snapshot.get("reason") or "敏感邻列写前快照读取失败",
            "snapshot": {
                "ok": snapshot.get("ok"),
                "fast_path": snapshot.get("fast_path"),
                "fast_context": snapshot.get("fast_context"),
                "method": snapshot.get("method"),
                "semantic_fallback_used": snapshot.get("semantic_fallback_used"),
                "path": snapshot.get("path"),
                "row_count": snapshot.get("row_count"),
                "col_count": snapshot.get("col_count"),
                "reason": snapshot.get("reason"),
            },
        }
    values = {str(col): cells.get(str(col)) for col in cols}
    values[str(field["col"])] = cells.get(str(field["col"]))
    return {
        "ok": True,
        "skipped": False,
        "method": "sensitive-neighbor-snapshot",
        "seconds": round(time.perf_counter() - started_at, 3),
        "row": int(row_index),
        "target_col": int(field["col"]),
        "cols": cols,
        "values": values,
        "snapshot": {
            "ok": snapshot.get("ok"),
            "fast_path": snapshot.get("fast_path"),
            "fast_context": snapshot.get("fast_context"),
            "method": snapshot.get("method"),
            "semantic_fallback_used": snapshot.get("semantic_fallback_used"),
            "path": snapshot.get("path"),
            "row_count": snapshot.get("row_count"),
            "col_count": snapshot.get("col_count"),
            "reason": snapshot.get("reason"),
        },
    }


def validate_sensitive_neighbors_after_write(field, before_guard, cells):
    cols = sensitive_neighbor_cols(field)
    if not cols:
        return {
            "ok": True,
            "skipped": True,
            "reason": "字段未配置敏感邻列哨兵",
        }
    if not before_guard or not before_guard.get("ok") or before_guard.get("skipped"):
        return {
            "ok": False,
            "skipped": False,
            "reason": "敏感邻列缺少写前快照，不能确认邻列未被污染",
            "before": before_guard,
        }
    before_values = before_guard.get("values") or {}
    changes = []
    for col in cols:
        key = str(col)
        before = before_values.get(key)
        after = cells.get(key)
        if not cell_text_equivalent(before, after):
            changes.append(
                {"col": col, "before": before_values.get(key), "after": cells.get(key)}
            )
    return {
        "ok": not changes,
        "skipped": False,
        "method": "sensitive-neighbor-compare",
        "cols": cols,
        "changes": changes,
        "reason": None if not changes else f"敏感邻列疑似被写入/改动：{changes}",
    }


def detail_diagnostic_cells(cells, cols=DETAIL_DIAGNOSTIC_COLS):
    return {str(col): (cells or {}).get(str(col)) for col in cols}


def normalize_cell_text(value):
    return str(value or "").strip()


def cell_text_equivalent(left, right):
    left_text = normalize_cell_text(left).replace(",", "")
    right_text = normalize_cell_text(right).replace(",", "")
    if left_text == right_text:
        return True
    if left_text and right_text:
        try:
            return Decimal(left_text) == Decimal(right_text)
        except (InvalidOperation, ValueError):
            return False
    return False


def current_col_after_commit(target_col, commit_key):
    key = commit_key or KEYBOARD_INPUT_COMMIT_KEY
    if str(key).lower() in {"enter", "right"}:
        return int(target_col) + 1
    if str(key).lower() == "left":
        return max(int(target_col) - 1, 0)
    return int(target_col)


def verify_detail_field_now(
    jab,
    located,
    row_index,
    field,
    business,
    readback_context=None,
):
    started_at = time.perf_counter()
    verify_cols = [int(field["col"])] + sensitive_neighbor_cols(field)
    snapshot, cells = read_row_cells_for_verify(
        jab,
        located,
        row_index,
        verify_cols,
        "field_readback",
        readback_context=readback_context,
    )
    actual = cells.get(str(field["col"]))
    expected = field_expected_value(field, business)
    ok = bool(snapshot.get("ok")) and field_matches(
        actual,
        str(business[field["value_key"]]),
        field.get("kind"),
    )
    return {
        "ok": ok,
        "seconds": round(time.perf_counter() - started_at, 3),
        "row": int(row_index),
        "col": field["col"],
        "name": field["name"],
        "expected": expected,
        "actual": actual,
        "snapshot": {
            "ok": snapshot.get("ok"),
            "fast_path": snapshot.get("fast_path"),
            "fast_context": snapshot.get("fast_context"),
            "method": snapshot.get("method"),
            "semantic_fallback_used": snapshot.get("semantic_fallback_used"),
            "path": snapshot.get("path"),
            "row_count": snapshot.get("row_count"),
            "col_count": snapshot.get("col_count"),
            "read_columns": snapshot.get("read_columns"),
            "fast_context_fallback_used": snapshot.get("fast_context_fallback_used"),
            "fast_context_failure": snapshot.get("fast_context_failure"),
            "reason": snapshot.get("reason"),
        },
        "reason": None
        if ok
        else (
            snapshot.get("reason")
            or f"即时校验未匹配：字段={field['name']}，期望={expected!r}，实际={actual!r}"
        ),
        "cells": cells,
    }


def ensure_field_immediate_verified(
    jab,
    located,
    table_window,
    row_index,
    field,
    business,
    first_attempt,
    recover_after_failure=None,
    readback_context=None,
):
    if not field.get("immediate_verify") or not first_attempt.get("ok"):
        return first_attempt, None
    max_attempts = max(int(field.get("immediate_verify_attempts") or 1), 1)
    wait_seconds = max(float(field.get("immediate_verify_wait") or 0.0), 0.0)
    verifications = []
    rewrites = []
    for attempt_index in range(max_attempts):
        if wait_seconds:
            time.sleep(wait_seconds)
        verification = verify_detail_field_now(
            jab,
            located,
            row_index,
            field,
            business,
            readback_context=readback_context,
        )
        verification["attempt_index"] = attempt_index + 1
        verification_cells = verification.pop("cells", {}) or {}
        neighbor_guard_after = validate_sensitive_neighbors_after_write(
            field,
            first_attempt.get("neighbor_guard_before"),
            verification_cells,
        )
        verification["neighbor_guard_after"] = neighbor_guard_after
        if verification.get("ok") and not neighbor_guard_after.get("ok"):
            verification["ok"] = False
            verification["reason"] = neighbor_guard_after.get("reason")
        if not verification.get("ok"):
            verification["diagnostic_cells"] = detail_diagnostic_cells(
                verification_cells
            )
        verifications.append(verification)
        if verification.get("ok"):
            return first_attempt, {
                "ok": True,
                "verified": True,
                "field": field["name"],
                "verifications": verifications,
                "rewrites": rewrites,
            }
        if attempt_index >= max_attempts - 1:
            break
        focus, navigation = focus_entry_for_field(
            jab,
            located,
            table_window,
            row_index,
            field,
        )
        activation = (
            activate_field_before_write(jab, located, row_index, field)
            if focus.get("ok") and navigation.get("ok")
            else {"ok": False, "skipped": True, "reason": "定位或导航失败，未激活"}
        )
        if not focus.get("ok") or not navigation.get("ok") or not activation.get("ok"):
            rewrites.append(
                {
                    "ok": False,
                    "stage": "focus_or_navigation",
                    "focus": focus,
                    "navigation": navigation,
                    "activation": activation,
                    "reason": (
                        focus.get("reason")
                        or navigation.get("reason")
                        or activation.get("reason")
                    ),
                }
            )
            break
        screen = write_field_value_to_focused_cell(
            table_window,
            str(business[field["value_key"]]),
            field,
            recover_after_failure=recover_after_failure,
        )
        rewrites.append(
            {
                "ok": bool(screen.get("ok")),
                "stage": "rewrite",
                "focus": focus,
                "navigation": navigation,
                "activation": activation,
                "screen_timing": screen.get("screen_timing"),
                "screen_commit": screen.get("commit"),
                "reason": screen.get("reason"),
            }
        )
        if not screen.get("ok"):
            break
    failed = dict(first_attempt)
    failed["ok"] = False
    failed["input_ok"] = False
    failed["commit_ok"] = False
    failed["input_reason"] = (
        (verifications[-1] if verifications else {}).get("reason")
        or (rewrites[-1] if rewrites else {}).get("reason")
        or f"{field['name']} 即时校验失败"
    )
    failed["commit_reason"] = failed["input_reason"]
    return failed, {
        "ok": False,
        "verified": False,
        "field": field["name"],
        "verifications": verifications,
        "rewrites": rewrites,
        "reason": failed["input_reason"],
    }


def write_detail_line_by_screen(
    jab,
    business,
    located,
    fields=None,
    row_index=0,
    after_field=None,
    recover_after_failure=None,
):
    line_started_at = time.perf_counter()
    line_timing = {}
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
    stage_started = time.perf_counter()
    readback_context = open_detail_readback_context(jab, located)
    line_timing["open_readback_context"] = round(
        time.perf_counter() - stage_started,
        4,
    )
    try:
        for index, field in enumerate(fields):
            field_loop_started = time.perf_counter()
            stage_started = time.perf_counter()
            if is_stop_hotkey_pressed():
                stop_hotkey_seconds = round(time.perf_counter() - stage_started, 4)
                steps.append(
                    {
                        "ok": False,
                        "changed": False,
                        "partial_success": bool(steps),
                        "name": field["name"],
                        "value": business[field["value_key"]],
                        "reason": f"检测到紧急停止键 {STOP_HOTKEY}",
                        "line_timing": {"stop_hotkey": stop_hotkey_seconds},
                    }
                )
                break
            stop_hotkey_seconds = round(time.perf_counter() - stage_started, 4)

            stage_started = time.perf_counter()
            step = make_detail_step(field, business, row_index, row_count, col_count)
            make_step_seconds = round(time.perf_counter() - stage_started, 4)
            stage_started = time.perf_counter()
            attempt = write_field_once(
                jab,
                located,
                table_window,
                row_index,
                field,
                business,
                attempt_no=1,
                current_col=current_col,
                recover_after_failure=recover_after_failure,
                readback_context=readback_context,
            )
            write_once_seconds = round(time.perf_counter() - stage_started, 4)
            if attempt.get("ok"):
                commit_col = attempt.get("commit_col")
                if commit_col is not None:
                    current_col = int(commit_col)
            stage_started = time.perf_counter()
            attempt, immediate_verify = ensure_field_immediate_verified(
                jab,
                located,
                table_window,
                row_index,
                field,
                business,
                attempt,
                recover_after_failure=recover_after_failure,
                readback_context=readback_context,
            )
            immediate_verify_seconds = round(time.perf_counter() - stage_started, 4)
            if immediate_verify is not None:
                step["immediate_verify"] = immediate_verify
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
                stage_started = time.perf_counter()
                step["async_verify_task"] = after_field(
                    row_index, field, business, step
                )
                after_field_seconds = round(time.perf_counter() - stage_started, 4)
            else:
                after_field_seconds = 0.0
            if not attempt.get("ok"):
                step["reason"] = attempt.get("input_reason") or attempt.get(
                    "commit_reason"
                )
            field_loop_seconds = round(time.perf_counter() - field_loop_started, 4)
            known_field_seconds = (
                stop_hotkey_seconds
                + make_step_seconds
                + write_once_seconds
                + immediate_verify_seconds
                + after_field_seconds
            )
            step["line_timing"] = {
                "stop_hotkey": stop_hotkey_seconds,
                "make_step": make_step_seconds,
                "write_once": write_once_seconds,
                "immediate_verify": immediate_verify_seconds,
                "after_field": after_field_seconds,
                "field_loop": field_loop_seconds,
                "unaccounted": round(
                    max(field_loop_seconds - known_field_seconds, 0.0),
                    4,
                ),
            }
            steps.append(step)
            if not attempt.get("ok"):
                break
    finally:
        stage_started = time.perf_counter()
        release_detail_readback_context(jab, readback_context)
        line_timing["release_readback_context"] = round(
            time.perf_counter() - stage_started,
            4,
        )
    line_timing["total_before_finalize"] = round(
        time.perf_counter() - line_started_at,
        4,
    )
    line_timing["field_loops"] = round(
        sum((step.get("line_timing") or {}).get("field_loop") or 0 for step in steps),
        4,
    )
    line_timing["unaccounted_before_finalize"] = round(
        max(
            line_timing["total_before_finalize"]
            - line_timing["open_readback_context"]
            - line_timing["release_readback_context"]
            - line_timing["field_loops"],
            0.0,
        ),
        4,
    )
    for step in steps:
        immediate_verify = step.get("immediate_verify")
        if immediate_verify is not None:
            step_ok = bool(step.get("input_ok")) and bool(immediate_verify.get("ok"))
            if not step_ok and not step.get("reason"):
                step["reason"] = immediate_verify.get("reason")
        else:
            step_ok = bool(step.get("input_ok"))
        step["ok"] = step_ok
        step["blocked"] = not step["ok"]
        step["reason"] = None if step["ok"] else step.get("reason")
        step["actual"] = None
        step["deferred_readback"] = {
            "ok": True,
            "reason": (
                "后台 pipeline verifier 批量读回校验"
                if after_field
                else "调用方后续整表读回/行数闭包校验"
            ),
        }
        step["line_timing_summary"] = line_timing
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
