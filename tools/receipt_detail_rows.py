# 职责：编排收款单明细手续费行增行、写入、清空账户和删多余行流程
# 不做什么：不定义字段映射，不负责 CLI/打印，不直接实现通用删行循环
# 允许依赖层：tools.receipt_detail_fields/reader/writer/row_cleanup、tools.receipt_body_table_locator
# 谁不应该 import：配置校验、Sheet 写入、收款匹配模块不应 import

import time

from tools.receipt_body_table_locator import locate_receipt_body_table_cached
from tools.receipt_detail_fields import (
    BUSINESS_TYPE_COL,
    FEE_FIELDS,
    SUBJECT_COL,
    build_fee_business,
    cells_from_steps,
    field_matches,
    normalize_text,
)
from tools.receipt_detail_reader import (
    read_located_body_table,
    read_row_cells,
    wait_body_row_count,
)
from tools.receipt_detail_row_cleanup import (
    clear_fee_account_if_filled,
    delete_extra_row_if_present,
    read_fee_prepare_row_count,
    skip_fee_account_clear_enabled,
    skip_fee_extra_row_delete_enabled,
)
from tools.receipt_detail_writer import write_detail_line_by_screen
from tools.tmp_receipt_cell_probe_run import guarded_send_ctrl_i

ADD_FEE_ROW_HOTKEY = "Ctrl+I"


class StepTimer:
    def __init__(self):
        self.items = []

    def measure(self, name, func, *args, **kwargs):
        started_at = time.perf_counter()
        result = func(*args, **kwargs)
        self.add(name, time.perf_counter() - started_at)
        return result

    def add(self, name, seconds):
        self.items.append({"name": name, "seconds": round(float(seconds), 3)})


def guarded_add_fee_row_by_ctrl_i(jab, located, scope_hwnd=None):
    started_at = time.perf_counter()
    before = read_located_body_table(jab, located, "before_fee_row_add", scope_hwnd)
    if not before.get("ok"):
        return {
            "ok": False,
            "reason": f"增行前无法读取明细表：{before.get('reason')}",
            "before": before,
            "seconds": round(time.perf_counter() - started_at, 3),
        }

    best = located.get("best") or {}
    table_window = best.get("window") or {}
    pressed = guarded_send_ctrl_i(table_window)
    after = (
        wait_body_row_count(
            jab,
            located,
            expected_rows=int(before.get("row_count") or 0) + 1,
            label="after_fee_row_add",
            scope_hwnd=scope_hwnd,
        ).get("snapshot")
        or {}
    )
    before_rows = int(before.get("row_count") or 0)
    after_rows = int(after.get("row_count") or 0)
    ok = (
        bool(pressed.get("ok"))
        and bool(after.get("ok"))
        and after_rows == before_rows + 1
    )
    return {
        "ok": ok,
        "hotkey": ADD_FEE_ROW_HOTKEY,
        "before_rows": before_rows,
        "after_rows": after_rows,
        "before": before,
        "after": after,
        "pressed": pressed,
        "seconds": round(time.perf_counter() - started_at, 3),
        "reason": None
        if ok
        else (
            pressed.get("reason")
            or after.get("reason")
            or f"行数未按预期从 {before_rows} 变为 {before_rows + 1}，实际 {after_rows}"
        ),
    }


def located_with_row_count(located, row_count):
    best = dict((located.get("best") or {}))
    if row_count:
        best["row_count"] = int(row_count)
    return {**located, "best": best}


def read_fee_row_overwrite_guard(jab, located, row_index, fee_business):
    snapshot, cells = read_row_cells(jab, row_index, located)
    if not snapshot.get("ok"):
        return {
            "ok": False,
            "changed": False,
            "reason": f"覆盖手续费行前无法读取第 {row_index + 1} 行：{snapshot.get('reason')}",
            "snapshot": snapshot,
        }
    business_type = normalize_text(cells.get(str(BUSINESS_TYPE_COL)))
    subject = normalize_text(cells.get(str(SUBJECT_COL)))
    empty = not business_type and not subject
    already_fee = business_type == fee_business["fee_business_type"] and field_matches(
        subject,
        fee_business["fee_subject"],
        "code_prefix",
    )
    ok = empty or already_fee
    return {
        "ok": ok,
        "changed": False,
        "empty": empty,
        "already_fee": already_fee,
        "cells": cells,
        "snapshot": snapshot,
        "reason": None
        if ok
        else (
            f"第 {row_index + 1} 行已有非手续费业务，拒绝覆盖："
            f"业务类型={business_type!r}，科目={subject!r}，"
            f"期望空行或手续费/{fee_business['fee_subject']}"
        ),
    }


def run_fee_only(jab, located, fee_amount, scope_hwnd=None):
    timings = StepTimer()
    fee_business = build_fee_business(fee_amount)
    before = timings.measure(
        "fee.read-before-prepare",
        read_fee_prepare_row_count,
        jab,
        located,
        scope_hwnd,
    )
    before_rows = int(before.get("row_count") or 0)
    if not before.get("ok"):
        add_row = {
            "ok": False,
            "reason": f"手续费准备前无法读取明细表：{before.get('reason')}",
            "before": before,
        }
    elif before_rows > 2:
        cleanup_extra = timings.measure(
            "fee.cleanup-to-second-row",
            delete_extra_row_if_present,
            jab,
            located,
            expected_rows=2,
            scope_hwnd=scope_hwnd,
        )
        if not cleanup_extra.get("ok"):
            cleanup_extra["timings"] = timings.items
            return (
                cleanup_extra,
                [],
                {
                    "ok": False,
                    "skipped": True,
                    "reason": "清理到第 2 行失败，未清空手续费账户",
                },
                cleanup_extra,
            )
        located = timings.measure(
            "fee.locate-after-cleanup",
            locate_receipt_body_table_cached,
            jab,
            cached=located,
            max_rows=5,
            scope_hwnd=scope_hwnd
            or ((located.get("best") or {}).get("window") or {}).get("hwnd"),
        )
        before_rows = 2
        overwrite_guard = timings.measure(
            "fee.guard-row2-after-cleanup",
            read_fee_row_overwrite_guard,
            jab,
            located,
            1,
            fee_business,
        )
        if not overwrite_guard.get("ok"):
            overwrite_guard["timings"] = timings.items
            return (
                overwrite_guard,
                [],
                {
                    "ok": False,
                    "skipped": True,
                    "reason": "手续费第 2 行覆盖守卫失败，未清空账户",
                },
                {
                    "ok": False,
                    "skipped": True,
                    "reason": "手续费第 2 行覆盖守卫失败，未删除多余行",
                },
            )
        add_row = {
            "ok": True,
            "skipped": True,
            "reason": "当前超过 2 行，已删到 2 行并覆盖第 2 行为手续费行",
            "hotkey": ADD_FEE_ROW_HOTKEY,
            "before_rows": before.get("row_count"),
            "after_rows": before_rows,
            "before": before,
            "overwrite_guard": overwrite_guard,
        }
    elif before_rows == 1:
        add_row = guarded_add_fee_row_by_ctrl_i(jab, located, scope_hwnd=scope_hwnd)
        timings.add("fee.add-row", add_row.get("seconds") or 0)
    elif before_rows == 2:
        overwrite_guard = timings.measure(
            "fee.guard-row2-before-overwrite",
            read_fee_row_overwrite_guard,
            jab,
            located,
            1,
            fee_business,
        )
        if not overwrite_guard.get("ok"):
            add_row = {
                "ok": False,
                "skipped": True,
                "reason": overwrite_guard.get("reason"),
                "hotkey": ADD_FEE_ROW_HOTKEY,
                "before_rows": before_rows,
                "after_rows": before_rows,
                "before": before,
                "overwrite_guard": overwrite_guard,
            }
        else:
            add_row = {
                "ok": True,
                "skipped": True,
                "reason": "当前已有 2 行，第 2 行为空或已是手续费行，直接覆盖为手续费行",
                "hotkey": ADD_FEE_ROW_HOTKEY,
                "before_rows": before_rows,
                "after_rows": before_rows,
                "before": before,
                "after": before,
                "overwrite_guard": overwrite_guard,
            }
    else:
        add_row = {
            "ok": False,
            "reason": f"手续费行固定第 2 行，但清理后当前仍有 {before_rows} 行",
            "before_rows": before_rows,
            "after_rows": before_rows,
            "before": before,
        }

    if not add_row.get("ok"):
        add_row["timings"] = timings.items
        return (
            add_row,
            [],
            {"ok": False, "skipped": True, "reason": "增行失败，未清空手续费账户"},
            {"ok": False, "skipped": True, "reason": "增行失败，未删除多余行"},
        )

    refreshed = located_with_row_count(
        located, int(add_row.get("after_rows") or before_rows)
    )

    target_row = 1
    steps = timings.measure(
        "fee.write-line",
        write_detail_line_by_screen,
        jab,
        fee_business,
        refreshed,
        fields=FEE_FIELDS,
        row_index=target_row,
    )
    if skip_fee_account_clear_enabled():
        clear_account = {
            "ok": True,
            "skipped": True,
            "reason": "RECEIPT_SKIP_FEE_ACCOUNT_CLEAR=1，试验不清手续费账户",
        }
    else:
        clear_account = timings.measure(
            "fee.clear-account-if-filled",
            clear_fee_account_if_filled,
            jab,
            refreshed,
            target_row,
            known_cells=cells_from_steps(steps),
        )
    if skip_fee_extra_row_delete_enabled():
        delete_extra = {
            "ok": True,
            "skipped": True,
            "reason": "RECEIPT_SKIP_FEE_EXTRA_ROW_DELETE=1，试验保留空白行直接保存",
            "timings": timings.items,
        }
    else:
        delete_extra = timings.measure(
            "fee.delete-extra-after-write",
            delete_extra_row_if_present,
            jab,
            refreshed,
            expected_rows=2,
            scope_hwnd=scope_hwnd,
        )
    delete_extra["timings"] = timings.items
    return add_row, steps, clear_account, delete_extra
