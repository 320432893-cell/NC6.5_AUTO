# 职责：编排收款单明细手续费行增行、写入、清空账户和删多余行流程
# 不做什么：不定义字段映射，不负责 CLI/打印，不直接实现通用删行循环
# 允许依赖层：core.receipt_detail_fields/reader/writer/row_cleanup、core.receipt_body_table_locator
# 谁不应该 import：配置校验、Sheet 写入、收款匹配模块不应 import

import time

from core.receipt_body_table_locator import locate_receipt_body_table_cached
from core.receipt_detail_fields import (
    ACCOUNT_COL,
    FEE_FIELDS,
    build_fee_business,
    cells_from_steps,
    normalize_text,
)
from core.receipt_detail_reader import (
    read_located_body_table,
    wait_body_row_count,
)
from core.receipt_detail_row_cleanup import (
    delete_extra_row_if_present,
    read_fee_prepare_row_count,
    skip_fee_extra_row_delete_enabled,
)
from core.receipt_detail_writer import write_detail_line_by_screen
from core.receipt_keyboard_utils import guarded_send_ctrl_i

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


def run_fee_only(
    jab,
    located,
    fee_amount,
    scope_hwnd=None,
    after_field=None,
    known_row_count=None,
    defer_delete_wait=False,
    recover_after_failure=None,
):
    timings = StepTimer()
    fee_business = build_fee_business(fee_amount)
    if known_row_count is None:
        before = timings.measure(
            "fee.read-before-prepare",
            read_fee_prepare_row_count,
            jab,
            located,
            scope_hwnd,
        )
    else:
        before = {
            "ok": True,
            "fast_path": True,
            "row_count": int(known_row_count),
            "source": "known_row_count",
        }
        timings.add("fee.read-before-prepare", 0.0)
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
        add_row = {
            "ok": True,
            "skipped": True,
            "reason": "当前超过 2 行，已删到 2 行；第 2 行无条件覆盖为手续费行",
            "hotkey": ADD_FEE_ROW_HOTKEY,
            "before_rows": before.get("row_count"),
            "after_rows": before_rows,
            "before": before,
        }
    elif before_rows == 1:
        add_row = guarded_add_fee_row_by_ctrl_i(jab, located, scope_hwnd=scope_hwnd)
        timings.add("fee.add-row", add_row.get("seconds") or 0)
    elif before_rows == 2:
        add_row = {
            "ok": True,
            "skipped": True,
            "reason": "当前已有 2 行，第 2 行无条件覆盖为手续费行",
            "hotkey": ADD_FEE_ROW_HOTKEY,
            "before_rows": before_rows,
            "after_rows": before_rows,
            "before": before,
            "after": before,
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
        after_field=after_field,
        recover_after_failure=recover_after_failure,
    )
    repair = repair_fee_business_type_if_needed(
        timings,
        jab,
        fee_business,
        refreshed,
        steps,
        after_field=after_field,
        recover_after_failure=recover_after_failure,
    )
    if repair.get("attempted"):
        steps.append(repair)
        if repair.get("ok"):
            steps = replace_fee_business_type_step_with_repair(steps, repair)
    fee_cells = cells_from_steps(steps)
    clear_account = {
        "ok": normalize_text(fee_cells.get(str(ACCOUNT_COL))) == "",
        "skipped": True,
        "source": "fee.write-line",
        "before": fee_cells.get(str(ACCOUNT_COL)),
        "after": fee_cells.get(str(ACCOUNT_COL)),
        "reason": None
        if normalize_text(fee_cells.get(str(ACCOUNT_COL))) == ""
        else "手续费账户列按顺序清空后读回仍非空",
    }
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
            known_row_count=int(add_row.get("after_rows") or before_rows) + 1,
            defer_wait=defer_delete_wait,
        )
    delete_extra["timings"] = timings.items
    return add_row, steps, clear_account, delete_extra


def repair_fee_business_type_if_needed(
    timings,
    jab,
    fee_business,
    located,
    steps,
    after_field=None,
    recover_after_failure=None,
):
    failed = first_failed_fee_business_type_step(steps)
    if not failed:
        return {
            "ok": True,
            "attempted": False,
            "skipped": True,
            "reason": "手续费业务类型无需修复",
        }
    retry_steps = timings.measure(
        "fee.rewrite-business-type",
        write_detail_line_by_screen,
        jab,
        fee_business,
        located,
        fields=[FEE_FIELDS[0]],
        row_index=1,
        after_field=after_field,
        recover_after_failure=recover_after_failure,
    )
    ok = bool(retry_steps) and all(step.get("ok") for step in retry_steps)
    return {
        "ok": ok,
        "attempted": True,
        "name": "手续费业务类型修复",
        "field": "收款业务类型",
        "target": {"row": 1, "col": 1},
        "retry_steps": retry_steps,
        "reason": None
        if ok
        else summarize_fee_business_type_repair_failure(retry_steps),
    }


def first_failed_fee_business_type_step(steps):
    for step in steps or []:
        if step.get("ok"):
            continue
        name = str(step.get("name") or "").strip()
        target_col = (step.get("target") or {}).get("col") or 1
        try:
            target_col = int(target_col)
        except (TypeError, ValueError):
            target_col = 1
        if name == "收款业务类型" and target_col == 1:
            return step
    return None


def replace_fee_business_type_step_with_repair(steps, repair):
    retry_steps = repair.get("retry_steps") or []
    if not retry_steps:
        return steps
    replacement = dict(retry_steps[0])
    replacement["repair_of"] = "fee_business_type"
    replacement["repair"] = {
        "ok": True,
        "method": "rewrite-second-row-business-type-once",
    }
    replaced = False
    result = []
    for step in steps:
        if step is repair:
            continue
        if not replaced and first_failed_fee_business_type_step([step]):
            result.append(replacement)
            replaced = True
        else:
            result.append(step)
    if not replaced:
        result.insert(0, replacement)
    return result


def summarize_fee_business_type_repair_failure(retry_steps):
    for step in retry_steps or []:
        if not step.get("ok"):
            return str(step.get("reason") or "手续费业务类型重写后仍未通过").strip()
    return "手续费业务类型重写未返回有效步骤"
