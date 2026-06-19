# 职责：明细 pipeline 校验失败后,用已定位 path 修复字段/行数并提交二次校验
# 不做什么：不重扫表格,不切换旧语义兜底,不做报告渲染,不做整行编排
# 允许依赖层：tools.receipt_flow_entry_state 的 run_with_jab_lock
# 谁不应该 import：core 层模块不应 import；本模块不应反向 import row_runner

import sys

from tools.receipt_flow_entry_state import run_with_jab_lock


class _FlowNamespace:
    # 按调用时从已加载的入口模块取属性：让测试对
    # tools.receipt_full_flow_entry.write_field_once /
    # delete_extra_row_if_present 的 monkeypatch 与拆分前一致地生效，
    # 且不在加载期 import 入口模块以避免成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_full_flow_entry"], name)


_flow = _FlowNamespace()


def force_one_detail_field_pending(report, field_task_ids):
    forced = dict(report or {})
    results = dict(forced.get("results") or {})
    target_id = next(
        (task_id for task_id in field_task_ids if task_id in results), None
    )
    if target_id is None and field_task_ids:
        target_id = field_task_ids[-1]
    if target_id and target_id in results:
        results.pop(target_id, None)
    submitted = list(forced.get("submitted") or [])
    if target_id and target_id not in submitted:
        submitted.append(target_id)
    pending = int(forced.get("pending") or 0)
    forced.update(
        {
            "ok": False,
            "pending": max(1, pending),
            "results": results,
            "submitted": submitted,
            "forced_detail_repair_drill": True,
            "forced_pending_field_task_id": target_id,
        }
    )
    return forced


def repair_detail_pipeline_failures(
    jab,
    jab_lock,
    located,
    pipeline_verifier,
    pipeline_report,
    pipeline_field_tasks,
    pipeline_row_count_task_id,
    expected_rows,
    scope_hwnd,
    recover_after_failure=None,
):
    results = (pipeline_report or {}).get("results") or {}
    best = (located or {}).get("best") or {}
    table_window = best.get("window") or {}
    row_count = int(best.get("row_count") or 0)
    repair = {
        "ok": False,
        "policy": (
            "只用当前已定位的明细表 path 修复一次；不重扫表格，不切换到旧语义兜底"
        ),
        "field_repairs": [],
        "row_count_repair": None,
        "wait_ids": [],
        "snapshot_task_id": None,
    }
    repair_field_ids = []
    for task_id, task in (pipeline_field_tasks or {}).items():
        result = results.get(task_id)
        if result and result.get("ok"):
            continue
        field = task["field"]
        if not best or not table_window:
            attempt = {
                "ok": False,
                "reason": "明细表缓存窗口不可用，不能执行字段修复",
            }
        else:
            attempt = run_with_jab_lock(
                jab_lock,
                _flow.write_field_once,
                jab,
                located,
                table_window,
                int(task["row_index"]),
                row_count,
                field,
                task["business"],
                2,
                current_col=None,
                recover_after_failure=recover_after_failure,
            )
        field_report = {
            "original_task_id": task_id,
            "name": field.get("name"),
            "row_index": int(task["row_index"]),
            "col": field.get("col"),
            "attempt": attempt,
        }
        if attempt.get("ok"):
            verify_task_id = pipeline_verifier.submit_field(
                int(task["row_index"]),
                field,
                task["business"],
            )
            repair_field_ids.append(verify_task_id)
            field_report["verify_task_id"] = verify_task_id
        else:
            field_report["reason"] = (
                attempt.get("input_reason")
                or attempt.get("commit_reason")
                or attempt.get("reason")
            )
        repair["field_repairs"].append(field_report)

    row_count_result = results.get(pipeline_row_count_task_id)
    row_count_needs_repair = bool(pipeline_row_count_task_id) and (
        row_count_result is None or not row_count_result.get("ok")
    )
    row_count_wait_id = None
    if row_count_needs_repair:
        row_repair = run_with_jab_lock(
            jab_lock,
            _flow.delete_extra_row_if_present,
            jab,
            located,
            int(expected_rows),
            scope_hwnd=scope_hwnd,
        )
        repair["row_count_repair"] = row_repair
        if row_repair.get("ok"):
            row_count_wait_id = pipeline_verifier.submit_row_count(int(expected_rows))

    if repair_field_ids:
        repair["snapshot_task_id"] = pipeline_verifier.submit_snapshot(
            "after-detail-repair",
            max_rows=max(3, int(expected_rows) + 1),
            min_matches=len(repair_field_ids),
        )
    wait_ids = [*repair_field_ids]
    if row_count_wait_id:
        wait_ids.append(row_count_wait_id)
    repair["wait_ids"] = wait_ids
    attempted = bool(repair["field_repairs"]) or bool(repair["row_count_repair"])
    if not attempted:
        repair["ok"] = False
        repair["reason"] = "pipeline 失败但没有可修复的字段或行数任务"
    elif not wait_ids:
        repair["ok"] = False
        repair["reason"] = "已尝试修复，但没有成功提交二次校验任务"
    else:
        repair["ok"] = True
    return repair
