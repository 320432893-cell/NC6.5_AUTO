# 职责：后台用阶段性快照批量验证收款单明细字段和最终行数
# 不做什么：不写入 NC，不决定业务字段顺序，不处理保存/暂存
# 允许依赖层：core JAB 操作、tools.receipt_detail_fields/reader
# 谁不应该 import：底层 core、Excel/Sheet 写入模块不应 import
# 生命周期：正式完整流程组件

from __future__ import annotations

import queue
import threading
import time
import uuid

from core.jab_operator import JABOperator
from tools.receipt_body_table_locator import locate_receipt_body_table_cached
from tools.receipt_detail_fields import field_expected_value, field_matches
from tools.receipt_detail_reader import read_body_table_by_path


class DetailPipelineVerifier:
    def __init__(
        self,
        config,
        located,
        flow_started_at=None,
        force_cached_path_fail=False,
        jab=None,
        jab_lock=None,
    ):
        self.config = config
        self.located = self._clone_located(located)
        self.force_cached_path_fail = bool(force_cached_path_fail)
        self.jab = jab
        self.jab_lock = jab_lock
        self.forced_cached_path_report = None
        self.flow_started_at = flow_started_at
        self.started_at = None
        self._queue = queue.Queue()
        self._lock = threading.Lock()
        self._results = {}
        self._submitted = []
        self._pending_fields = {}
        self._snapshots = []
        self._row_count_tasks = {}
        self._stopped = threading.Event()
        self._thread = None

    def start(self):
        self.started_at = time.perf_counter()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def submit_field(self, row_index, field, business):
        task_id = self._new_task_id("field")
        expected = {
            "id": task_id,
            "type": "field",
            "row_index": int(row_index),
            "field": dict(field),
            "expected": field_expected_value(field, business),
            "raw_expected": str(business[field["value_key"]]),
            "submitted_at": time.perf_counter(),
        }
        with self._lock:
            self._submitted.append(task_id)
            self._pending_fields[task_id] = expected
        return task_id

    def submit_snapshot(
        self,
        label,
        max_rows=5,
        timeout=1.2,
        interval=0.08,
        min_matches=0,
    ):
        task_id = self._new_task_id("snap")
        self._queue.put(
            {
                "id": task_id,
                "type": "snapshot",
                "label": label,
                "max_rows": int(max_rows),
                "timeout": float(timeout),
                "interval": float(interval),
                "min_matches": int(min_matches or 0),
                "submitted_at": time.perf_counter(),
            }
        )
        return task_id

    def submit_row_count(self, expected_rows, timeout=1.1, interval=0.06):
        task_id = self._new_task_id("rows")
        task = {
            "id": task_id,
            "type": "row_count",
            "expected_rows": int(expected_rows),
            "timeout": float(timeout),
            "interval": float(interval),
            "submitted_at": time.perf_counter(),
        }
        with self._lock:
            self._submitted.append(task_id)
            self._row_count_tasks[task_id] = task
        self._queue.put(task)
        return task_id

    def wait(self, task_ids, timeout=2.0):
        ids = [task_ids] if isinstance(task_ids, str) else list(task_ids or [])
        deadline = time.perf_counter() + float(timeout)
        while True:
            with self._lock:
                ready = all(task_id in self._results for task_id in ids)
            if ready:
                break
            if time.perf_counter() >= deadline:
                break
            time.sleep(0.03)
        return self.snapshot()

    def snapshot(self):
        with self._lock:
            results = dict(self._results)
            submitted = list(self._submitted)
            pending = dict(self._pending_fields)
            snapshots = list(self._snapshots)
        done = len(results)
        failed = [
            task_id for task_id, result in results.items() if not result.get("ok")
        ]
        return {
            "status": "running" if self._thread and self._thread.is_alive() else "done",
            "submitted": submitted,
            "done": done,
            "pending": len(pending),
            "ok": bool(submitted)
            and done == len(submitted)
            and not pending
            and not failed,
            "failed": failed,
            "results": results,
            "snapshots": snapshots,
            "semantic_table_fallback_used": any(
                (item.get("snapshot") or {}).get("semantic_fallback_used")
                for item in snapshots
            )
            or any(
                ((result.get("snapshot") or {}).get("semantic_fallback_used"))
                for result in results.values()
            ),
            "forced_cached_path_report": self.forced_cached_path_report,
            "started_offset_seconds": round(self.started_at - self.flow_started_at, 3)
            if self.started_at and self.flow_started_at is not None
            else None,
        }

    def close(self, timeout=1.0):
        self._stopped.set()
        self._queue.put(None)
        if self._thread is not None:
            self._thread.join(timeout=max(float(timeout), 0.0))

    def _new_task_id(self, prefix):
        return f"{prefix}-{uuid.uuid4().hex[:10]}"

    def _clone_located(self, located):
        cloned = dict(located or {})
        if isinstance(cloned.get("best"), dict):
            cloned["best"] = dict(cloned["best"])
            if isinstance(cloned["best"].get("window"), dict):
                cloned["best"]["window"] = dict(cloned["best"]["window"])
        return cloned

    def _force_cached_path_fail(self):
        best = self.located.get("best") or {}
        original_path = best.get("path")
        forced_path = f"{original_path}.9999" if original_path else "0.9999"
        best["path"] = forced_path
        self.located["best"] = best
        self.forced_cached_path_report = {
            "ok": True,
            "original_path": original_path,
            "forced_path": forced_path,
        }

    def _run(self):
        jab = self.jab or JABOperator(self.config)
        owns_jab = self.jab is None
        try:
            if owns_jab:
                jab.ensure_started()
            self._with_jab_lock(self._preload_table, jab)
            if self.force_cached_path_fail:
                self._force_cached_path_fail()
            while not self._stopped.is_set():
                task = self._queue.get()
                if task is None:
                    break
                if task.get("type") == "snapshot":
                    self._run_snapshot(jab, task)
                elif task.get("type") == "row_count":
                    result = self._verify_row_count(jab, task)
                    with self._lock:
                        self._results[task["id"]] = result
                else:
                    with self._lock:
                        self._results[task.get("id") or "unknown-task"] = {
                            "ok": False,
                            # 给人看的业务原因；内部任务类型只留作开发诊断字段。
                            "reason": "明细后台校验未执行：收到无法识别的校验请求，本行明细未完成后台核对。",
                            "error_detail": f"unknown_task_type={task.get('type')}",
                        }
        except Exception as exc:
            with self._lock:
                self._results["worker-error"] = {
                    "ok": False,
                    # 给人看的业务原因；异常类型/原文只留作开发诊断字段。
                    "reason": "明细后台校验中断：未能完成本行明细的后台读回核对，请人工核对当前明细行。",
                    "error_detail": f"{type(exc).__name__}: {exc}",
                }
        finally:
            if owns_jab:
                jab.close()

    def _with_jab_lock(self, func, *args, **kwargs):
        if self.jab_lock is None:
            return func(*args, **kwargs)
        with self.jab_lock:
            return func(*args, **kwargs)

    def _preload_table(self, jab):
        started_at = time.perf_counter()
        scope_hwnd = (((self.located or {}).get("best") or {}).get("window") or {}).get(
            "hwnd"
        )
        preloaded = locate_receipt_body_table_cached(
            jab,
            cached=self.located,
            max_rows=5,
            scope_hwnd=scope_hwnd,
        )
        best = preloaded.get("best")
        if best:
            self.located["best"] = best
        with self._lock:
            self._snapshots.append(
                {
                    "id": "preload",
                    "label": "table-preload",
                    "ok": bool(best),
                    "seconds": round(time.perf_counter() - started_at, 3),
                    "cache_hit": bool(preloaded.get("cache_hit")),
                    "fallback_used": bool(preloaded.get("fallback_used")),
                    "path": (best or {}).get("path"),
                    "row_count": (best or {}).get("row_count"),
                    "col_count": (best or {}).get("col_count"),
                }
            )

    def _run_snapshot(self, jab, task):
        started_at = time.perf_counter()
        deadline = started_at + task["timeout"]
        last_snapshot = None
        while True:
            snapshot = self._with_jab_lock(
                read_body_table_by_path,
                jab,
                self.located,
                task["label"],
                max_rows=task["max_rows"],
                semantic_fallback=True,
            )
            last_snapshot = snapshot
            if snapshot.get("ok"):
                matched = self._apply_snapshot_to_pending(task, snapshot, started_at)
                enough_matches = len(matched) >= int(task.get("min_matches") or 0)
                if not enough_matches and time.perf_counter() < deadline:
                    time.sleep(task["interval"])
                    continue
                with self._lock:
                    self._snapshots.append(
                        {
                            "id": task["id"],
                            "label": task["label"],
                            "ok": enough_matches,
                            "seconds": round(time.perf_counter() - started_at, 3),
                            "matched": matched,
                            "min_matches": int(task.get("min_matches") or 0),
                            "snapshot": snapshot,
                            "reason": None
                            if enough_matches
                            else "后台快照匹配字段数不足",
                        }
                    )
                return
            if time.perf_counter() >= deadline:
                with self._lock:
                    self._snapshots.append(
                        {
                            "id": task["id"],
                            "label": task["label"],
                            "ok": False,
                            "seconds": round(time.perf_counter() - started_at, 3),
                            "snapshot": last_snapshot,
                            "reason": "后台快照读取超时",
                        }
                    )
                return
            time.sleep(task["interval"])

    def _apply_snapshot_to_pending(self, task, snapshot, started_at):
        rows = {
            int(row.get("row_index")): (row.get("cells") or {})
            for row in snapshot.get("rows") or []
        }
        matched = []
        with self._lock:
            pending_items = list(self._pending_fields.items())
        for task_id, expected in pending_items:
            field = expected["field"]
            cells = rows.get(expected["row_index"]) or {}
            actual = cells.get(str(field["col"]))
            if not field_matches(actual, expected["raw_expected"], field.get("kind")):
                continue
            result = {
                "ok": True,
                "type": "field",
                "row_index": expected["row_index"],
                "col": field["col"],
                "name": field["name"],
                "expected": expected["expected"],
                "actual": actual,
                "verified_by_snapshot": task["id"],
                "seconds": round(time.perf_counter() - expected["submitted_at"], 3),
                "snapshot": {
                    "step": snapshot.get("step"),
                    "ok": snapshot.get("ok"),
                    "fast_path": snapshot.get("fast_path"),
                    "semantic_fallback_used": snapshot.get("semantic_fallback_used"),
                    "path": snapshot.get("path"),
                    "row_count": snapshot.get("row_count"),
                    "col_count": snapshot.get("col_count"),
                    "path_validation": snapshot.get("path_validation"),
                },
            }
            with self._lock:
                if task_id in self._pending_fields:
                    self._pending_fields.pop(task_id, None)
                    self._results[task_id] = result
                    matched.append(task_id)
        return matched

    def _verify_row_count(self, jab, task):
        started_at = time.perf_counter()
        deadline = started_at + task["timeout"]
        last_snapshot = None
        while True:
            snapshot = self._with_jab_lock(
                read_body_table_by_path,
                jab,
                self.located,
                "row_count_readback",
                max_rows=0,
                semantic_fallback=True,
            )
            last_snapshot = snapshot
            actual_rows = (
                int(snapshot.get("row_count") or 0) if snapshot.get("ok") else 0
            )
            if snapshot.get("ok") and actual_rows == task["expected_rows"]:
                return {
                    "ok": True,
                    "type": "row_count",
                    "expected_rows": task["expected_rows"],
                    "actual_rows": actual_rows,
                    "seconds": round(time.perf_counter() - started_at, 3),
                    "snapshot": snapshot,
                }
            if time.perf_counter() >= deadline:
                return {
                    "ok": False,
                    "type": "row_count",
                    "expected_rows": task["expected_rows"],
                    "actual_rows": actual_rows,
                    "seconds": round(time.perf_counter() - started_at, 3),
                    "snapshot": last_snapshot,
                    "reason": "后台行数验证超时或行数不匹配",
                }
            time.sleep(task["interval"])
