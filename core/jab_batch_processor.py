import time
import re
import subprocess
import sys
import itertools
from collections import defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path

from core.data_handler import DataHandler
from core.jab_operator import JABOperator
from core.logger import log
from core.nc_state import NCStateDetector, normalize_generated_voucher
from core.perf import PerfRecorder
from core.run_state import RunStateRecorder
from core.utils import check_abort


class JABBatchProcessor:
    """Batch workflow driven by Excel order and Java Access Bridge table matches."""

    def __init__(
        self,
        config,
        perf_enabled=False,
        perf_label=None,
        command=None,
        generated_date_value=None,
        save_trigger=None,
        hotkey_activate_policy=None,
    ):
        self.cfg = config
        self.batch_cfg = config.get("jab_batch", {})
        self.data_handler = DataHandler(config)
        self.jab = JABOperator(config)
        self.perf = PerfRecorder(enabled=perf_enabled, label=perf_label)
        self.run_state = RunStateRecorder(command=command, config=config)
        self.match_mode = self.batch_cfg.get("match_mode", "exact")
        self.save_strategy = self.batch_cfg.get("save_strategy", "single")
        self.max_batch_size = int(self.batch_cfg.get("max_batch_size", 50))
        self.save_wait = float(self.batch_cfg.get("save_wait", 0.5))
        self.save_trigger = save_trigger or self.batch_cfg.get(
            "save_trigger", "jab_button"
        )
        self.hotkey_activate_policy = hotkey_activate_policy or self.batch_cfg.get(
            "hotkey_activate_policy", "always"
        )
        self.hotkey_save_attempts = 0
        self.wait_between_save_batches = float(
            self.batch_cfg.get("wait_between_save_batches", self.save_wait)
        )
        self.write_generated_status_each_save = bool(
            self.batch_cfg.get("write_generated_status_each_save", True)
        )
        self.use_voucher_queue_cache = bool(
            self.batch_cfg.get("use_voucher_queue_cache", False)
        )
        self.voucher_order_fallback_mode = self.batch_cfg.get(
            "voucher_order_fallback_mode", "strict"
        )
        self.foreign_currency_rate = self.parse_optional_decimal(
            self.batch_cfg.get("foreign_currency_rate")
        )
        self.foreign_currency_rate_tolerance = self.parse_optional_decimal(
            self.batch_cfg.get("foreign_currency_rate_tolerance", 0.02)
        ) or Decimal("0.02")
        self.save_success_timeout = float(
            self.batch_cfg.get("save_success_timeout", 8.0)
        )
        self.generated_status = self.batch_cfg.get("generated_status", "已生成待回填")
        self.voucher_col = int(self.batch_cfg.get("generated_voucher_col", 22))
        self.verify_voucher_advance = self.batch_cfg.get("verify_voucher_advance", True)
        self.voucher_record_timeout = float(
            self.batch_cfg.get("voucher_record_timeout", 8.0)
        )
        self.voucher_window_title = self.batch_cfg.get("voucher_window_title", "制单")
        self.voucher_window_class = self.batch_cfg.get(
            "voucher_window_class", "SunAwtDialog"
        )
        self.generated_date_col = self.batch_cfg.get("generated_date_col", 18)
        self.generated_date_value = (
            generated_date_value
            or self.batch_cfg.get("generated_date_value")
            or date.today().isoformat()
        )
        self.generated_voucher_max = int(
            self.batch_cfg.get("generated_voucher_max", 9999)
        )
        self.config_path = config.get("_config_path", "config.json")
        self.state_detector = NCStateDetector(
            self.jab,
            self.batch_cfg,
            self.generated_date_value,
            self.generated_date_col,
            self.voucher_col,
            self.generated_voucher_max,
            self.record_event,
            self.record_transition,
        )

    def close(self):
        self.jab.close()

    def finish_run_state(self, status, error=None):
        self.run_state.finish(status, error=error)

    def record_event(self, name, **kwargs):
        self.run_state.event(name, **kwargs)
        self.perf.event(name, **kwargs)

    def record_transition(self, event, from_state=None, to_state=None, **kwargs):
        self.record_event(
            "state_transition",
            event=event,
            from_state=from_state,
            to_state=to_state,
            **kwargs,
        )

    def load_pending_items(
        self,
        skip_filled=True,
        skip_any_status=False,
        limit=None,
        start_row=None,
        end_row=None,
    ):
        with self.perf.span(
            "excel_load",
            skip_filled=skip_filled,
            skip_any_status=skip_any_status,
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        ):
            items = self.data_handler.load_jab_batch_data(
                skip_filled=skip_filled,
                skip_any_status=skip_any_status,
            )
        if start_row is not None:
            items = [item for item in items if item["row"] >= start_row]
        if end_row is not None:
            items = [item for item in items if item["row"] <= end_row]
        if limit:
            items = items[:limit]
        self.perf.event(
            "excel_loaded",
            rows=len(items),
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        )
        self.run_state.update_counts(excel_loaded=len(items))
        self.run_state.event(
            "excel_loaded",
            rows=len(items),
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        )
        return items

    def dry_run(self, limit=None, start_row=None, end_row=None):
        self.run_state.set_stage(
            "plan_load_excel",
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        )
        items = self.load_pending_items(
            skip_filled=True,
            skip_any_status=True,
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        )
        parsed_items = [item for item in items if not item.get("parse_error")]
        parse_errors = [item for item in items if item.get("parse_error")]

        self.require_page_state("pending", parsed_items, command="plan")
        self.run_state.set_stage("plan_match_pending_table")
        matches, issues = self.match_current_table(parsed_items)
        batches = self.build_increasing_batches(matches)
        self.run_state.update_counts(
            parse_errors=len(parse_errors),
            matches=len(matches),
            issues=len(issues),
            batches=len(batches),
        )
        self.run_state.set_stage("plan_done")

        self._log_plan(parsed_items, parse_errors, matches, issues, batches)
        return {
            "items": parsed_items,
            "parse_errors": parse_errors,
            "matches": matches,
            "issues": issues,
            "batches": batches,
        }

    def generate_and_save(
        self,
        limit=None,
        max_batches=None,
        start_row=None,
        end_row=None,
    ):
        with self.perf.span(
            "generate_total",
            limit=limit,
            max_batches=max_batches,
            start_row=start_row,
            end_row=end_row,
        ):
            self.run_state.set_stage(
                "generate_load_excel",
                limit=limit,
                max_batches=max_batches,
                start_row=start_row,
                end_row=end_row,
            )
            items = self.load_pending_items(
                skip_filled=True,
                skip_any_status=True,
                limit=limit,
                start_row=start_row,
                end_row=end_row,
            )
            pending = [item for item in items if not item.get("parse_error")]
            parse_errors = [item for item in items if item.get("parse_error")]
            with self.perf.span("excel_save_split_columns", rows=len(pending)):
                self.data_handler.save_jab_split_columns(pending)
            self.perf.event(
                "generate_items_prepared",
                total=len(items),
                pending=len(pending),
                parse_errors=len(parse_errors),
            )
            self.run_state.update_counts(
                generate_items_total=len(items),
                generate_pending=len(pending),
                parse_errors=len(parse_errors),
            )
            if parse_errors:
                self.run_state.set_stage(
                    "generate_write_parse_errors",
                    excel_rows=[item["row"] for item in parse_errors],
                )
                with self.perf.span("excel_save_parse_errors", rows=len(parse_errors)):
                    self.data_handler.save_jab_results(
                        {
                            item["row"]: f"格式错误-{item['parse_error']}"
                            for item in parse_errors
                        }
                    )

            total_saved = 0
            total_batches = 0
            issue_updates = {}

            if pending:
                check_abort()
                self.require_page_state("pending", pending, command="generate")
                self.run_state.set_stage(
                    "generate_match_pending_table",
                    excel_rows=[item["row"] for item in pending],
                )
                matches, issues = self.match_current_table(pending)
                self.perf.event(
                    "pending_match_done",
                    pending=len(pending),
                    matches=len(matches),
                    issues=len(issues),
                )
                if issues:
                    issue_updates.update(self.format_issue_updates(issues))

                if matches:
                    self.run_state.update_counts(
                        pending_matches=len(matches),
                        pending_issues=len(issues),
                    )
                    log.info(
                        "开始执行 JAB 全量生成: "
                        f"size={len(matches)} excel_rows={[m['item']['row'] for m in matches]} "
                        f"nc_rows={[m['nc_row'] for m in matches]}"
                    )
                    saved_matches, total_batches = self.process_full_selection(
                        matches,
                        max_save_batches=max_batches,
                    )
                    total_saved = len(saved_matches)

            if issue_updates:
                self.run_state.set_stage(
                    "generate_write_issue_updates",
                    rows=len(issue_updates),
                )
                with self.perf.span(
                    "excel_save_issue_updates", rows=len(issue_updates)
                ):
                    self.data_handler.save_jab_results(issue_updates)

            self.perf.event(
                "generate_done",
                batches=total_batches,
                saved=total_saved,
                perf_path=str(self.perf.path) if self.perf.path else "",
            )
            self.run_state.update_counts(
                save_batches=total_batches,
                saved=total_saved,
            )
            self.run_state.set_stage("generate_done")
            log.info(f"JAB 生成保存完成: batches={total_batches}, saved={total_saved}")
            return total_saved

    def backfill_generated_vouchers(self, limit=None, start_row=None, end_row=None):
        with self.perf.span(
            "backfill_total",
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        ):
            self.run_state.set_stage(
                "backfill_load_excel",
                limit=limit,
                start_row=start_row,
                end_row=end_row,
            )
            items = self.load_pending_items(
                skip_filled=False,
                skip_any_status=False,
                limit=limit,
                start_row=start_row,
                end_row=end_row,
            )
            items = [
                item
                for item in items
                if not item.get("parse_error")
                and str(item.get("voucher") or "").strip() == self.generated_status
            ]
            with self.perf.span("excel_save_split_columns", rows=len(items)):
                self.data_handler.save_jab_split_columns(items)
            self.perf.event("backfill_items_prepared", rows=len(items))
            self.run_state.update_counts(backfill_pending=len(items))
            if not items:
                self.perf.event(
                    "backfill_done",
                    vouchers=0,
                    issues=0,
                    updates=0,
                    perf_path=str(self.perf.path) if self.perf.path else "",
                )
                self.run_state.set_stage("backfill_done")
                log.info("JAB 回填完成: 没有待回填行")
                return {}
            self.run_state.set_stage(
                "backfill_match_generated_table",
                excel_rows=[item["row"] for item in items],
            )
            self.require_page_state("generated", items, command="backfill")
            matches, issues = self.match_current_table(
                items,
                voucher_col=self.voucher_col,
                prefer_generated_date=True,
            )
            self.perf.event(
                "backfill_match_done",
                items=len(items),
                matches=len(matches),
                issues=len(issues),
            )

            updates = {}
            for match in matches:
                raw_voucher = str(match["row_data"].get("voucher_text", "")).strip()
                voucher = self.normalize_generated_voucher(raw_voucher)
                if voucher is not None:
                    updates[match["item"]["row"]] = voucher
                    self.perf.event(
                        "backfill_voucher_match",
                        excel_row=match["item"]["row"],
                        voucher=voucher,
                        generated_row=match["row_data"].get("row_index"),
                        amount=str(match["item"]["amount"]),
                        partner=match["item"]["partner"],
                    )
                elif raw_voucher:
                    updates[match["item"]["row"]] = f"凭证号异常-{raw_voucher}"
                else:
                    updates[match["item"]["row"]] = "已生成未取到凭证号"

            updates.update(self.format_issue_updates(issues, prefix="回填"))
            self.run_state.set_stage("backfill_write_excel", rows=len(updates))
            with self.perf.span("backfill_excel_save", rows=len(updates)):
                self.data_handler.save_jab_results(updates)
            self.perf.event(
                "backfill_done",
                vouchers=len(matches),
                issues=len(issues),
                updates=len(updates),
                perf_path=str(self.perf.path) if self.perf.path else "",
            )
            self.run_state.update_counts(
                backfill_matches=len(matches),
                backfill_issues=len(issues),
                backfill_updates=len(updates),
            )
            self.run_state.set_stage("backfill_done")
            log.info(f"JAB 回填完成: vouchers={len(matches)}, issues={len(issues)}")
            return updates

    def switch_to_generated_list(self):
        with self.perf.span("switch_generated_total"):
            self.perf.event(
                "switch_generated_date",
                generated_date_value=self.generated_date_value,
            )
            self.run_state.set_stage("switch_detect_state")
            state = self.get_nc_workflow_state()
            if state == "voucher_empty":
                close_cfg = self.batch_cfg.get("close_voucher_window", {})
                with self.perf.span("switch_close_empty_voucher_window"):
                    self.jab.close_window_by_title(
                        close_cfg.get("title", self.voucher_window_title),
                        class_name=close_cfg.get(
                            "class_name", self.voucher_window_class
                        ),
                        wait=float(close_cfg.get("wait", 0.5)),
                    )
                state = self.get_nc_workflow_state()

            if state != "parent_ready":
                raise RuntimeError(
                    f"当前不能切换正式单据，必须先回到制单父界面: state={state}"
                )

            open_query = self.batch_cfg.get("open_query", {})
            query_method = open_query.get("method")
            self.run_state.set_stage(
                "switch_find_query_window",
                query_method=query_method,
                nc_state=state,
            )
            query_hwnd = self.find_query_window(open_query, timeout=0.5)
            if query_method == "jab_action":
                self.run_state.set_stage("switch_open_query_jab_action")
                self.record_event(
                    "event_query_open_start",
                    method="jab_action",
                    existing=bool(query_hwnd),
                )
                with self.perf.span("switch_open_query_jab_action"):
                    if query_hwnd:
                        log.info("查询窗口已存在，跳过 JAB 查询入口触发")
                    else:
                        try:
                            self.open_query_with_jab_action(open_query)
                        except RuntimeError:
                            if open_query.get("fallback_method") != "hotkey":
                                raise
                            log.warning("JAB 查询入口失败，回退 F3 热键")
                        query_hwnd = self.find_query_window(
                            open_query,
                            timeout=float(open_query.get("timeout", 5)),
                        )
                    if not query_hwnd and open_query.get("fallback_method") == "hotkey":
                        log.warning("JAB 查询入口未打开查询窗口，回退 F3 热键")
                        self.open_query_with_hotkey(open_query)
                        query_hwnd = self.find_query_window(
                            open_query,
                            timeout=float(open_query.get("timeout", 5)),
                        )
                    if not query_hwnd:
                        raise RuntimeError("未检测到查询窗口")
                self.record_transition(
                    "query_opened",
                    from_state="pending",
                    to_state="query_open",
                    method="jab_action",
                )
            elif query_method == "hotkey":
                self.run_state.set_stage("switch_open_query_hotkey")
                self.record_event(
                    "event_query_open_start",
                    method="hotkey",
                    existing=bool(query_hwnd),
                )
                with self.perf.span("switch_open_query"):
                    if not query_hwnd:
                        self.open_query_with_hotkey(open_query)
                        query_hwnd = self.find_query_window(
                            open_query,
                            timeout=float(open_query.get("timeout", 5)),
                        )
                    if not query_hwnd:
                        raise RuntimeError("按快捷键后未检测到查询窗口")
                self.record_transition(
                    "query_opened",
                    from_state="pending",
                    to_state="query_open",
                    method="hotkey",
                )
            elif query_method:
                raise RuntimeError(f"不支持的 open_query.method: {query_method}")

            steps = self.batch_cfg.get("switch_generated_steps", [])
            if not steps:
                raise RuntimeError(
                    "未配置 switch_generated_steps，暂不能自动切换到已生成列表"
                )
            try:
                self.run_state.set_stage("switch_run_query_steps")
                self.require_page_state(
                    "query_open",
                    items=None,
                    command="switch-generated",
                )
                self.record_event(
                    "event_query_confirm_start",
                    steps=len(steps),
                    generated_date_value=self.generated_date_value,
                )
                with self.perf.span("switch_run_steps", steps=len(steps)):
                    self.run_switch_generated_steps(
                        open_query,
                        steps,
                        query_method,
                        query_hwnd=query_hwnd,
                    )
                self.record_transition(
                    "query_confirmed",
                    from_state="query_open",
                    to_state="loading",
                    generated_date_value=self.generated_date_value,
                )
            except RuntimeError:
                if (
                    query_method == "jab_action"
                    and open_query.get("fallback_method") == "hotkey"
                ):
                    log.warning("JAB 查询窗口步骤失败，回退 F3 重新打开查询窗口")
                    self.open_query_with_hotkey(open_query)
                    query_hwnd = self.find_query_window(
                        open_query,
                        timeout=float(open_query.get("timeout", 5)),
                    )
                    if not query_hwnd:
                        raise RuntimeError("F3 回退后未检测到查询窗口")
                    with self.perf.span("switch_run_steps_fallback", steps=len(steps)):
                        self.run_switch_generated_steps(
                            open_query,
                            steps,
                            "hotkey",
                            query_hwnd=query_hwnd,
                        )
                else:
                    raise
            with self.perf.span("switch_generated_snapshot"):
                self.run_state.set_stage("switch_verify_generated_snapshot")
                state = self.require_page_state(
                    "generated",
                    items=None,
                    command="switch-generated",
                )
            self.record_transition(
                "generated_list_loaded",
                from_state="loading",
                to_state="generated",
                rows=state.table.get("row_count", 0) if state.table else 0,
            )
            rows = state.table.get("row_count", 0) if state.table else 0
            with_voucher = state.table.get("voucher_values", []) if state.table else []
            log.info(
                "已切换到已生成列表: "
                f"rows={rows} sample_voucher_count={len(with_voucher)}"
            )
            self.run_state.update_counts(generated_snapshot_rows=rows)
            self.run_state.set_stage("switch_generated_done")
            return True

    def run_switch_generated_steps(
        self,
        open_query,
        steps,
        query_method,
        query_hwnd=None,
    ):
        dialog_title = open_query.get("dialog_title", "查询")
        dialog_class = open_query.get("dialog_class", "SunAwtDialog")

        try:
            for index, step in enumerate(steps, start=1):
                step_name = self.get_switch_step_name(step, index)
                if isinstance(step, dict) and step.get("runner") == "subprocess":
                    subprocess_step = {
                        **step,
                        "title": dialog_title,
                        "class_name": dialog_class,
                    }
                    with self.perf.span(
                        f"switch_step_{step_name}",
                        step_index=index,
                        path=step.get("path"),
                        control_name=step.get("name"),
                        role=step.get("role"),
                        runner="subprocess",
                    ):
                        self.run_jab_action_subprocess(subprocess_step)
                        time.sleep(float(step.get("wait", 0.2)))
                else:
                    with self.perf.span(
                        f"switch_step_{step_name}",
                        step_index=index,
                        path=step.get("path") if isinstance(step, dict) else None,
                        control_name=step.get("name")
                        if isinstance(step, dict)
                        else None,
                        role=step.get("role") if isinstance(step, dict) else None,
                    ):
                        self.run_query_window_step(
                            step,
                            dialog_title=dialog_title,
                            dialog_class=dialog_class,
                        )
            return True
        except RuntimeError as exc:
            if query_method == "jab_action":
                raise
            log.warning(
                "查询窗口内步骤失败，回退全局控件查找: "
                f"window={dialog_title!r}/{dialog_class!r}, error={exc}"
            )
            self.jab.run_named_steps(steps)
            return True

    def get_switch_step_name(self, step, index):
        if not isinstance(step, dict):
            return f"{index}"
        if step.get("perf_name"):
            return str(step["perf_name"])
        name = step.get("name")
        set_text = step.get("set_text")
        path = step.get("path")
        if name == "正式单据":
            return "formal_action"
        if name == "确定":
            return "confirm_action"
        if set_text == "{generated_date_value}" and path:
            if str(path).endswith(".1.0.0"):
                return "date_from"
            if str(path).endswith(".1.2.0"):
                return "date_to"
            return "date_text"
        return f"{index}"

    def run_query_window_step(self, step, dialog_title, dialog_class):
        path = step.get("path") if isinstance(step, dict) else None
        if not path:
            self.jab.run_named_steps_in_window(
                [step],
                window_title=dialog_title,
                window_class=dialog_class,
                visible_only=True,
            )
            return

        wait_info = self.jab.wait_context_by_path(
            path,
            title=dialog_title,
            class_name=dialog_class,
            name=step.get("name"),
            role=step.get("role"),
            require_showing=bool(step.get("require_showing", False)),
            timeout=float(step.get("guard_timeout", step.get("timeout", 3))),
        )
        if not wait_info:
            raise RuntimeError(
                f"查询窗口控件未就绪: path={path} name={step.get('name')}"
            )

        set_text = step.get("set_text")
        if set_text == "{generated_date_value}":
            set_text = self.generated_date_value
        if set_text is not None:
            ok = self.jab.set_text_by_path(
                path,
                set_text,
                title=dialog_title,
                class_name=dialog_class,
                name=step.get("name"),
                role=step.get("role"),
                wait=float(step.get("wait", 0.0)),
                timeout=float(step.get("timeout", 1)),
                require_showing=bool(step.get("require_showing", False)),
            )
        else:
            ok = self.jab.do_action_by_path(
                path,
                title=dialog_title,
                class_name=dialog_class,
                name=step.get("name"),
                role=step.get("role"),
                action_name=step.get("action"),
                click_mode=step.get("click_mode"),
                wait=float(step.get("wait", 0.0)),
                timeout=float(step.get("timeout", 1)),
                require_showing=bool(step.get("require_showing", False)),
            )
        if not ok:
            raise RuntimeError(f"查询窗口步骤失败: path={path} name={step.get('name')}")

    def find_query_window(self, open_query, timeout):
        return self.jab.wait_window_by_title(
            open_query.get("dialog_title", "查询"),
            class_name=open_query.get("dialog_class", "SunAwtDialog"),
            timeout=timeout,
            include_children=bool(open_query.get("dialog_include_children", True)),
            visible_only=bool(open_query.get("dialog_visible_only", True)),
        )

    def open_query_with_jab_action(self, open_query):
        if open_query.get("runner") == "subprocess":
            self.run_jab_action_subprocess(open_query)
            return

        if (
            bool(open_query.get("wait_dialog", True))
            and open_query.get("click_mode") != "bounds"
        ):
            thread = self._trigger_query_action_path(open_query)
            if not thread:
                raise RuntimeError(f"JAB 查询入口未找到: path={open_query.get('path')}")
            return

        if not self._do_query_action_path(open_query):
            raise RuntimeError(f"JAB 查询入口执行失败: path={open_query.get('path')}")

    def _trigger_query_action_path(self, open_query):
        path = open_query.get("path")
        kwargs = {
            "name": open_query.get("name"),
            "role": open_query.get("role"),
            "action_name": open_query.get("action"),
            "timeout": float(open_query.get("return_timeout", 0.2)),
            "require_showing": bool(open_query.get("require_showing", False)),
        }
        attempts = [
            {
                "title": open_query.get("window_title"),
                "class_name": open_query.get("window_class"),
            },
            {
                "title": open_query.get("main_title"),
                "class_name": open_query.get("main_class"),
            },
            {"title": None, "class_name": None},
        ]
        for attempt in attempts:
            thread = self.jab.trigger_action_by_path_async(path, **kwargs, **attempt)
            if thread:
                return thread
        return None

    def _do_query_action_path(self, open_query):
        path = open_query.get("path")
        kwargs = {
            "name": open_query.get("name"),
            "role": open_query.get("role"),
            "action_name": open_query.get("action"),
            "click_mode": open_query.get("click_mode"),
            "wait": float(open_query.get("wait", 0.8)),
            "timeout": float(open_query.get("timeout", 5)),
            "require_showing": bool(open_query.get("require_showing", False)),
        }
        attempts = [
            {
                "title": open_query.get("window_title"),
                "class_name": open_query.get("window_class"),
            },
            {
                "title": open_query.get("main_title"),
                "class_name": open_query.get("main_class"),
            },
            {"title": None, "class_name": None},
        ]
        for attempt in attempts:
            if self.jab.do_action_by_path(path, **kwargs, **attempt):
                return True
        return False

    def run_jab_action_subprocess(self, step):
        path = step.get("path")
        if not path:
            raise RuntimeError("JAB 子进程动作缺少 path")

        timeout = float(step.get("process_timeout", 1.0))
        action_timeout = float(step.get("timeout", 3.0))
        action_wait = float(step.get("action_wait", 0.0))
        script = Path(__file__).resolve().parents[1] / "tools" / "jab_action_once.py"
        cmd = [
            sys.executable,
            str(script),
            "--config",
            self.config_path,
            "--path",
            str(path),
            "--timeout",
            str(action_timeout),
            "--wait",
            str(action_wait),
        ]

        title = step.get("title", step.get("window_title"))
        class_name = step.get("class_name", step.get("window_class"))
        if title is not None:
            cmd.extend(["--title", str(title)])
        if class_name is not None:
            cmd.extend(["--class-name", str(class_name)])
        if step.get("name") is not None:
            cmd.extend(["--name", str(step.get("name"))])
        if step.get("role") is not None:
            cmd.extend(["--role", str(step.get("role"))])
        if step.get("action") is not None:
            cmd.extend(["--action", str(step.get("action"))])
        click_mode = step.get("click_mode")
        if click_mode is not None:
            cmd.extend(["--click-mode", str(click_mode)])

        set_text = step.get("set_text")
        if set_text == "{generated_date_value}":
            set_text = self.generated_date_value
        if set_text is not None:
            cmd.extend(["--set-text", str(set_text)])
        if step.get("set_text_near_label") is not None:
            cmd.extend(["--set-text-near-label", str(step.get("set_text_near_label"))])
        if step.get("guard_path") is not None:
            cmd.extend(["--guard-path", str(step.get("guard_path"))])
        if step.get("guard_name") is not None:
            cmd.extend(["--guard-name", str(step.get("guard_name"))])
        if step.get("guard_role") is not None:
            cmd.extend(["--guard-role", str(step.get("guard_role"))])

        if bool(step.get("require_showing", False)):
            cmd.append("--require-showing")

        log.info(
            "启动 JAB action 子进程: "
            f"path={path} name={step.get('name')} "
            f"set_text={set_text is not None} timeout={timeout}"
        )
        try:
            result = subprocess.run(
                cmd,
                cwd=str(Path(__file__).resolve().parents[1]),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired:
            log.warning(
                "JAB action 子进程超时，按 UI 状态继续判断: "
                f"path={path} name={step.get('name')} timeout={timeout}"
            )
            return False

        output = ((result.stdout or "") + (result.stderr or "")).strip()
        if result.returncode != 0:
            log.warning(
                "JAB action 子进程未确认成功，按 UI 状态继续判断: "
                f"returncode={result.returncode} output={output[:500]}"
            )
            return False

        log.info(f"JAB action 子进程完成: output={output[:500]}")
        return True

    def open_query_with_hotkey(self, open_query):
        self.jab.activate_window_by_title(
            open_query.get("main_title", ""),
            class_name=open_query.get("main_class"),
            timeout=float(open_query.get("timeout", 5)),
        )
        self.jab.press_key(
            open_query.get("key", "f3"),
            wait=float(open_query.get("wait", 0.8)),
        )

    def get_nc_workflow_state(self):
        voucher_exists = self.jab.window_exists(
            self.voucher_window_title,
            class_name=self.voucher_window_class,
        )
        if voucher_exists:
            tables = self.jab.read_window_table_cells(
                self.voucher_window_title,
                max_rows=20,
                max_cols=13,
            )
            voucher_tables = [
                table
                for table in tables
                if table["row_count"] > 0 and table["col_count"] == 13
            ]
            if voucher_tables:
                rows = sum(table["row_count"] for table in voucher_tables)
                log.info(f"NC 状态: 制单子窗口打开，制单表 rows={rows}")
                return "voucher_open"
            log.info("NC 状态: 制单子窗口打开，但制单表为空/不可读")
            return "voucher_empty"

        main_title = self.batch_cfg.get("open_query", {}).get("main_title", "")
        main_class = self.batch_cfg.get("open_query", {}).get("main_class")
        if not main_title or self.jab.window_exists(main_title, class_name=main_class):
            log.info("NC 状态: 已回到制单父界面")
            return "parent_ready"

        log.warning("NC 状态: 未检测到制单子窗口，也未确认父界面")
        return "unknown"

    def require_page_state(self, expected, items=None, command=""):
        return self.state_detector.require_page_state(expected, items, command)

    def detect_page_state(self, items=None):
        return self.state_detector.detect_page_state(items)

    def detect_voucher_window_state(self):
        return self.state_detector.detect_voucher_window_state()

    def collect_page_controls(self):
        return self.state_detector.probe.collect_named_controls(
            ("单据生成", "查询", "生成", "前台生成", "正式单据")
        )

    def collect_window_controls(self, window_title, window_class, names):
        return self.state_detector.probe.collect_named_controls(
            names,
            window_title=window_title,
            window_class=window_class,
        )

    def read_page_table_signatures(self):
        return self.state_detector.probe.read_page_table_signatures(
            self.generated_date_col,
            self.voucher_col,
            self.jab.amount_col,
            self.jab.partner_col,
        )

    def describe_signature_table(self, table):
        date_values = self.sample_table_col(table, self.generated_date_col)
        voucher_values = self.sample_table_col(table, self.voucher_col)
        return {
            "table_index": table["table_index"],
            "window_title": table.get("window_title"),
            "window_class": table.get("window_class"),
            "row_count": table["row_count"],
            "col_count": table["col_count"],
            "date_values": date_values,
            "voucher_values": voucher_values,
            "rows": [
                {
                    "row_index": row["row_index"],
                    "amount": self.jab.normalize_amount(
                        row["cells"][self.jab.amount_col]
                        if self.jab.amount_col < len(row["cells"])
                        else ""
                    ),
                    "partner": self.jab.normalize_text(
                        row["cells"][self.jab.partner_col]
                        if self.jab.partner_col < len(row["cells"])
                        else ""
                    ),
                }
                for row in table.get("rows", [])
            ],
        }

    def sample_table_col(self, table, col):
        if col is None:
            return []
        values = []
        for row in table.get("rows", []):
            cells = row.get("cells", [])
            if 0 <= col < len(cells):
                text = str(cells[col]).strip()
                if text:
                    values.append(text)
        return values

    def choose_main_signature_table(self, tables):
        from core.nc_state import choose_main_signature_table

        return choose_main_signature_table(tables)

    def is_generated_signature(self, table, require_formal=True):
        return self.state_detector.is_generated_signature(table, require_formal)

    def is_pending_signature(self, table, visible_names):
        from core.nc_state import is_pending_signature

        return is_pending_signature(table, visible_names)

    def table_match_ratio(self, rows, items):
        return self.state_detector.table_match_ratio(rows, items)

    def looks_loading(self, controls, tables):
        from core.nc_state import looks_loading

        return looks_loading(controls, tables)

    def normalize_generated_voucher(self, raw_voucher):
        return normalize_generated_voucher(raw_voucher, self.generated_voucher_max)

    def match_current_table(self, items, voucher_col=None, prefer_generated_date=False):
        extra_cols = [self.generated_date_col] if prefer_generated_date else None
        with self.perf.span(
            "pending_snapshot_read",
            items=len(items),
            voucher_col=voucher_col,
            prefer_generated_date=prefer_generated_date,
        ):
            snapshot = self.jab.read_table_snapshot(
                voucher_col=voucher_col,
                extra_cols=extra_cols,
            )
        self.perf.event("pending_snapshot_loaded", rows=len(snapshot))
        self.run_state.event(
            "table_snapshot_loaded",
            rows=len(snapshot),
            voucher_col=voucher_col,
            prefer_generated_date=prefer_generated_date,
        )
        index = defaultdict(list)
        for row in snapshot:
            if row["amount"] is None or not row["partner"]:
                continue
            index[(row["amount"], row["partner"])].append(row)

        matches = []
        issues = []
        for item in items:
            key = (
                self._as_decimal(item["amount"]),
                self.jab.normalize_text(item["partner"]),
            )
            rows = index.get(key, [])
            if len(rows) > 1 and prefer_generated_date:
                dated_rows = self.filter_generated_date_rows(rows)
                if dated_rows:
                    rows = dated_rows
            if len(rows) == 1:
                matches.append(
                    {
                        "item": item,
                        "nc_row": rows[0]["row_index"],
                        "row_data": rows[0],
                    }
                )
            elif not rows and self.match_mode == "contains":
                contains_rows = self._find_contains(snapshot, key)
                self._append_match_or_issue(matches, issues, item, contains_rows)
            else:
                issues.append(
                    {
                        "item": item,
                        "reason": "未找到" if not rows else f"重复{len(rows)}条",
                        "rows": [row["row_index"] for row in rows],
                    }
                )

        return matches, issues

    def filter_generated_date_rows(self, rows):
        if self.generated_date_col is None or not self.generated_date_value:
            return []

        target = str(self.generated_date_value).strip()
        dated_rows = []
        for row in rows:
            text = str(
                row.get("extra_text", {}).get(self.generated_date_col, "")
            ).strip()
            if text == target:
                dated_rows.append(row)
        return dated_rows

    def build_increasing_batches(self, matches):
        batches = []
        current = []
        last_nc_row = None

        for match in matches:
            nc_row = match["nc_row"]
            should_split = current and (
                nc_row <= last_nc_row
                or (self.max_batch_size > 0 and len(current) >= self.max_batch_size)
            )
            if should_split:
                batches.append(current)
                current = []
                last_nc_row = None

            current.append(match)
            last_nc_row = nc_row

        if current:
            batches.append(current)
        return batches

    def process_full_selection(self, matches, max_save_batches=None):
        rows = [match["nc_row"] for match in matches]
        self.run_state.set_stage(
            "pending_rows_select",
            excel_rows=[match["item"]["row"] for match in matches],
            nc_rows=rows,
        )
        self.perf.event(
            "full_selection_start",
            matches=len(matches),
            excel_rows=[match["item"]["row"] for match in matches],
            nc_rows=rows,
        )
        self.record_event(
            "event_pending_rows_select_start",
            rows=len(rows),
            excel_rows=[match["item"]["row"] for match in matches],
            nc_rows=rows,
        )
        with self.perf.span("pending_rows_select", rows=len(rows)):
            if not self.jab.select_table_rows(rows):
                raise RuntimeError(f"选中 NC 行失败: {rows}")
        self.record_event(
            "event_pending_rows_selected",
            rows=len(rows),
            nc_rows=rows,
        )

        self.run_state.set_stage("front_generate_click", nc_rows=rows)
        self.record_event("event_front_generate_click", rows=len(rows), nc_rows=rows)
        with self.perf.span("front_generate_click", rows=len(rows)):
            if not self.jab.do_generate_front():
                raise RuntimeError("点击 生成 -> 前台生成 失败")
        self.record_transition(
            "front_generate_clicked",
            from_state="pending",
            to_state="voucher_open",
            rows=len(rows),
        )

        self.run_state.set_stage("front_generate_wait")
        with self.perf.span("front_generate_wait", wait=self.save_wait):
            time.sleep(self.save_wait)

        pending = list(matches)
        self.require_page_state("voucher_open", pending, command="generate")
        saved_matches = []
        save_batches = 0

        while pending:
            check_abort()
            self.run_state.set_stage(
                "voucher_match",
                pending_excel_rows=[match["item"]["row"] for match in pending],
                save_batch_index=save_batches + 1,
            )
            with self.perf.span(
                "voucher_match",
                pending=len(pending),
                save_batch_index=save_batches + 1,
            ):
                voucher_matches, issues = self.match_voucher_table(pending)
            self.perf.event(
                "voucher_match_done",
                pending=len(pending),
                matches=len(voucher_matches),
                issues=len(issues),
                save_batch_index=save_batches + 1,
            )
            if issues:
                detail = "; ".join(
                    f"Excel行{issue['item']['row']} {issue['reason']}"
                    for issue in issues
                )
                raise RuntimeError(f"制单表匹配失败: {detail}")

            new_saved, new_batches = self.save_current_voucher_matches(voucher_matches)
            saved_matches.extend(new_saved)
            saved_rows = {match["item"]["row"] for match in new_saved}
            if max_save_batches and save_batches >= max_save_batches:
                raise RuntimeError(
                    "已达到 max_batches，但全量生成模式不能中途留下制单窗口，"
                    "请不要在全量生成时使用 max_batches"
                )
            pending = [
                match for match in pending if match["item"]["row"] not in saved_rows
            ]
            save_batches += new_batches
            if pending:
                with self.perf.span(
                    "save_batch_wait",
                    wait=self.save_wait,
                    save_batch_index=save_batches,
                ):
                    time.sleep(self.save_wait)

        self.run_state.set_stage(
            "pending_final_verify",
            saved_excel_rows=[match["item"]["row"] for match in saved_matches],
        )
        with self.perf.span("pending_final_verify", rows=len(saved_matches)):
            self.close_and_verify_pending_removed(saved_matches)
        return saved_matches, save_batches

    def resume_current_voucher_window(self, limit=None, start_row=None, end_row=None):
        with self.perf.span(
            "resume_voucher_total",
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        ):
            self.run_state.set_stage(
                "resume_load_excel",
                limit=limit,
                start_row=start_row,
                end_row=end_row,
            )
            items = self.load_pending_items(
                skip_filled=True,
                skip_any_status=True,
                limit=limit,
                start_row=start_row,
                end_row=end_row,
            )
            pending = [item for item in items if not item.get("parse_error")]
            if not pending:
                log.info("当前没有可用于恢复保存的 Excel 行")
                return 0

            self.run_state.set_stage(
                "resume_read_voucher_window",
                excel_rows=[item["row"] for item in pending],
            )
            tables = self.read_voucher_tables(len(pending))
            matches = [
                {
                    "item": item,
                    "nc_row": None,
                    "row_data": {},
                }
                for item in pending
            ]
            voucher_matches, issues = self.match_voucher_table(matches, tables=tables)
            if issues:
                table_rows = sum(table["row_count"] for table in tables)
                if len(voucher_matches) != table_rows:
                    detail = "; ".join(
                        f"Excel行{issue['item']['row']} {issue['reason']}"
                        for issue in issues
                    )
                    raise RuntimeError(f"恢复制单表匹配失败: {detail}")
                log.warning(
                    "恢复保存时忽略未进入当前制单窗口的 Excel 行: "
                    f"excel_rows={[issue['item']['row'] for issue in issues]}"
                )

            saved_matches, save_batches = self.save_current_voucher_matches(
                voucher_matches
            )
            self.run_state.set_stage(
                "resume_pending_final_verify",
                saved_excel_rows=[match["item"]["row"] for match in saved_matches],
            )
            with self.perf.span("resume_pending_final_verify", rows=len(saved_matches)):
                self.close_and_verify_pending_removed(saved_matches)
            log.info(
                f"JAB 恢复制单窗口保存完成: batches={save_batches}, saved={len(saved_matches)}"
            )
            self.run_state.update_counts(
                resume_save_batches=save_batches,
                resume_saved=len(saved_matches),
            )
            self.run_state.set_stage("resume_done")
            return len(saved_matches)

    def save_current_voucher_matches(self, voucher_matches):
        pending_source = list(voucher_matches)
        saved_matches = []
        pending_status_updates = {}
        cached_voucher_matches = None
        save_batches = 0

        while pending_source:
            check_abort()
            if cached_voucher_matches is not None:
                refreshed_matches = cached_voucher_matches
                self.perf.event(
                    "voucher_queue_cache_used",
                    pending=len(refreshed_matches),
                    save_batch_index=save_batches + 1,
                )
            else:
                tables = self.read_voucher_tables(len(pending_source))
                refreshed_matches, issues = self.match_voucher_table(
                    pending_source,
                    tables=tables,
                )
                if issues:
                    detail = "; ".join(
                        f"Excel行{issue['item']['row']} {issue['reason']}"
                        for issue in issues
                    )
                    raise RuntimeError(f"刷新制单表匹配失败: {detail}")
                if self.should_use_voucher_queue_cache(refreshed_matches):
                    cached_voucher_matches = refreshed_matches
                    self.perf.event(
                        "voucher_queue_cache_started",
                        pending=len(cached_voucher_matches),
                        save_batch_index=save_batches + 1,
                    )

            pending = refreshed_matches
            voucher_batches = self.build_voucher_save_batches(pending)
            voucher_batch = voucher_batches[0]
            before_count = voucher_batch[0]["table_rows"]
            row_indexes = [match["voucher_row"] for match in voucher_batch]
            selection_row_indexes = self.get_voucher_selection_rows(voucher_batch)
            self.run_state.set_stage(
                "voucher_rows_select",
                save_batch_index=save_batches + 1,
                excel_rows=[match["item"]["row"] for match in voucher_batch],
                voucher_rows=row_indexes,
                select_rows=selection_row_indexes,
                strategy=self.save_strategy,
            )
            log.info(
                "保存制单批次: "
                f"size={len(voucher_batch)} excel_rows={[m['item']['row'] for m in voucher_batch]} "
                f"voucher_rows={row_indexes} select_rows={selection_row_indexes} "
                f"strategy={self.save_strategy} before_count={before_count}"
            )

            with self.perf.span(
                "voucher_rows_select",
                rows=len(selection_row_indexes),
                table_index=voucher_batch[0]["table_index"],
                logical_rows=row_indexes,
                select_rows=selection_row_indexes,
                save_batch_index=save_batches + 1,
            ):
                if not self.jab.select_visible_table_rows(
                    voucher_batch[0]["table_index"],
                    selection_row_indexes,
                    window_title=self.voucher_window_title,
                ):
                    raise RuntimeError(f"选中制单表行失败: {selection_row_indexes}")

            self.run_state.set_stage(
                "voucher_save_click",
                save_batch_index=save_batches + 1,
                excel_rows=[match["item"]["row"] for match in voucher_batch],
            )
            self.require_page_state(
                "voucher_open",
                voucher_batch,
                command="voucher-save",
            )
            self.record_event(
                "event_voucher_save_click",
                save_batch_index=save_batches + 1,
                excel_rows=[match["item"]["row"] for match in voucher_batch],
                rows=len(voucher_batch),
                save_trigger=self.save_trigger,
            )
            with self.perf.span(
                "voucher_save_click",
                rows=len(voucher_batch),
                excel_rows=[match["item"]["row"] for match in voucher_batch],
                save_batch_index=save_batches + 1,
                save_trigger=self.save_trigger,
                hotkey_activate_policy=self.hotkey_activate_policy,
            ):
                if not self.trigger_voucher_save():
                    raise RuntimeError(
                        f"点击保存失败: Excel行{voucher_batch[0]['item']['row']}"
                    )
            self.record_transition(
                "voucher_save_clicked",
                from_state="voucher_open",
                to_state="voucher_open",
                save_batch_index=save_batches + 1,
                rows=len(voucher_batch),
            )

            self.run_state.set_stage(
                "voucher_save_verify",
                save_batch_index=save_batches + 1,
                excel_rows=[match["item"]["row"] for match in voucher_batch],
            )
            with self.perf.span(
                "voucher_save_verify",
                rows=len(voucher_batch),
                save_batch_index=save_batches + 1,
            ):
                verify_result = self.verify_voucher_batch_removed(
                    voucher_batch, before_count
                )
            self.record_event(
                "event_voucher_save_verified",
                save_batch_index=save_batches + 1,
                result=verify_result,
                rows=len(voucher_batch),
            )
            saved_matches.extend(voucher_batch)
            saved_rows = {match["item"]["row"] for match in voucher_batch}
            batch_status_updates = {
                match["item"]["row"]: self.generated_status for match in voucher_batch
            }
            pending_status_updates.update(batch_status_updates)
            self.run_state.event(
                "voucher_batch_saved",
                save_batch_index=save_batches + 1,
                excel_rows=[match["item"]["row"] for match in voucher_batch],
                pending_excel_status_updates=sorted(pending_status_updates),
            )
            if self.write_generated_status_each_save:
                with self.perf.span(
                    "excel_save_generated_status",
                    rows=len(voucher_batch),
                    save_batch_index=save_batches + 1,
                ):
                    self.run_state.set_stage(
                        "excel_save_generated_status",
                        save_batch_index=save_batches + 1,
                        excel_rows=[match["item"]["row"] for match in voucher_batch],
                    )
                    self.data_handler.save_jab_results(batch_status_updates)
                pending_status_updates = {}
            else:
                self.run_state.set_stage(
                    "excel_defer_generated_status",
                    pending_rows=sorted(pending_status_updates),
                )
            pending = [
                match for match in pending if match["item"]["row"] not in saved_rows
            ]
            pending_source = [
                match
                for match in pending_source
                if match["item"]["row"] not in saved_rows
            ]
            save_batches += 1
            if pending_source and verify_result in ("empty_window", "window_closed"):
                raise RuntimeError(
                    "制单窗口已空/关闭但仍有未保存 Excel 行，停止复核: "
                    f"remaining_excel_rows={[m['item']['row'] for m in pending_source]}"
                )
            if cached_voucher_matches is not None:
                cached_voucher_matches = self.advance_voucher_queue_cache(
                    cached_voucher_matches,
                    saved_rows,
                )
            if pending_source:
                with self.perf.span(
                    "save_batch_wait",
                    wait=self.wait_between_save_batches,
                    save_batch_index=save_batches,
                ):
                    if self.wait_between_save_batches > 0:
                        time.sleep(self.wait_between_save_batches)

        if pending_status_updates:
            self.run_state.set_stage(
                "excel_save_generated_status_bulk",
                rows=len(pending_status_updates),
                excel_rows=sorted(pending_status_updates),
            )
            with self.perf.span(
                "excel_save_generated_status_bulk",
                rows=len(pending_status_updates),
            ):
                self.data_handler.save_jab_results(pending_status_updates)

        return saved_matches, save_batches

    def trigger_voucher_save(self):
        if self.save_trigger == "jab_button":
            return self.jab.click_save(timeout=self.save_success_timeout)
        if self.save_trigger == "hotkey":
            if not self.jab.window_exists(
                self.voucher_window_title,
                class_name=self.voucher_window_class,
            ):
                log.warning("Ctrl+S 保存前未检测到制单窗口")
                return False
            if not self.prepare_hotkey_save_focus():
                return False
            self.jab.press_hotkey("ctrl", "s", wait=0)
            self.hotkey_save_attempts += 1
            return True
        raise ValueError(f"不支持的保存触发方式: {self.save_trigger!r}")

    def prepare_hotkey_save_focus(self):
        if self.hotkey_activate_policy == "always":
            return self.jab.activate_window_by_title(
                self.voucher_window_title,
                class_name=self.voucher_window_class,
                timeout=1,
            )

        if self.hotkey_activate_policy == "first":
            if self.hotkey_save_attempts == 0:
                return self.jab.activate_window_by_title(
                    self.voucher_window_title,
                    class_name=self.voucher_window_class,
                    timeout=1,
                )
            return True

        if self.hotkey_activate_policy == "foreground_guard":
            if self.jab.foreground_window_matches(
                self.voucher_window_title,
                class_name=self.voucher_window_class,
            ):
                return True
            foreground = self.jab.get_foreground_window_info()
            log.warning(
                "Ctrl+S 保存前前台窗口不是制单，停止避免误触发: "
                f"foreground={foreground}"
            )
            return False

        raise ValueError(f"不支持的 Ctrl+S 激活策略: {self.hotkey_activate_policy!r}")

    def should_use_voucher_queue_cache(self, matches):
        if not self.use_voucher_queue_cache:
            return False
        if self.save_strategy not in ("single", "safe_batch_by_pending_row"):
            return False
        if not matches:
            return False
        table_index = matches[0]["table_index"]
        table_rows = matches[0]["table_rows"]
        return all(
            match["table_index"] == table_index and match["table_rows"] == table_rows
            for match in matches
        )

    def advance_voucher_queue_cache(self, cached_matches, saved_rows):
        deleted_rows = [
            match["voucher_row"]
            for match in cached_matches
            if match["item"]["row"] in saved_rows
        ]
        advanced = []
        for match in cached_matches:
            if match["item"]["row"] in saved_rows:
                continue
            next_row = match["voucher_row"]
            next_row -= sum(1 for deleted_row in deleted_rows if next_row > deleted_row)
            advanced.append(
                {
                    **match,
                    "voucher_row": next_row,
                    "table_rows": match["table_rows"] - len(deleted_rows),
                }
            )
        if advanced and any(match["voucher_row"] < 0 for match in advanced):
            log.warning("制单表队列缓存行号异常，停止使用缓存")
            return None
        return advanced

    def read_voucher_tables(self, pending_count):
        with self.perf.span("voucher_table_read", pending=pending_count):
            tables = self.jab.read_window_table_cells(
                self.voucher_window_title,
                max_rows=500,
                max_cols=13,
            )
        return self.filter_voucher_tables(tables)

    def filter_voucher_tables(self, tables):
        voucher_tables = [
            table
            for table in tables
            if table["row_count"] > 0 and table["col_count"] == 13
        ]
        self.perf.event(
            "voucher_table_loaded",
            tables=len(voucher_tables),
            rows=sum(table["row_count"] for table in voucher_tables),
        )
        if not voucher_tables:
            raise RuntimeError("未找到制单窗口表格")
        return voucher_tables

    def match_voucher_table(self, matches, tables=None):
        if tables is None:
            voucher_tables = self.read_voucher_tables(len(matches))
        else:
            voucher_tables = self.filter_voucher_tables(tables)
        row_records = []
        for table in voucher_tables:
            for row in table["rows"]:
                amount = None
                for cell in row["cells"]:
                    amount = self.jab.normalize_amount(cell)
                    if amount is not None:
                        break
                row_text = "".join(
                    self.jab.normalize_text(cell) for cell in row["cells"]
                )
                if amount is None:
                    continue
                row_records.append(
                    {
                        "table": table,
                        "row": row,
                        "amount": amount,
                        "row_text": row_text,
                        "partner_key_text": self.normalize_partner_match_text(row_text),
                    }
                )

        index = defaultdict(list)
        partner_index = defaultdict(list)
        for record in row_records:
            for match in matches:
                partner_key = self.normalize_partner_match_text(
                    match["item"]["partner"]
                )
                if partner_key and partner_key in record["partner_key_text"]:
                    partner_index[match["item"]["row"]].append(
                        {
                            "table": record["table"],
                            "row": record["row"],
                            "amount": record["amount"],
                            "fallback_reason": "partner_normalized",
                        }
                    )
                if (
                    record["amount"] == self._as_decimal(match["item"]["amount"])
                    and partner_key
                    and partner_key in record["partner_key_text"]
                ):
                    index[match["item"]["row"]].append(
                        {
                            "table": record["table"],
                            "row": record["row"],
                            "amount": record["amount"],
                            "fallback_reason": "",
                        }
                    )

        voucher_matches = []
        issues = []
        assigned_rows = set()
        rate_group_matches = self.find_partner_rate_group_matches(matches, row_records)
        for ordinal, match in enumerate(matches):
            rows = index.get(match["item"]["row"], [])
            if len(rows) == 1:
                found = rows[0]
                self._append_voucher_match(voucher_matches, match, found, assigned_rows)
                continue

            partner_rows = partner_index.get(match["item"]["row"], [])
            if len(partner_rows) == 1:
                found = partner_rows[0]
                self._append_voucher_match(voucher_matches, match, found, assigned_rows)
                log.warning(
                    "制单表金额与 Excel 不一致，按唯一对手方匹配: "
                    f"Excel行{match['item']['row']} expected_amount={match['item']['amount']} "
                    f"voucher_amount={found['amount']} voucher_row={found['row']['row_index']}"
                )
                continue

            rate_group_match = rate_group_matches.get(match["item"]["row"])
            if rate_group_match:
                self._append_voucher_match(
                    voucher_matches,
                    match,
                    rate_group_match,
                    assigned_rows,
                )
                log.warning(
                    "制单表金额与 Excel 不一致，按归一化对手方+汇率一致性匹配: "
                    f"Excel行{match['item']['row']} expected_amount={match['item']['amount']} "
                    f"voucher_amount={rate_group_match['amount']} "
                    f"voucher_row={rate_group_match['row']['row_index']}"
                )
                continue

            fallback = self.find_voucher_order_fallback(
                match, ordinal, matches, row_records
            )
            if not fallback and self.voucher_order_fallback_mode == "same_count":
                fallback = self.find_voucher_same_count_order_fallback(
                    ordinal,
                    matches,
                    row_records,
                    assigned_rows,
                )
            if fallback:
                self._append_voucher_match(
                    voucher_matches, match, fallback, assigned_rows
                )
                log.warning(
                    "制单表金额不一致，按本批顺序+对手方匹配: "
                    f"Excel行{match['item']['row']} expected_amount={match['item']['amount']} "
                    f"voucher_amount={fallback['amount']} voucher_row={fallback['row']['row_index']}"
                )
            else:
                issues.append(
                    {
                        "item": match["item"],
                        "reason": "未找到" if not rows else f"重复{len(rows)}条",
                        "rows": [row["row"]["row_index"] for row in rows],
                    }
                )

        return voucher_matches, issues

    def _append_voucher_match(self, voucher_matches, match, found, assigned_rows):
        row_key = (
            found["table"]["table_index"],
            found["row"]["row_index"],
        )
        if row_key in assigned_rows:
            return
        assigned_rows.add(row_key)
        voucher_matches.append(
            {
                **match,
                "table_index": found["table"]["table_index"],
                "table_rows": found["table"]["row_count"],
                "voucher_row": found["row"]["row_index"],
                "voucher_cells": found["row"]["cells"],
            }
        )

    def find_voucher_order_fallback(self, match, ordinal, matches, row_records):
        """NC sometimes changes voucher amount during front generation.

        Only trust the fallback when the generated voucher table has the same
        row count as the current batch and the row at the same ordinal contains
        the expected partner name.
        """
        partner = self.jab.normalize_text(match["item"]["partner"])
        if not partner:
            return None

        candidates = []
        for record in row_records:
            table = record["table"]
            row = record["row"]
            if table["row_count"] != len(matches):
                continue
            if row["row_index"] != ordinal:
                continue
            if partner not in record["row_text"]:
                continue
            candidates.append(
                {
                    **record,
                    "fallback_reason": "order_partner",
                }
            )

        if len(candidates) == 1:
            return candidates[0]
        return None

    def find_partner_rate_group_matches(self, matches, row_records):
        by_partner = defaultdict(list)
        records_by_partner = defaultdict(list)
        for match in matches:
            partner_key = self.normalize_partner_match_text(match["item"]["partner"])
            if partner_key:
                by_partner[partner_key].append(match)
        for record in row_records:
            for partner_key in by_partner:
                if partner_key in record["partner_key_text"]:
                    records_by_partner[partner_key].append(record)

        result = {}
        for partner_key, partner_matches in by_partner.items():
            partner_records = records_by_partner.get(partner_key, [])
            if len(partner_matches) <= 1:
                continue
            if len(partner_records) != len(partner_matches):
                continue
            if len(partner_matches) > 7:
                continue
            assignment = self.choose_rate_consistent_assignment(
                partner_matches,
                partner_records,
            )
            if not assignment:
                continue
            for match, record in assignment:
                result[match["item"]["row"]] = {
                    "table": record["table"],
                    "row": record["row"],
                    "amount": record["amount"],
                    "fallback_reason": "partner_rate_group",
                }
        return result

    def choose_rate_consistent_assignment(self, matches, records):
        best = None
        best_score = None
        for permutation in itertools.permutations(records):
            rates = []
            valid = True
            for match, record in zip(matches, permutation):
                excel_amount = self._as_decimal(match["item"]["amount"])
                if not excel_amount:
                    valid = False
                    break
                rates.append(record["amount"] / excel_amount)
            if not valid:
                continue
            expected_rate = self.foreign_currency_rate
            if expected_rate:
                score = sum(abs(rate - expected_rate) / expected_rate for rate in rates)
            else:
                mean_rate = sum(rates) / len(rates)
                if not mean_rate:
                    continue
                score = sum(abs(rate - mean_rate) / mean_rate for rate in rates)
            if best_score is None or score < best_score:
                best_score = score
                best = list(zip(matches, permutation))

        if best is None or best_score is None:
            return None
        if best_score > self.foreign_currency_rate_tolerance:
            return None
        return best

    def normalize_partner_match_text(self, value):
        text = self.jab.normalize_text(value)
        return re.sub(r"[^0-9A-Z\u4e00-\u9fff]+", "", text.upper())

    def parse_optional_decimal(self, value):
        if value in (None, ""):
            return None
        return Decimal(str(value))

    def find_voucher_same_count_order_fallback(
        self,
        ordinal,
        matches,
        row_records,
        assigned_rows,
    ):
        if len(row_records) != len(matches):
            return None
        if ordinal >= len(row_records):
            return None
        record = row_records[ordinal]
        row_key = (
            record["table"]["table_index"],
            record["row"]["row_index"],
        )
        if row_key in assigned_rows:
            return None
        return {
            **record,
            "fallback_reason": "same_count_order",
        }

    def build_voucher_save_batches(self, matches):
        if self.save_strategy == "single":
            return [[match] for match in matches]
        if self.save_strategy == "safe_batch_by_pending_row":
            return self.build_safe_pending_row_batches(matches)
        if self.save_strategy == "bottom_up":
            return self.build_voucher_bottom_up_batches(matches)
        raise ValueError(f"不支持的保存策略: {self.save_strategy!r}")

    def build_safe_pending_row_batches(self, matches):
        batches = []
        current = []
        last_nc_row = None
        last_voucher_row = None

        for match in matches:
            nc_row = match.get("nc_row")
            voucher_row = match["voucher_row"]
            can_extend = (
                current
                and match["table_index"] == current[-1]["table_index"]
                and nc_row is not None
                and last_nc_row is not None
                and nc_row > last_nc_row
                and voucher_row > last_voucher_row
                and (self.max_batch_size <= 0 or len(current) < self.max_batch_size)
            )
            if current and not can_extend:
                batches.append(current)
                current = []

            current.append(match)
            last_nc_row = nc_row
            last_voucher_row = voucher_row

        if current:
            batches.append(current)

        self.perf.event(
            "safe_pending_row_batches",
            batch_sizes=[len(batch) for batch in batches],
            excel_rows=[[match["item"]["row"] for match in batch] for batch in batches],
            nc_rows=[[match.get("nc_row") for match in batch] for batch in batches],
            voucher_rows=[
                [match["voucher_row"] for match in batch] for batch in batches
            ],
        )
        return batches

    def build_voucher_bottom_up_batches(self, matches):
        batches = []
        current = []
        last_row = None

        for match in matches:
            row = match["voucher_row"]
            should_split = current and (
                row >= last_row
                or match["table_index"] != current[-1]["table_index"]
                or (self.max_batch_size > 0 and len(current) >= self.max_batch_size)
            )
            if should_split:
                batches.append(current)
                current = []
                last_row = None

            current.append(match)
            last_row = row

        if current:
            batches.append(current)
        return batches

    def get_voucher_selection_rows(self, voucher_batch):
        return [match["voucher_row"] for match in voucher_batch]

    def verify_voucher_batch_removed(self, voucher_batch, before_count):
        expected_removed = len(voucher_batch)
        deadline = time.time() + self.voucher_record_timeout
        target_rows = {match["item"]["row"] for match in voucher_batch}

        while time.time() < deadline:
            check_abort()
            window_exists = self.jab.window_exists(
                self.voucher_window_title,
                class_name=self.voucher_window_class,
            )
            with self.perf.span(
                "voucher_save_verify_counts",
                rows=expected_removed,
                before=before_count,
            ):
                table_counts = self.jab.read_window_table_counts(
                    self.voucher_window_title
                )
            counts = [
                table["row_count"]
                for table in table_counts
                if table["row_count"] > 0 and table["col_count"] == 13
            ]
            if not counts:
                if window_exists:
                    log.info(
                        "制单窗口仍存在但制单表为空，转待生成表复核: "
                        f"excel_rows={sorted(target_rows)} before={before_count} "
                        f"expected_removed={expected_removed}"
                    )
                    return "empty_window"
                log.info(
                    f"制单窗口已关闭，转待生成表复核: excel_rows={sorted(target_rows)}"
                )
                return "window_closed"

            min_count = min(counts) if counts else 0
            if min_count <= before_count - expected_removed:
                log.info(
                    "制单批次保存行数验证通过: "
                    f"excel_rows={sorted(target_rows)} before={before_count} after={min_count}"
                )
                return True

            tables = self.jab.read_window_table_cells(
                self.voucher_window_title,
                max_rows=500,
                max_cols=13,
            )
            voucher_tables = [
                table
                for table in tables
                if table["row_count"] > 0 and table["col_count"] == 13
            ]
            remaining_matches, issues = self.match_voucher_table(
                voucher_batch,
                tables=voucher_tables,
            )
            remaining_rows = {match["item"]["row"] for match in remaining_matches}

            if not remaining_rows and min_count <= before_count - expected_removed:
                log.info(
                    "制单批次保存验证通过: "
                    f"excel_rows={sorted(target_rows)} before={before_count} after={min_count}"
                )
                return True

            time.sleep(0.3)

        raise RuntimeError(
            "制单批次保存后仍可匹配到记录或行数未减少: "
            f"excel_rows={sorted(target_rows)} before={before_count}"
        )

    def close_and_verify_pending_removed(self, voucher_batch):
        close_cfg = self.batch_cfg.get("close_voucher_window", {})
        if self.jab.window_exists(
            self.voucher_window_title,
            class_name=self.voucher_window_class,
        ):
            self.record_event(
                "event_voucher_window_close_start",
                excel_rows=[match["item"]["row"] for match in voucher_batch],
            )
            self.jab.close_window_by_title(
                close_cfg.get("title", self.voucher_window_title),
                class_name=close_cfg.get("class_name", self.voucher_window_class),
                wait=float(close_cfg.get("wait", 0.5)),
            )
            self.record_transition(
                "voucher_window_closed",
                from_state="voucher_open",
                to_state="pending",
                excel_rows=[match["item"]["row"] for match in voucher_batch],
            )

        self.record_event(
            "event_pending_refresh",
            key="f5",
            wait=float(self.batch_cfg.get("pending_refresh_wait", 1.0)),
        )
        self.jab.press_key(
            "f5", wait=float(self.batch_cfg.get("pending_refresh_wait", 1.0))
        )
        snapshot = self.jab.read_table_snapshot()
        index = defaultdict(list)
        for row in snapshot:
            if row["amount"] is None or not row["partner"]:
                continue
            index[(row["amount"], row["partner"])].append(row)

        still_present = []
        for match in voucher_batch:
            item = match["item"]
            key = (
                self._as_decimal(item["amount"]),
                self.jab.normalize_text(item["partner"]),
            )
            if index.get(key):
                still_present.append(item["row"])

        if still_present:
            raise RuntimeError(f"待生成表刷新后仍存在本批记录: Excel行{still_present}")

        log.info(
            "待生成表复核通过，本批记录已消失: "
            f"excel_rows={[match['item']['row'] for match in voucher_batch]}"
        )
        return True

    def verify_current_voucher_record(self, match):
        item = match["item"]
        found = self.jab.wait_for_record_visible(
            item["amount"],
            item["partner"],
            timeout=self.voucher_record_timeout,
            window_title=self.voucher_window_title,
        )
        if not found:
            raise RuntimeError(
                "制单界面未找到当前记录，停止以避免错保存: "
                f"Excel行{item['row']} amount={item['amount']} partner={item['partner']}"
            )
        log.debug(
            "制单当前记录验证通过: "
            f"Excel行{item['row']} table={found['table_index']} row={found['row_index']} "
            f"selected={found['selected']}"
        )
        return found

    def wait_for_next_voucher_record(self, next_match):
        item = next_match["item"]
        found = self.jab.wait_for_record_visible(
            item["amount"],
            item["partner"],
            timeout=self.voucher_record_timeout,
            window_title=self.voucher_window_title,
        )
        if not found:
            raise RuntimeError(
                "保存后未检测到制单界面推进到下一条，停止以避免误标记: "
                f"下一Excel行{item['row']} amount={item['amount']} partner={item['partner']}"
            )
        log.info(
            "保存后已推进到下一条: "
            f"Excel行{item['row']} table={found['table_index']} row={found['row_index']} "
            f"selected={found['selected']}"
        )
        return found

    def format_issue_updates(self, issues, prefix=""):
        updates = {}
        for issue in issues:
            reason = issue["reason"]
            if issue.get("rows"):
                reason = f"{reason}-NC行{','.join(str(r) for r in issue['rows'][:5])}"
            updates[issue["item"]["row"]] = f"{prefix}{reason}"
        return updates

    def _append_match_or_issue(self, matches, issues, item, rows):
        if len(rows) == 1:
            matches.append(
                {
                    "item": item,
                    "nc_row": rows[0]["row_index"],
                    "row_data": rows[0],
                }
            )
        else:
            issues.append(
                {
                    "item": item,
                    "reason": "未找到" if not rows else f"重复{len(rows)}条",
                    "rows": [row["row_index"] for row in rows],
                }
            )

    def _find_contains(self, snapshot, key):
        amount, partner = key
        rows = []
        for row in snapshot:
            if row["amount"] == amount and partner in row["partner"]:
                rows.append(row)
        return rows

    def _as_decimal(self, value):
        if isinstance(value, Decimal):
            return value
        return self.jab.normalize_amount(value)

    def _log_plan(self, items, parse_errors, matches, issues, batches):
        log.info(
            "JAB dry-run: "
            f"items={len(items)} parse_errors={len(parse_errors)} "
            f"matched={len(matches)} issues={len(issues)} batches={len(batches)}"
        )
        for issue in parse_errors:
            log.warning(f"Excel行{issue['row']} 格式错误: {issue.get('parse_error')}")
        for issue in issues:
            log.warning(
                f"Excel行{issue['item']['row']} {issue['reason']}: "
                f"amount={issue['item']['amount']} partner={issue['item']['partner']} "
                f"nc_rows={issue.get('rows', [])}"
            )
        for index, batch in enumerate(batches, start=1):
            log.info(
                f"批次{index}: size={len(batch)} "
                f"excel_rows={[m['item']['row'] for m in batch]} "
                f"nc_rows={[m['nc_row'] for m in batch]}"
            )
