import time
from collections import defaultdict
from datetime import date
from decimal import Decimal

from core.data_handler import DataHandler
from core.jab_operator import JABOperator
from core.logger import log
from core.nc_backfill_workflow import NCBackfillWorkflow
from core.nc_state import NCStateDetector, normalize_generated_voucher
from core.nc_switch_generated_workflow import NCSwitchGeneratedWorkflow
from core.nc_voucher_workflow import NCVoucherWorkflow
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
        self.backfill_workflow = NCBackfillWorkflow(self)
        self.switch_generated_workflow = NCSwitchGeneratedWorkflow(self)
        self.voucher_workflow = NCVoucherWorkflow(self)

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
        return self.backfill_workflow.backfill_generated_vouchers(
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        )

    def switch_to_generated_list(self, *args, **kwargs):
        return self.switch_generated_workflow.switch_to_generated_list(*args, **kwargs)

    def run_switch_generated_steps(self, *args, **kwargs):
        return self.switch_generated_workflow.run_switch_generated_steps(
            *args, **kwargs
        )

    def get_switch_step_name(self, *args, **kwargs):
        return self.switch_generated_workflow.get_switch_step_name(*args, **kwargs)

    def run_query_window_step(self, *args, **kwargs):
        return self.switch_generated_workflow.run_query_window_step(*args, **kwargs)

    def find_query_window(self, *args, **kwargs):
        return self.switch_generated_workflow.find_query_window(*args, **kwargs)

    def open_query_with_jab_action(self, *args, **kwargs):
        return self.switch_generated_workflow.open_query_with_jab_action(
            *args, **kwargs
        )

    def _trigger_query_action_path(self, *args, **kwargs):
        return self.switch_generated_workflow._trigger_query_action_path(
            *args, **kwargs
        )

    def _do_query_action_path(self, *args, **kwargs):
        return self.switch_generated_workflow._do_query_action_path(*args, **kwargs)

    def run_jab_action_subprocess(self, *args, **kwargs):
        return self.switch_generated_workflow.run_jab_action_subprocess(*args, **kwargs)

    def open_query_with_hotkey(self, *args, **kwargs):
        return self.switch_generated_workflow.open_query_with_hotkey(*args, **kwargs)

    def get_nc_workflow_state(self, *args, **kwargs):
        return self.switch_generated_workflow.get_nc_workflow_state(*args, **kwargs)

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
            self.require_page_state(
                "voucher_open",
                pending,
                command="resume-voucher",
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

    def save_current_voucher_matches(self, *args, **kwargs):
        return self.voucher_workflow.save_current_voucher_matches(*args, **kwargs)

    def trigger_voucher_save(self, *args, **kwargs):
        return self.voucher_workflow.trigger_voucher_save(*args, **kwargs)

    def prepare_hotkey_save_focus(self, *args, **kwargs):
        return self.voucher_workflow.prepare_hotkey_save_focus(*args, **kwargs)

    def voucher_verify_result_state(self, *args, **kwargs):
        return self.voucher_workflow.voucher_verify_result_state(*args, **kwargs)

    def should_use_voucher_queue_cache(self, *args, **kwargs):
        return self.voucher_workflow.should_use_voucher_queue_cache(*args, **kwargs)

    def advance_voucher_queue_cache(self, *args, **kwargs):
        return self.voucher_workflow.advance_voucher_queue_cache(*args, **kwargs)

    def read_voucher_tables(self, *args, **kwargs):
        return self.voucher_workflow.read_voucher_tables(*args, **kwargs)

    def filter_voucher_tables(self, *args, **kwargs):
        return self.voucher_workflow.filter_voucher_tables(*args, **kwargs)

    def match_voucher_table(self, *args, **kwargs):
        return self.voucher_workflow.match_voucher_table(*args, **kwargs)

    def _append_voucher_match(self, *args, **kwargs):
        return self.voucher_workflow._append_voucher_match(*args, **kwargs)

    def find_voucher_order_fallback(self, *args, **kwargs):
        return self.voucher_workflow.find_voucher_order_fallback(*args, **kwargs)

    def find_partner_rate_group_matches(self, *args, **kwargs):
        return self.voucher_workflow.find_partner_rate_group_matches(*args, **kwargs)

    def choose_rate_consistent_assignment(self, *args, **kwargs):
        return self.voucher_workflow.choose_rate_consistent_assignment(*args, **kwargs)

    def normalize_partner_match_text(self, *args, **kwargs):
        return self.voucher_workflow.normalize_partner_match_text(*args, **kwargs)

    def find_voucher_same_count_order_fallback(self, *args, **kwargs):
        return self.voucher_workflow.find_voucher_same_count_order_fallback(
            *args, **kwargs
        )

    def build_voucher_save_batches(self, *args, **kwargs):
        return self.voucher_workflow.build_voucher_save_batches(*args, **kwargs)

    def build_safe_pending_row_batches(self, *args, **kwargs):
        return self.voucher_workflow.build_safe_pending_row_batches(*args, **kwargs)

    def build_voucher_bottom_up_batches(self, *args, **kwargs):
        return self.voucher_workflow.build_voucher_bottom_up_batches(*args, **kwargs)

    def get_voucher_selection_rows(self, *args, **kwargs):
        return self.voucher_workflow.get_voucher_selection_rows(*args, **kwargs)

    def verify_voucher_batch_removed(self, *args, **kwargs):
        return self.voucher_workflow.verify_voucher_batch_removed(*args, **kwargs)

    def close_and_verify_pending_removed(self, *args, **kwargs):
        return self.voucher_workflow.close_and_verify_pending_removed(*args, **kwargs)

    def verify_current_voucher_record(self, *args, **kwargs):
        return self.voucher_workflow.verify_current_voucher_record(*args, **kwargs)

    def wait_for_next_voucher_record(self, *args, **kwargs):
        return self.voucher_workflow.wait_for_next_voucher_record(*args, **kwargs)

    def parse_optional_decimal(self, value):
        if value in (None, ""):
            return None
        return Decimal(str(value))

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
