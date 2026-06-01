from datetime import date
from decimal import Decimal

from core.data_handler import DataHandler
from core.jab_operator import JABOperator
from core.nc_backfill_workflow import NCBackfillWorkflow
from core.nc_pending_workflow import NCPendingWorkflow
from core.nc_state import NCStateDetector, normalize_generated_voucher
from core.nc_switch_generated_workflow import NCSwitchGeneratedWorkflow
from core.nc_table_matcher import NCTableMatcher
from core.nc_voucher_workflow import NCVoucherWorkflow
from core.perf import PerfRecorder
from core.run_state import RunStateRecorder


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
        self.pending_workflow = NCPendingWorkflow(self)
        self.backfill_workflow = NCBackfillWorkflow(self)
        self.switch_generated_workflow = NCSwitchGeneratedWorkflow(self)
        self.table_matcher = NCTableMatcher(self)
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

    def load_pending_items(self, *args, **kwargs):
        return self.pending_workflow.load_pending_items(*args, **kwargs)

    def dry_run(self, *args, **kwargs):
        return self.pending_workflow.dry_run(*args, **kwargs)

    def generate_and_save(self, *args, **kwargs):
        return self.pending_workflow.generate_and_save(*args, **kwargs)

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

    def match_current_table(self, *args, **kwargs):
        return self.table_matcher.match_current_table(*args, **kwargs)

    def filter_generated_date_rows(self, *args, **kwargs):
        return self.table_matcher.filter_generated_date_rows(*args, **kwargs)

    def build_increasing_batches(self, *args, **kwargs):
        return self.table_matcher.build_increasing_batches(*args, **kwargs)

    def process_full_selection(self, *args, **kwargs):
        return self.pending_workflow.process_full_selection(*args, **kwargs)

    def resume_current_voucher_window(self, *args, **kwargs):
        return self.pending_workflow.resume_current_voucher_window(*args, **kwargs)

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

    def format_issue_updates(self, *args, **kwargs):
        return self.pending_workflow.format_issue_updates(*args, **kwargs)

    def _log_plan(self, *args, **kwargs):
        return self.pending_workflow._log_plan(*args, **kwargs)
