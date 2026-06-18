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
        self.duplicate_match_policy = self.batch_cfg.get(
            "duplicate_match_policy", "stop"
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

    def generate_and_collect_saved(self, *args, **kwargs):
        return self.pending_workflow.generate_and_collect_saved(*args, **kwargs)

    def generate_and_backfill(self, limit=None, max_batches=None, start_row=None, end_row=None):
        result = self.generate_and_collect_saved(
            limit=limit,
            max_batches=max_batches,
            start_row=start_row,
            end_row=end_row,
        )
        excel_rows = result.get("excel_rows") or []
        if not excel_rows:
            return {"saved": 0, "batches": 0, "updates": {}}

        updates = self.backfill_generated_vouchers(
            limit=None,
            start_row=min(excel_rows),
            end_row=max(excel_rows),
            auto_switch=True,
            require_generated_status=False,
            assume_pending_ready=True,
        )
        return {
            "saved": int(result.get("saved", 0)),
            "batches": int(result.get("batches", 0)),
            "updates": updates,
            "excel_rows": excel_rows,
        }

    def backfill_generated_vouchers(
        self,
        limit=None,
        start_row=None,
        end_row=None,
        auto_switch=True,
        require_generated_status=True,
        assume_pending_ready=False,
    ):
        return self.backfill_workflow.backfill_generated_vouchers(
            limit=limit,
            start_row=start_row,
            end_row=end_row,
            auto_switch=auto_switch,
            require_generated_status=require_generated_status,
            assume_pending_ready=assume_pending_ready,
        )

    def switch_to_generated_list(self, *args, **kwargs):
        return self.switch_generated_workflow.switch_to_generated_list(*args, **kwargs)

    def require_page_state(self, expected, items=None, command=""):
        return self.state_detector.require_page_state(expected, items, command)

    def wait_for_page_state(self, expected, items=None, command="", timeout=None):
        return self.state_detector.wait_for_page_state(expected, items, command, timeout)

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

    def resume_current_voucher_window(self, *args, **kwargs):
        return self.pending_workflow.resume_current_voucher_window(*args, **kwargs)

    def parse_optional_decimal(self, value):
        if value in (None, ""):
            return None
        return Decimal(str(value))
