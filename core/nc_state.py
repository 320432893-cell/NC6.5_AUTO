import time
from dataclasses import dataclass

from core.errors import WorkflowStateError
from core.logger import log
from core.nc_page_probe import NCPageProbe


@dataclass
class NCPageState:
    name: str
    reason: str
    table: dict | None = None
    match_ratio: float | None = None


class NCStateDetector:
    def __init__(
        self,
        jab,
        batch_cfg,
        generated_date_value,
        generated_date_col,
        voucher_col,
        generated_voucher_max,
        record_event,
        record_transition,
    ):
        self.jab = jab
        self.batch_cfg = batch_cfg
        self.generated_date_value = generated_date_value
        self.generated_date_col = generated_date_col
        self.voucher_col = voucher_col
        self.generated_voucher_max = generated_voucher_max
        self.record_event = record_event
        self.record_transition = record_transition
        self.voucher_window_title = batch_cfg.get("voucher_window_title", "制单")
        self.voucher_window_class = batch_cfg.get(
            "voucher_window_class", "SunAwtDialog"
        )
        self.state_wait_timeout = float(batch_cfg.get("state_wait_timeout", 2.0))
        self.state_wait_interval = float(batch_cfg.get("state_wait_interval", 0.2))
        self.probe = NCPageProbe(jab, batch_cfg)

    def require_page_state(self, expected, items=None, command=""):
        deadline = time.time() + self.state_wait_timeout
        last_state = None
        while True:
            last_state = self.detect_page_state(items=items)
            self.record_event(
                "nc_page_state",
                command=command,
                expected=expected,
                actual=last_state.name,
                reason=last_state.reason,
                match_ratio=last_state.match_ratio,
            )
            if last_state.name == expected:
                self.record_transition(
                    "state_guard_passed",
                    to_state=last_state.name,
                    command=command,
                    expected=expected,
                    reason=last_state.reason,
                )
                log.info(
                    f"NC 页面状态确认: expected={expected} reason={last_state.reason}"
                )
                return last_state
            if last_state.name != "loading" or time.time() >= deadline:
                break
            time.sleep(self.state_wait_interval)

        raise WorkflowStateError(
            f"NC 页面状态不正确: command={command} expected={expected} "
            f"actual={last_state.name} reason={last_state.reason}"
        )

    def detect_page_state(self, items=None):
        voucher_state = self.detect_voucher_window_state()
        if voucher_state.name == "voucher_open":
            return voucher_state

        open_query = self.batch_cfg.get("open_query", {})
        if self.jab.window_exists(
            open_query.get("dialog_title", "查询"),
            class_name=open_query.get("dialog_class", "SunAwtDialog"),
        ):
            return NCPageState("query_open", "查询子窗口打开，父页面被阻塞")

        controls = self.probe.collect_named_controls(
            ("单据生成", "查询", "生成", "前台生成", "正式单据")
        )
        has_parent = any(control["name"] == "单据生成" for control in controls)
        visible_names = {
            control["name"] for control in controls if control.get("showing")
        }
        has_buttons = {"查询", "生成"}.issubset(visible_names)
        has_formal = "正式单据" in visible_names

        tables = self.probe.read_page_table_signatures(
            self.generated_date_col,
            self.voucher_col,
            self.jab.amount_col,
            self.jab.partner_col,
        )
        main = choose_main_signature_table(tables)

        if not has_parent or not main:
            if looks_loading(controls, tables):
                return NCPageState(
                    "loading",
                    f"父页面/主表暂不完整 parent={has_parent} tables={len(tables)}",
                )
            return NCPageState(
                "error",
                f"缺少父页面或主表 parent={has_parent} tables={len(tables)}",
            )

        if not has_buttons:
            return NCPageState(
                "error",
                f"按钮布局不完整 visible={sorted(visible_names)}",
                table=main,
            )

        if self.is_generated_signature(main, require_formal=has_formal):
            return NCPageState("generated", "父页+按钮+已生成表特征", table=main)

        ratio = self.table_match_ratio(main.get("rows", []), items or [])
        if is_pending_signature(main, visible_names):
            return NCPageState(
                "pending",
                f"父页+按钮+待生成表特征 sample_match_ratio={ratio:.3f}",
                table=main,
                match_ratio=ratio,
            )

        if looks_loading(controls, tables):
            return NCPageState("loading", "父页面存在但表格特征暂不完整", table=main)

        return NCPageState(
            "error",
            (
                f"未知页面 rows={main.get('row_count')} cols={main.get('col_count')} "
                f"match_ratio={ratio:.3f}"
            ),
            table=main,
            match_ratio=ratio,
        )

    def detect_voucher_window_state(self):
        tables = [
            table
            for table in self.jab.read_window_table_cells(
                self.voucher_window_title,
                max_rows=5,
                max_cols=13,
            )
            if table.get("window_class") == self.voucher_window_class
        ]
        voucher_tables = [
            table
            for table in tables
            if table.get("row_count", 0) > 0 and table.get("col_count") == 13
        ]
        buttons = self.probe.collect_named_controls(
            ("修改", "保存"),
            window_title=self.voucher_window_title,
            window_class=self.voucher_window_class,
        )
        visible_buttons = {button["name"] for button in buttons if button["showing"]}
        button_tokens = {
            button["token"]
            for button in self.probe.collect_visible_buttons_by_desc_tokens(
                ("Ctrl+E", "Ctrl+S"),
                window_title=self.voucher_window_title,
                window_class=self.voucher_window_class,
            )
        }
        if voucher_tables:
            rows = sum(table["row_count"] for table in voucher_tables)
            return NCPageState(
                "voucher_open",
                "制单子窗口打开，父页面被阻塞 "
                f"rows={rows} buttons={sorted(visible_buttons)} tokens={sorted(button_tokens)}",
            )
        if not tables and not visible_buttons:
            return NCPageState("not_voucher", "未检测到制单子窗口")
        return NCPageState(
            "error",
            (
                "制单窗口特征不完整: "
                f"tables={[(t.get('row_count'), t.get('col_count')) for t in tables]} "
                f"buttons={sorted(visible_buttons)} tokens={sorted(button_tokens)}"
            ),
        )

    def is_generated_signature(self, table, require_formal=True):
        if table.get("col_count") != 23:
            return False
        if require_formal is False:
            return False
        target = str(self.generated_date_value).strip()
        date_values = [str(value).strip() for value in table.get("date_values", [])]
        if target and not date_values:
            return False
        if target and any(value != target for value in date_values[:5]):
            return False
        vouchers = table.get("voucher_values", [])
        return any(
            normalize_generated_voucher(value, self.generated_voucher_max)
            for value in vouchers
        )

    def table_match_ratio(self, rows, items):
        parsed = [item for item in items if not item.get("parse_error")]
        if not parsed:
            return 0.0
        index = {
            (row.get("amount"), row.get("partner"))
            for row in rows
            if row.get("amount") is not None and row.get("partner")
        }
        if not index:
            return 0.0
        matched = 0
        for item in parsed:
            key = (
                self._as_decimal(item["amount"]),
                self.jab.normalize_text(item["partner"]),
            )
            if key in index:
                matched += 1
        return matched / len(parsed)

    def _as_decimal(self, value):
        if value is None:
            return None
        return self.jab.normalize_amount(str(value))


def normalize_generated_voucher(raw_voucher, generated_voucher_max):
    import re

    text = str(raw_voucher or "").strip()
    match = re.search(r"\d+", text)
    if not match:
        return None

    value = int(match.group(0))
    if value <= 0 or value > generated_voucher_max:
        return None
    return value


def choose_main_signature_table(tables):
    candidates = [table for table in tables if table.get("col_count", 0) > 1]
    if not candidates:
        return None
    return max(
        candidates,
        key=lambda table: table.get("row_count", 0) * table.get("col_count", 0),
    )


def is_pending_signature(table, visible_names):
    if table.get("col_count") != 25:
        return False
    if table.get("col_count") == 23 and table.get("voucher_values"):
        return False
    if "前台生成" not in visible_names:
        log.info("未枚举到前台生成按钮，按父页按钮和表格列数判定待生成页")
    return not table.get("voucher_values")


def looks_loading(controls, tables):
    if not controls and not tables:
        return True
    if controls and not tables:
        return True
    main = choose_main_signature_table(tables)
    return bool(main and main.get("row_count", 0) == 0)
