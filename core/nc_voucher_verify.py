# 职责：制单保存后的验证复核——批次移除校验、待生成移除复核
# 不做什么：不直接持有 JAB/processor 状态(经 NCVoucherWorkflow 的 self 代理到 processor)
# 允许依赖层：core.errors/models/logger/utils;由 NCVoucherWorkflow 多继承组合
# 谁不应该 import：其它 nc_*_workflow 不应 import(import-linter 独立性约束)

import time
from collections import defaultdict

from core.errors import (
    ContractViolation,
)
from core.logger import log
from core.models import VoucherSaveMatch
from core.utils import check_abort


class NCVoucherVerifyMixin:
    def verify_voucher_batch_removed(
        self, voucher_batch: list[VoucherSaveMatch], before_count
    ):
        expected_removed = len(voucher_batch)
        deadline = time.time() + self.voucher_record_timeout
        target_rows = {match.item.row for match in voucher_batch}

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
            remaining_rows = {match.item.row for match in remaining_matches}

            if not remaining_rows and min_count <= before_count - expected_removed:
                log.info(
                    "制单批次保存验证通过: "
                    f"excel_rows={sorted(target_rows)} before={before_count} after={min_count}"
                )
                return True

            time.sleep(0.3)

        raise ContractViolation(
            "制单批次保存后仍可匹配到记录或行数未减少: "
            f"excel_rows={sorted(target_rows)} before={before_count}"
        )

    def close_and_verify_pending_removed(self, voucher_batch: list[VoucherSaveMatch]):
        close_cfg = self.batch_cfg.get("close_voucher_window", {})
        if self.jab.window_exists(
            self.voucher_window_title,
            class_name=self.voucher_window_class,
        ):
            self.record_event(
                "event_voucher_window_close_start",
                excel_rows=[match.item.row for match in voucher_batch],
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
                excel_rows=[match.item.row for match in voucher_batch],
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
            item = match.item
            key = (
                self.table_matcher._as_decimal(item.amount),
                self.jab.normalize_text(item.partner),
            )
            if index.get(key):
                still_present.append(item.row)

        if still_present:
            raise ContractViolation(
                f"待生成表刷新后仍存在本批记录: Excel行{still_present}"
            )

        log.info(
            "待生成表复核通过，本批记录已消失: "
            f"excel_rows={[match.item.row for match in voucher_batch]}"
        )
        return True
