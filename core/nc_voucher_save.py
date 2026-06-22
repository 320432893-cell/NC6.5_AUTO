# 职责：凭证制单保存编排——选行/触发保存/批次构建/队列缓存推进
# 不做什么：不直接持有 JAB/processor 状态(经 NCVoucherWorkflow 的 self 代理到 processor)
# 允许依赖层：core.errors/models/logger/utils;由 NCVoucherWorkflow 多继承组合
# 谁不应该 import：其它 nc_*_workflow 不应 import(import-linter 独立性约束)

import time
from dataclasses import replace

from core.errors import (
    ContractViolation,
    JABActionError,
    TableMatchError,
)
from core.logger import log
from core.models import VoucherSaveMatch
from core.utils import check_abort


class NCVoucherSaveMixin:
    def save_current_voucher_matches(
        self, voucher_matches: list[VoucherSaveMatch]
    ) -> tuple[list[VoucherSaveMatch], int]:
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
                        f"Excel行{issue.item.row} {issue.reason}" for issue in issues
                    )
                    raise TableMatchError(f"刷新制单表匹配失败: {detail}")
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
            before_count = voucher_batch[0].table_rows
            row_indexes = [match.voucher_row for match in voucher_batch]
            selection_row_indexes = self.get_voucher_selection_rows(voucher_batch)
            self.run_state.set_stage(
                "voucher_rows_select",
                save_batch_index=save_batches + 1,
                excel_rows=[match.item.row for match in voucher_batch],
                voucher_rows=row_indexes,
                select_rows=selection_row_indexes,
                strategy=self.save_strategy,
            )
            log.info(
                "保存制单批次: "
                f"size={len(voucher_batch)} excel_rows={[m.item.row for m in voucher_batch]} "
                f"voucher_rows={row_indexes} select_rows={selection_row_indexes} "
                f"strategy={self.save_strategy} before_count={before_count}"
            )

            with self.perf.span(
                "voucher_rows_select",
                rows=len(selection_row_indexes),
                table_index=voucher_batch[0].table_index,
                logical_rows=row_indexes,
                select_rows=selection_row_indexes,
                save_batch_index=save_batches + 1,
            ):
                if not self.jab.select_visible_table_rows(
                    voucher_batch[0].table_index,
                    selection_row_indexes,
                    window_title=self.voucher_window_title,
                ):
                    raise JABActionError(f"选中制单表行失败: {selection_row_indexes}")

            self.run_state.set_stage(
                "voucher_save_click",
                save_batch_index=save_batches + 1,
                excel_rows=[match.item.row for match in voucher_batch],
            )
            self.require_page_state(
                "voucher_open",
                voucher_batch,
                command="voucher-save",
            )
            self.record_event(
                "event_voucher_save_click",
                save_batch_index=save_batches + 1,
                excel_rows=[match.item.row for match in voucher_batch],
                rows=len(voucher_batch),
                save_trigger=self.save_trigger,
            )
            with self.perf.span(
                "voucher_save_click",
                rows=len(voucher_batch),
                excel_rows=[match.item.row for match in voucher_batch],
                save_batch_index=save_batches + 1,
                save_trigger=self.save_trigger,
                hotkey_activate_policy=self.hotkey_activate_policy,
            ):
                if not self.trigger_voucher_save():
                    raise JABActionError(
                        f"点击保存失败: Excel行{voucher_batch[0].item.row}"
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
                excel_rows=[match.item.row for match in voucher_batch],
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
            self.record_transition(
                "voucher_save_verified",
                from_state="voucher_open",
                to_state=self.voucher_verify_result_state(verify_result),
                save_batch_index=save_batches + 1,
                result=verify_result,
                rows=len(voucher_batch),
            )
            saved_matches.extend(voucher_batch)
            # 累计已落库行数,供中途异常时上报真实"已存 N 单"(而非误报 0)
            self.run_state.add_count("voucher_saved", len(voucher_batch))
            saved_rows = {match.item.row for match in voucher_batch}
            batch_status_updates = {
                match.item.row: self.generated_status for match in voucher_batch
            }
            pending_status_updates.update(batch_status_updates)
            self.run_state.event(
                "voucher_batch_saved",
                save_batch_index=save_batches + 1,
                excel_rows=[match.item.row for match in voucher_batch],
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
                        excel_rows=[match.item.row for match in voucher_batch],
                    )
                    self.data_handler.save_jab_results(batch_status_updates)
                pending_status_updates = {}
            else:
                self.run_state.set_stage(
                    "excel_defer_generated_status",
                    pending_rows=sorted(pending_status_updates),
                )
            pending = [match for match in pending if match.item.row not in saved_rows]
            pending_source = [
                match for match in pending_source if match.item.row not in saved_rows
            ]
            save_batches += 1
            if pending_source and verify_result in ("empty_window", "window_closed"):
                raise ContractViolation(
                    "制单窗口已空/关闭但仍有未保存 Excel 行，停止复核: "
                    f"remaining_excel_rows={[m.item.row for m in pending_source]}"
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

    def voucher_verify_result_state(self, verify_result):
        if verify_result is True:
            return "voucher_open"
        if verify_result == "empty_window":
            return "voucher_open_empty"
        if verify_result == "window_closed":
            return "pending"
        return "error"

    def should_use_voucher_queue_cache(self, matches):
        if not self.use_voucher_queue_cache:
            return False
        if self.save_strategy not in ("single", "safe_batch_by_pending_row"):
            return False
        if not matches:
            return False
        table_index = matches[0].table_index
        table_rows = matches[0].table_rows
        return all(
            match.table_index == table_index and match.table_rows == table_rows
            for match in matches
        )

    def advance_voucher_queue_cache(self, cached_matches, saved_rows):
        deleted_rows = [
            match.voucher_row
            for match in cached_matches
            if match.item.row in saved_rows
        ]
        advanced = []
        for match in cached_matches:
            if match.item.row in saved_rows:
                continue
            next_row = match.voucher_row
            next_row -= sum(1 for deleted_row in deleted_rows if next_row > deleted_row)
            advanced.append(
                replace(
                    match,
                    voucher_row=next_row,
                    table_rows=match.table_rows - len(deleted_rows),
                )
            )
        if advanced and any(match.voucher_row < 0 for match in advanced):
            log.warning("制单表队列缓存行号异常，停止使用缓存")
            return None
        return advanced

    def build_voucher_save_batches(self, matches: list[VoucherSaveMatch]):
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
            nc_row = match.nc_row
            voucher_row = match.voucher_row
            can_extend = (
                current
                and match.table_index == current[-1].table_index
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
            excel_rows=[[match.item.row for match in batch] for batch in batches],
            nc_rows=[[match.nc_row for match in batch] for batch in batches],
            voucher_rows=[[match.voucher_row for match in batch] for batch in batches],
        )
        return batches

    def build_voucher_bottom_up_batches(self, matches):
        batches = []
        current = []
        last_row = None

        for match in matches:
            row = match.voucher_row
            should_split = current and (
                row >= last_row
                or match.table_index != current[-1].table_index
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

    def get_voucher_selection_rows(self, voucher_batch: list[VoucherSaveMatch]):
        return [match.voucher_row for match in voucher_batch]
