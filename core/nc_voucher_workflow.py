import time
import re
import itertools
from collections import defaultdict

from core.logger import log
from core.utils import check_abort


class NCVoucherWorkflow:
    def __init__(self, processor):
        super().__setattr__("processor", processor)

    def __getattr__(self, name):
        return getattr(self.processor, name)

    def __setattr__(self, name, value):
        if name == "processor":
            super().__setattr__(name, value)
            return
        setattr(self.processor, name, value)

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
            self.record_transition(
                "voucher_save_verified",
                from_state="voucher_open",
                to_state=self.voucher_verify_result_state(verify_result),
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
