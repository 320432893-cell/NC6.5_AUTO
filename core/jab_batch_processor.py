import time
import re
from collections import defaultdict
from datetime import date
from decimal import Decimal

from core.data_handler import DataHandler
from core.jab_operator import JABOperator
from core.logger import log
from core.utils import check_abort


class JABBatchProcessor:
    """Batch workflow driven by Excel order and Java Access Bridge table matches."""

    def __init__(self, config):
        self.cfg = config
        self.batch_cfg = config.get("jab_batch", {})
        self.data_handler = DataHandler(config)
        self.jab = JABOperator(config)
        self.match_mode = self.batch_cfg.get("match_mode", "exact")
        self.max_batch_size = int(self.batch_cfg.get("max_batch_size", 50))
        self.save_wait = float(self.batch_cfg.get("save_wait", 0.5))
        self.save_success_timeout = float(self.batch_cfg.get("save_success_timeout", 8.0))
        self.generated_status = self.batch_cfg.get("generated_status", "已生成待回填")
        self.voucher_col = int(self.batch_cfg.get("generated_voucher_col", 22))
        self.verify_voucher_advance = self.batch_cfg.get("verify_voucher_advance", True)
        self.voucher_record_timeout = float(self.batch_cfg.get("voucher_record_timeout", 8.0))
        self.voucher_window_title = self.batch_cfg.get("voucher_window_title", "制单")
        self.voucher_window_class = self.batch_cfg.get("voucher_window_class", "SunAwtDialog")
        self.generated_date_col = self.batch_cfg.get("generated_date_col", 18)
        self.generated_date_value = self.batch_cfg.get(
            "generated_date_value",
            date.today().isoformat(),
        )
        self.generated_voucher_max = int(self.batch_cfg.get("generated_voucher_max", 9999))

    def close(self):
        self.jab.close()

    def load_pending_items(self, skip_filled=True, skip_any_status=False, limit=None):
        items = self.data_handler.load_jab_batch_data(
            skip_filled=skip_filled,
            skip_any_status=skip_any_status,
        )
        if limit:
            items = items[:limit]
        return items

    def dry_run(self, limit=None):
        items = self.load_pending_items(skip_filled=True, skip_any_status=True, limit=limit)
        parsed_items = [item for item in items if not item.get("parse_error")]
        parse_errors = [item for item in items if item.get("parse_error")]

        matches, issues = self.match_current_table(parsed_items)
        batches = self.build_increasing_batches(matches)

        self._log_plan(parsed_items, parse_errors, matches, issues, batches)
        return {
            "items": parsed_items,
            "parse_errors": parse_errors,
            "matches": matches,
            "issues": issues,
            "batches": batches,
        }

    def generate_and_save(self, limit=None, max_batches=None):
        items = self.load_pending_items(skip_filled=True, skip_any_status=True, limit=limit)
        pending = [item for item in items if not item.get("parse_error")]
        parse_errors = [item for item in items if item.get("parse_error")]
        if parse_errors:
            self.data_handler.save_jab_results({
                item["row"]: f"格式错误-{item['parse_error']}"
                for item in parse_errors
            })

        total_saved = 0
        total_batches = 0
        issue_updates = {}

        if pending:
            check_abort()
            matches, issues = self.match_current_table(pending)
            if issues:
                issue_updates.update(self.format_issue_updates(issues))

            if matches:
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
            self.data_handler.save_jab_results(issue_updates)

        log.info(f"JAB 生成保存完成: batches={total_batches}, saved={total_saved}")
        return total_saved

    def backfill_generated_vouchers(self, limit=None):
        items = self.load_pending_items(skip_filled=False, skip_any_status=False, limit=limit)
        items = [item for item in items if not item.get("parse_error")]
        matches, issues = self.match_current_table(
            items,
            voucher_col=self.voucher_col,
            prefer_generated_date=True,
        )

        updates = {}
        for match in matches:
            raw_voucher = str(match["row_data"].get("voucher_text", "")).strip()
            voucher = self.normalize_generated_voucher(raw_voucher)
            if voucher is not None:
                updates[match["item"]["row"]] = voucher
            elif raw_voucher:
                updates[match["item"]["row"]] = f"凭证号异常-{raw_voucher}"
            else:
                updates[match["item"]["row"]] = "已生成未取到凭证号"

        updates.update(self.format_issue_updates(issues, prefix="回填"))
        self.data_handler.save_jab_results(updates)
        log.info(f"JAB 回填完成: vouchers={len(matches)}, issues={len(issues)}")
        return updates

    def switch_to_generated_list(self):
        close_cfg = self.batch_cfg.get("close_voucher_window", {})
        if close_cfg.get("enabled", True):
            self.jab.close_window_by_title(
                close_cfg.get("title", "制单"),
                class_name=close_cfg.get("class_name", "SunAwtDialog"),
                wait=float(close_cfg.get("wait", 0.5)),
            )

        open_query = self.batch_cfg.get("open_query", {})
        if open_query.get("method") == "hotkey":
            query_hwnd = self.jab.wait_window_by_title(
                open_query.get("dialog_title", "查询"),
                class_name=open_query.get("dialog_class", "SunAwtDialog"),
                timeout=0.5,
            )
            if not query_hwnd:
                self.jab.activate_window_by_title(
                    open_query.get("main_title", ""),
                    class_name=open_query.get("main_class"),
                    timeout=float(open_query.get("timeout", 5)),
                )
                self.jab.press_key(
                    open_query.get("key", "f3"),
                    wait=float(open_query.get("wait", 0.8)),
                )
                query_hwnd = self.jab.wait_window_by_title(
                    open_query.get("dialog_title", "查询"),
                    class_name=open_query.get("dialog_class", "SunAwtDialog"),
                    timeout=float(open_query.get("timeout", 5)),
                )
            if not query_hwnd:
                raise RuntimeError("按快捷键后未检测到查询窗口")

        steps = self.batch_cfg.get("switch_generated_steps", [])
        if not steps:
            raise RuntimeError("未配置 switch_generated_steps，暂不能自动切换到已生成列表")
        self.jab.run_named_steps(steps)
        rows = self.jab.read_table_snapshot(voucher_col=self.voucher_col)
        if not rows:
            raise RuntimeError("切换后未读到已生成列表表格")
        with_voucher = [row for row in rows[:50] if row.get("voucher_text")]
        log.info(
            "已切换到疑似已生成列表: "
            f"rows={len(rows)} sample_voucher_count={len(with_voucher)}"
        )
        return True

    def normalize_generated_voucher(self, raw_voucher):
        text = str(raw_voucher or "").strip()
        match = re.search(r"\d+", text)
        if not match:
            return None

        value = int(match.group(0))
        if value <= 0 or value > self.generated_voucher_max:
            return None
        return value

    def match_current_table(self, items, voucher_col=None, prefer_generated_date=False):
        extra_cols = [self.generated_date_col] if prefer_generated_date else None
        snapshot = self.jab.read_table_snapshot(
            voucher_col=voucher_col,
            extra_cols=extra_cols,
        )
        index = defaultdict(list)
        for row in snapshot:
            if row["amount"] is None or not row["partner"]:
                continue
            index[(row["amount"], row["partner"])].append(row)

        matches = []
        issues = []
        for item in items:
            key = (self._as_decimal(item["amount"]), self.jab.normalize_text(item["partner"]))
            rows = index.get(key, [])
            if len(rows) > 1 and prefer_generated_date:
                dated_rows = self.filter_generated_date_rows(rows)
                if dated_rows:
                    rows = dated_rows
            if len(rows) == 1:
                matches.append({
                    "item": item,
                    "nc_row": rows[0]["row_index"],
                    "row_data": rows[0],
                })
            elif not rows and self.match_mode == "contains":
                contains_rows = self._find_contains(snapshot, key)
                self._append_match_or_issue(matches, issues, item, contains_rows)
            else:
                issues.append({
                    "item": item,
                    "reason": "未找到" if not rows else f"重复{len(rows)}条",
                    "rows": [row["row_index"] for row in rows],
                })

        return matches, issues

    def filter_generated_date_rows(self, rows):
        if self.generated_date_col is None or not self.generated_date_value:
            return []

        target = str(self.generated_date_value).strip()
        dated_rows = []
        for row in rows:
            text = str(row.get("extra_text", {}).get(self.generated_date_col, "")).strip()
            if text == target:
                dated_rows.append(row)
        return dated_rows

    def build_increasing_batches(self, matches):
        batches = []
        current = []
        last_nc_row = None

        for match in matches:
            nc_row = match["nc_row"]
            should_split = (
                current
                and (
                    nc_row <= last_nc_row
                    or (self.max_batch_size > 0 and len(current) >= self.max_batch_size)
                )
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
        if not self.jab.select_table_rows(rows):
            raise RuntimeError(f"选中 NC 行失败: {rows}")

        if not self.jab.do_generate_front():
            raise RuntimeError("点击 生成 -> 前台生成 失败")

        time.sleep(self.save_wait)

        pending = list(matches)
        saved_matches = []
        save_batches = 0

        while pending:
            check_abort()
            voucher_matches, issues = self.match_voucher_table(pending)
            if issues:
                detail = "; ".join(
                    f"Excel行{issue['item']['row']} {issue['reason']}"
                    for issue in issues
                )
                raise RuntimeError(f"制单表匹配失败: {detail}")

            voucher_batches = self.build_voucher_increasing_batches(voucher_matches)
            voucher_batch = voucher_batches[0]
            before_count = voucher_batch[0]["table_rows"]
            row_indexes = [match["voucher_row"] for match in voucher_batch]
            log.info(
                "保存制单批次: "
                f"size={len(voucher_batch)} excel_rows={[m['item']['row'] for m in voucher_batch]} "
                f"voucher_rows={row_indexes} before_count={before_count}"
            )

            if not self.jab.select_visible_table_rows(
                voucher_batch[0]["table_index"],
                row_indexes,
                window_title=self.voucher_window_title,
            ):
                raise RuntimeError(f"选中制单表行失败: {row_indexes}")

            if not self.jab.click_save(timeout=self.save_success_timeout):
                raise RuntimeError(
                    f"点击保存失败: Excel行{voucher_batch[0]['item']['row']}"
                )

            verify_result = self.verify_voucher_batch_removed(voucher_batch, before_count)
            saved_matches.extend(voucher_batch)
            saved_rows = {match["item"]["row"] for match in voucher_batch}
            self.data_handler.save_jab_results({
                match["item"]["row"]: self.generated_status
                for match in voucher_batch
            })
            pending = [match for match in pending if match["item"]["row"] not in saved_rows]
            save_batches += 1
            if pending and verify_result in ("empty_window", "window_closed"):
                raise RuntimeError(
                    "制单窗口已空/关闭但仍有未保存 Excel 行，停止复核: "
                    f"remaining_excel_rows={[m['item']['row'] for m in pending]}"
                )
            if max_save_batches and save_batches >= max_save_batches:
                raise RuntimeError(
                    "已达到 max_batches，但全量生成模式不能中途留下制单窗口，"
                    "请不要在全量生成时使用 max_batches"
                )
            time.sleep(self.save_wait)

        self.close_and_verify_pending_removed(saved_matches)
        return saved_matches, save_batches

    def match_voucher_table(self, matches):
        tables = self.jab.read_window_table_cells(
            self.voucher_window_title,
            max_rows=500,
            max_cols=13,
        )
        voucher_tables = [
            table for table in tables
            if table["row_count"] > 0 and table["col_count"] == 13
        ]
        if not voucher_tables:
            raise RuntimeError("未找到制单窗口表格")

        row_records = []
        for table in voucher_tables:
            for row in table["rows"]:
                amount = None
                for cell in row["cells"]:
                    amount = self.jab.normalize_amount(cell)
                    if amount is not None:
                        break
                row_text = "".join(self.jab.normalize_text(cell) for cell in row["cells"])
                if amount is None:
                    continue
                row_records.append({
                    "table": table,
                    "row": row,
                    "amount": amount,
                    "row_text": row_text,
                })

        index = defaultdict(list)
        for record in row_records:
            for match in matches:
                partner = self.jab.normalize_text(match["item"]["partner"])
                if (
                    record["amount"] == self._as_decimal(match["item"]["amount"])
                    and partner in record["row_text"]
                ):
                    index[match["item"]["row"]].append({
                        "table": record["table"],
                        "row": record["row"],
                        "amount": record["amount"],
                        "fallback_reason": "",
                    })

        voucher_matches = []
        issues = []
        assigned_rows = set()
        for ordinal, match in enumerate(matches):
            rows = index.get(match["item"]["row"], [])
            if len(rows) == 1:
                found = rows[0]
                self._append_voucher_match(voucher_matches, match, found, assigned_rows)
                continue

            fallback = self.find_voucher_order_fallback(match, ordinal, matches, row_records)
            if fallback:
                self._append_voucher_match(voucher_matches, match, fallback, assigned_rows)
                log.warning(
                    "制单表金额不一致，按本批顺序+对手方匹配: "
                    f"Excel行{match['item']['row']} expected_amount={match['item']['amount']} "
                    f"voucher_amount={fallback['amount']} voucher_row={fallback['row']['row_index']}"
                )
            else:
                issues.append({
                    "item": match["item"],
                    "reason": "未找到" if not rows else f"重复{len(rows)}条",
                    "rows": [row["row"]["row_index"] for row in rows],
                })

        return voucher_matches, issues

    def _append_voucher_match(self, voucher_matches, match, found, assigned_rows):
        row_key = (
            found["table"]["table_index"],
            found["row"]["row_index"],
        )
        if row_key in assigned_rows:
            return
        assigned_rows.add(row_key)
        voucher_matches.append({
            **match,
            "table_index": found["table"]["table_index"],
            "table_rows": found["table"]["row_count"],
            "voucher_row": found["row"]["row_index"],
            "voucher_cells": found["row"]["cells"],
        })

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
            candidates.append({
                **record,
                "fallback_reason": "order_partner",
            })

        if len(candidates) == 1:
            return candidates[0]
        return None

    def build_voucher_increasing_batches(self, matches):
        batches = []
        current = []
        last_row = None

        for match in matches:
            row = match["voucher_row"]
            should_split = (
                current
                and (
                    row <= last_row
                    or match["table_index"] != current[-1]["table_index"]
                    or (self.max_batch_size > 0 and len(current) >= self.max_batch_size)
                )
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
            tables = self.jab.read_window_table_cells(
                self.voucher_window_title,
                max_rows=500,
                max_cols=13,
            )
            counts = [table["row_count"] for table in tables if table["col_count"] == 13]
            if not counts:
                if window_exists:
                    log.info(
                        "制单窗口仍存在但制单表为空，转待生成表复核: "
                        f"excel_rows={sorted(target_rows)} before={before_count} "
                        f"expected_removed={expected_removed}"
                    )
                    return "empty_window"
                log.info(f"制单窗口已关闭，转待生成表复核: excel_rows={sorted(target_rows)}")
                return "window_closed"

            remaining_matches, issues = self.match_voucher_table(voucher_batch)
            remaining_rows = {match["item"]["row"] for match in remaining_matches}
            min_count = min(counts) if counts else 0

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
            self.jab.close_window_by_title(
                close_cfg.get("title", self.voucher_window_title),
                class_name=close_cfg.get("class_name", self.voucher_window_class),
                wait=float(close_cfg.get("wait", 0.5)),
            )

        self.jab.press_key("f5", wait=float(self.batch_cfg.get("pending_refresh_wait", 1.0)))
        snapshot = self.jab.read_table_snapshot()
        index = defaultdict(list)
        for row in snapshot:
            if row["amount"] is None or not row["partner"]:
                continue
            index[(row["amount"], row["partner"])].append(row)

        still_present = []
        for match in voucher_batch:
            item = match["item"]
            key = (self._as_decimal(item["amount"]), self.jab.normalize_text(item["partner"]))
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
            matches.append({
                "item": item,
                "nc_row": rows[0]["row_index"],
                "row_data": rows[0],
            })
        else:
            issues.append({
                "item": item,
                "reason": "未找到" if not rows else f"重复{len(rows)}条",
                "rows": [row["row_index"] for row in rows],
            })

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
