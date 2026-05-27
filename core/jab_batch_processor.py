import time
import re
from collections import defaultdict
from datetime import date
from decimal import Decimal

from core.data_handler import DataHandler
from core.jab_operator import JABOperator
from core.logger import log
from core.perf import PerfRecorder
from core.utils import check_abort


class JABBatchProcessor:
    """Batch workflow driven by Excel order and Java Access Bridge table matches."""

    def __init__(self, config, perf_enabled=False, perf_label=None):
        self.cfg = config
        self.batch_cfg = config.get("jab_batch", {})
        self.data_handler = DataHandler(config)
        self.jab = JABOperator(config)
        self.perf = PerfRecorder(enabled=perf_enabled, label=perf_label)
        self.match_mode = self.batch_cfg.get("match_mode", "exact")
        self.save_strategy = self.batch_cfg.get("save_strategy", "batch_reverse_select")
        self.voucher_selection_order = self.batch_cfg.get(
            "voucher_selection_order", "reverse_excel"
        )
        self.max_batch_size = int(self.batch_cfg.get("max_batch_size", 50))
        self.save_wait = float(self.batch_cfg.get("save_wait", 0.5))
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
        self.generated_date_value = self.batch_cfg.get(
            "generated_date_value",
            date.today().isoformat(),
        )
        self.generated_voucher_max = int(
            self.batch_cfg.get("generated_voucher_max", 9999)
        )

    def close(self):
        self.jab.close()

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
        return items

    def dry_run(self, limit=None, start_row=None, end_row=None):
        items = self.load_pending_items(
            skip_filled=True,
            skip_any_status=True,
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        )
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
            if parse_errors:
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
            log.info(f"JAB 生成保存完成: batches={total_batches}, saved={total_saved}")
            return total_saved

    def backfill_generated_vouchers(self, limit=None, start_row=None, end_row=None):
        with self.perf.span(
            "backfill_total",
            limit=limit,
            start_row=start_row,
            end_row=end_row,
        ):
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
            if not items:
                self.perf.event(
                    "backfill_done",
                    vouchers=0,
                    issues=0,
                    updates=0,
                    perf_path=str(self.perf.path) if self.perf.path else "",
                )
                log.info("JAB 回填完成: 没有待回填行")
                return {}
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
            with self.perf.span("backfill_excel_save", rows=len(updates)):
                self.data_handler.save_jab_results(updates)
            self.perf.event(
                "backfill_done",
                vouchers=len(matches),
                issues=len(issues),
                updates=len(updates),
                perf_path=str(self.perf.path) if self.perf.path else "",
            )
            log.info(f"JAB 回填完成: vouchers={len(matches)}, issues={len(issues)}")
            return updates

    def switch_to_generated_list(self):
        with self.perf.span("switch_generated_total"):
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
            query_hwnd = self.find_query_window(open_query, timeout=0.5)
            if query_method == "jab_action":
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
            elif query_method == "hotkey":
                with self.perf.span("switch_open_query"):
                    if not query_hwnd:
                        self.open_query_with_hotkey(open_query)
                        query_hwnd = self.find_query_window(
                            open_query,
                            timeout=float(open_query.get("timeout", 5)),
                        )
                    if not query_hwnd:
                        raise RuntimeError("按快捷键后未检测到查询窗口")
            elif query_method:
                raise RuntimeError(f"不支持的 open_query.method: {query_method}")

            steps = self.batch_cfg.get("switch_generated_steps", [])
            if not steps:
                raise RuntimeError(
                    "未配置 switch_generated_steps，暂不能自动切换到已生成列表"
                )
            if query_method == "jab_action":
                self.jab.activate_window_by_title(
                    open_query.get("main_title", ""),
                    class_name=open_query.get("main_class"),
                    timeout=float(open_query.get("activate_timeout", 1)),
                )
            try:
                with self.perf.span("switch_run_steps", steps=len(steps)):
                    self.jab.run_named_steps(steps)
            except RuntimeError:
                if (
                    query_method == "jab_action"
                    and open_query.get("fallback_method") == "hotkey"
                ):
                    log.warning("JAB 查询窗口步骤失败，回退 F3 重新打开查询窗口")
                    self.open_query_with_hotkey(open_query)
                    with self.perf.span("switch_run_steps_fallback", steps=len(steps)):
                        self.jab.run_named_steps(steps)
                else:
                    raise
            with self.perf.span("switch_generated_snapshot"):
                rows = self.jab.read_table_snapshot(voucher_col=self.voucher_col)
            if not rows:
                raise RuntimeError("切换后未读到已生成列表表格")
            with_voucher = [row for row in rows[:50] if row.get("voucher_text")]
            log.info(
                "已切换到疑似已生成列表: "
                f"rows={len(rows)} sample_voucher_count={len(with_voucher)}"
            )
            return True

    def find_query_window(self, open_query, timeout):
        return self.jab.wait_window_by_title(
            open_query.get("dialog_title", "查询"),
            class_name=open_query.get("dialog_class", "SunAwtDialog"),
            timeout=timeout,
            include_children=bool(open_query.get("dialog_include_children", True)),
            visible_only=bool(open_query.get("dialog_visible_only", True)),
        )

    def open_query_with_jab_action(self, open_query):
        if bool(open_query.get("wait_dialog", True)):
            thread = self.jab.trigger_action_by_path_async(
                open_query.get("path"),
                title=open_query.get("window_title"),
                class_name=open_query.get("window_class"),
                name=open_query.get("name"),
                role=open_query.get("role"),
                action_name=open_query.get("action"),
                timeout=float(open_query.get("return_timeout", 0.2)),
                require_showing=bool(open_query.get("require_showing", False)),
            )
            if not thread:
                raise RuntimeError(f"JAB 查询入口未找到: path={open_query.get('path')}")
            return

        if not self.jab.do_action_by_path(
            open_query.get("path"),
            title=open_query.get("window_title"),
            class_name=open_query.get("window_class"),
            name=open_query.get("name"),
            role=open_query.get("role"),
            action_name=open_query.get("action"),
            wait=float(open_query.get("wait", 0.8)),
            timeout=float(open_query.get("timeout", 5)),
            require_showing=bool(open_query.get("require_showing", False)),
        ):
            raise RuntimeError(f"JAB 查询入口执行失败: path={open_query.get('path')}")

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
        self.perf.event(
            "full_selection_start",
            matches=len(matches),
            excel_rows=[match["item"]["row"] for match in matches],
            nc_rows=rows,
        )
        with self.perf.span("pending_rows_select", rows=len(rows)):
            if not self.jab.select_table_rows(rows):
                raise RuntimeError(f"选中 NC 行失败: {rows}")

        with self.perf.span("front_generate_click", rows=len(rows)):
            if not self.jab.do_generate_front():
                raise RuntimeError("点击 生成 -> 前台生成 失败")

        with self.perf.span("front_generate_wait", wait=self.save_wait):
            time.sleep(self.save_wait)

        pending = list(matches)
        saved_matches = []
        save_batches = 0

        while pending:
            check_abort()
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
            with self.perf.span("resume_pending_final_verify", rows=len(saved_matches)):
                self.close_and_verify_pending_removed(saved_matches)
            log.info(
                f"JAB 恢复制单窗口保存完成: batches={save_batches}, saved={len(saved_matches)}"
            )
            return len(saved_matches)

    def save_current_voucher_matches(self, voucher_matches):
        pending_source = list(voucher_matches)
        saved_matches = []
        save_batches = 0

        while pending_source:
            check_abort()
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

            pending = refreshed_matches
            voucher_batches = self.build_voucher_save_batches(pending)
            voucher_batch = voucher_batches[0]
            before_count = voucher_batch[0]["table_rows"]
            row_indexes = [match["voucher_row"] for match in voucher_batch]
            selection_row_indexes = self.get_voucher_selection_rows(voucher_batch)
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

            with self.perf.span(
                "voucher_save_click",
                rows=len(voucher_batch),
                excel_rows=[match["item"]["row"] for match in voucher_batch],
                save_batch_index=save_batches + 1,
            ):
                if not self.jab.click_save(timeout=self.save_success_timeout):
                    raise RuntimeError(
                        f"点击保存失败: Excel行{voucher_batch[0]['item']['row']}"
                    )

            with self.perf.span(
                "voucher_save_verify",
                rows=len(voucher_batch),
                save_batch_index=save_batches + 1,
            ):
                verify_result = self.verify_voucher_batch_removed(
                    voucher_batch, before_count
                )
            saved_matches.extend(voucher_batch)
            saved_rows = {match["item"]["row"] for match in voucher_batch}
            with self.perf.span(
                "excel_save_generated_status",
                rows=len(voucher_batch),
                save_batch_index=save_batches + 1,
            ):
                self.data_handler.save_jab_results(
                    {
                        match["item"]["row"]: self.generated_status
                        for match in voucher_batch
                    }
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
            if pending_source:
                with self.perf.span(
                    "save_batch_wait",
                    wait=self.save_wait,
                    save_batch_index=save_batches,
                ):
                    time.sleep(self.save_wait)

        return saved_matches, save_batches

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
                    }
                )

        index = defaultdict(list)
        partner_index = defaultdict(list)
        for record in row_records:
            for match in matches:
                partner = self.jab.normalize_text(match["item"]["partner"])
                if partner and partner in record["row_text"]:
                    partner_index[match["item"]["row"]].append(
                        {
                            "table": record["table"],
                            "row": record["row"],
                            "amount": record["amount"],
                            "fallback_reason": "partner_only",
                        }
                    )
                if (
                    record["amount"] == self._as_decimal(match["item"]["amount"])
                    and partner in record["row_text"]
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

            fallback = self.find_voucher_order_fallback(
                match, ordinal, matches, row_records
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

    def build_voucher_save_batches(self, matches):
        if self.save_strategy == "single":
            return [[match] for match in matches]
        if self.save_strategy in ("batch", "batch_reverse_select"):
            return self.build_voucher_batch_groups(matches)
        if self.save_strategy == "bottom_up":
            return self.build_voucher_bottom_up_batches(matches)
        raise ValueError(f"不支持的保存策略: {self.save_strategy!r}")

    def build_voucher_batch_groups(self, matches):
        batches = []
        current = []

        for match in matches:
            should_split = current and (
                match["table_index"] != current[-1]["table_index"]
                or (self.max_batch_size > 0 and len(current) >= self.max_batch_size)
            )
            if should_split:
                batches.append(current)
                current = []

            current.append(match)

        if current:
            batches.append(current)
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
        rows = [match["voucher_row"] for match in voucher_batch]
        if self.voucher_selection_order == "reverse_excel":
            return list(reversed(rows))
        if self.voucher_selection_order == "excel":
            return rows
        raise ValueError(f"不支持的选择顺序: {self.voucher_selection_order!r}")

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
            self.jab.close_window_by_title(
                close_cfg.get("title", self.voucher_window_title),
                class_name=close_cfg.get("class_name", self.voucher_window_class),
                wait=float(close_cfg.get("wait", 0.5)),
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
