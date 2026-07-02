import time

from core.errors import (
    ContractViolation,
    JABActionError,
    TableMatchError,
    WorkflowStateError,
)
from core.logger import log
from core.models import (
    ExcelVoucherItem,
    MatchIssue,
    PendingMatch,
    VoucherPendingMatch,
    VoucherSaveMatch,
)
from core.utils import check_abort
from core.voucher_plan_cache import (
    issues_from_plan_cache,
    load_voucher_plan_cache,
    matches_from_plan_cache,
    validate_voucher_plan_cache,
    write_voucher_plan_cache,
)


class NCPendingWorkflow:
    def __init__(self, processor):
        self.processor = processor

    def __getattr__(self, name):
        return getattr(self.processor, name)

    def load_pending_items(
        self,
        skip_filled=True,
        skip_any_status=False,
        limit=None,
        start_row=None,
        end_row=None,
    ) -> list[ExcelVoucherItem]:
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
            items = [item for item in items if item.row >= start_row]
        if end_row is not None:
            items = [item for item in items if item.row <= end_row]
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
        parsed_items = [item for item in items if not item.parse_error]
        parse_errors = [item for item in items if item.parse_error]
        preflight = self.data_handler.preflight_jab_items(
            items,
            start_row=start_row,
            end_row=end_row,
            limit=limit,
            skip_any_status=True,
            context="plan",
        )
        self.run_state.event("excel_preflight_passed", **preflight)

        self.require_page_state("pending", parsed_items, command="plan")
        self.run_state.set_stage("plan_match_pending_table")
        matches, issues = self.table_matcher.match_current_table(parsed_items)
        batches = self.table_matcher.build_increasing_batches(matches)
        plan_path = None
        if matches or issues:
            plan_path = write_voucher_plan_cache(
                config=self.cfg,
                limit=limit,
                start_row=start_row,
                end_row=end_row,
                matches=matches,
                issues=issues,
            )
        self.run_state.update_counts(
            parse_errors=len(parse_errors),
            matches=len(matches),
            issues=len(issues),
            batches=len(batches),
        )
        if plan_path:
            self.run_state.event(
                "voucher_precheck_plan_cached",
                path=str(plan_path),
                rows=len(matches),
                excel_rows=[match.item.row for match in matches],
                nc_rows=[match.nc_row for match in matches],
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
        result = self.generate_and_collect_saved(
            limit=limit,
            max_batches=max_batches,
            start_row=start_row,
            end_row=end_row,
        )
        return result["saved"]

    def generate_and_collect_saved(
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
            pending = [item for item in items if not item.parse_error]
            parse_errors = [item for item in items if item.parse_error]
            preflight = self.data_handler.preflight_jab_items(
                items,
                start_row=start_row,
                end_row=end_row,
                limit=limit,
                skip_any_status=True,
                context="generate",
            )
            self.run_state.event("excel_preflight_passed", **preflight)
            self.run_state.set_stage("excel_write_preflight", context="generate")
            self.data_handler.assert_excel_writable("正式生成前写入凭证状态/凭证号预检")
            self.run_state.event("excel_write_preflight_passed", context="generate")
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
                detail = "; ".join(
                    f"Excel行{item.row} {item.parse_error}" for item in parse_errors
                )
                raise TableMatchError(
                    "正式生成前 Excel 存在格式错误，已暂停，未写 Excel、未点击 NC。"
                    f" 请先预检查并修正。异常: {detail}"
                )

            total_saved = 0
            total_batches = 0
            issue_updates = {}
            saved_matches: list[VoucherSaveMatch] = []

            if pending:
                check_abort()
                self.run_state.set_stage(
                    "generate_load_precheck_plan",
                    excel_rows=[item.row for item in pending],
                )
                cache = load_voucher_plan_cache()
                validate_voucher_plan_cache(
                    cache=cache,
                    config=self.cfg,
                    limit=limit,
                    start_row=start_row,
                    end_row=end_row,
                    pending=pending,
                )
                matches = matches_from_plan_cache(cache, pending)
                issues: list[MatchIssue] = issues_from_plan_cache(cache, pending)
                self.perf.event(
                    "pending_match_from_precheck_plan",
                    pending=len(pending),
                    matches=len(matches),
                    excel_rows=[match.item.row for match in matches],
                    nc_rows=[match.nc_row for match in matches],
                )
                self.run_state.event(
                    "pending_match_from_precheck_plan",
                    rows=len(matches),
                    issues=len(issues),
                    excel_rows=[match.item.row for match in matches],
                    nc_rows=[match.nc_row for match in matches],
                )

                self.ensure_full_pending_match(pending, matches, issues)
                if issues and self.duplicate_match_policy == "skip":
                    issue_updates = self.format_issue_updates(issues, prefix="跳过：")
                    self.run_state.set_stage(
                        "generate_write_skipped_rows",
                        rows=len(issue_updates),
                    )
                    with self.perf.span("generate_write_skipped_rows", rows=len(issue_updates)):
                        self.data_handler.save_jab_results(issue_updates)
                    self.record_event(
                        "pending_issues_skipped",
                        rows=len(issue_updates),
                        excel_rows=sorted(issue_updates),
                        reasons=issue_updates,
                    )

                self.run_state.update_counts(
                    pending_matches=len(matches),
                    pending_issues=len(issues),
                    skipped=len(issue_updates),
                )
                if matches:
                    matched_pending = [match.item for match in matches]
                    log.info(
                        "开始执行 JAB 生成: "
                        f"size={len(matches)} excel_rows={[m.item.row for m in matches]} "
                        f"nc_rows={[m.nc_row for m in matches]} skipped={len(issue_updates)}"
                    )
                    saved_matches, total_batches = self.process_pending_generate_batches(
                        matched_pending,
                        matches,
                        max_save_batches=max_batches,
                    )
                    total_saved = len(saved_matches)

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
            return {
                "saved": total_saved,
                "batches": total_batches,
                "skipped": len(issue_updates),
                "issue_updates": issue_updates,
                "saved_matches": saved_matches,
                "excel_rows": [match.item.row for match in saved_matches],
            }


    def ensure_full_pending_match(self, pending, matches, issues) -> None:
        if self.duplicate_match_policy == "skip":
            return
        pending_rows = {item.row for item in pending}
        matched_rows = {match.item.row for match in matches}
        missing_rows = sorted(pending_rows - matched_rows)
        if not issues and not missing_rows and len(matches) == len(pending):
            return

        issue_detail = "; ".join(
            f"Excel行{issue.item.row} {issue.reason}"
            + (f" NC行{issue.rows}" if issue.rows else "")
            for issue in issues
        )
        missing_detail = f"未匹配Excel行{missing_rows}" if missing_rows else ""
        detail = "；".join(part for part in (issue_detail, missing_detail) if part)
        raise TableMatchError(
            "待生成表未全量匹配，已暂停，未点击生成。"
            "请先处理预检查提示后再批量生成。"
            f" 异常: {detail or '匹配数量不一致'}"
        )

    def ensure_full_voucher_match(self, pending, voucher_matches) -> None:
        pending_rows = {match.item.row for match in pending}
        matched_rows = {match.item.row for match in voucher_matches}
        missing_rows = sorted(pending_rows - matched_rows)
        if not missing_rows and len(voucher_matches) == len(pending):
            return
        raise TableMatchError(
            "制单表未全量匹配，已暂停，未保存。"
            "请人工检查前台生成后的制单表。"
            f" 待保存{len(pending)}行，匹配到{len(voucher_matches)}行，"
            f"未匹配Excel行{missing_rows}"
        )

    def process_pending_generate_batches(
        self,
        pending,
        matches,
        max_save_batches=None,
    ):
        cap = int(self.batch_cfg.get("pending_generate_batch_size", 120) or 0)
        if cap <= 0 or len(matches) <= cap:
            return self.process_full_selection(
                matches,
                max_save_batches=max_save_batches,
            )
        if max_save_batches:
            raise ContractViolation(
                "已启用待生成页分批生成，不能同时使用 max_batches；"
                "请取消 max_batches 或调大待生成页单批上限。"
            )

        remaining = list(pending)
        first_matches_by_excel_row = {match.item.row: match for match in matches}
        saved_matches: list[VoucherSaveMatch] = []
        total_save_batches = 0
        pending_batch_index = 0

        self.record_event(
            "pending_generate_batches_start",
            total=len(remaining),
            pending_generate_batch_size=cap,
            estimated_batches=(len(remaining) + cap - 1) // cap,
            excel_rows=[item.row for item in remaining],
        )

        while remaining:
            check_abort()
            pending_batch_index += 1
            self.run_state.set_stage(
                "pending_batch_match",
                pending_batch_index=pending_batch_index,
                remaining=len(remaining),
                pending_generate_batch_size=cap,
                excel_rows=[item.row for item in remaining],
            )

            if pending_batch_index == 1:
                current_matches = [
                    first_matches_by_excel_row[item.row]
                    for item in remaining
                    if item.row in first_matches_by_excel_row
                ]
                current_issues: list[MatchIssue] = []
            else:
                self.require_page_state("pending", remaining, command="generate")
                current_matches, current_issues = self.table_matcher.match_current_table(
                    remaining
                )
            self.ensure_full_pending_match(remaining, current_matches, current_issues)

            batch_matches = current_matches[:cap]
            self.record_event(
                "pending_generate_batch",
                pending_batch_index=pending_batch_index,
                rows=len(batch_matches),
                remaining_before=len(remaining),
                excel_rows=[match.item.row for match in batch_matches],
                nc_rows=[match.nc_row for match in batch_matches],
            )
            new_saved, new_save_batches = self.process_full_selection(batch_matches)
            saved_matches.extend(new_saved)
            total_save_batches += new_save_batches

            saved_excel_rows = {match.item.row for match in new_saved}
            if not saved_excel_rows:
                raise TableMatchError(
                    "待生成页分批生成未保存任何凭证，已停止后续批次。"
                )
            remaining = [
                item for item in remaining if item.row not in saved_excel_rows
            ]
            self.record_event(
                "pending_generate_batch_done",
                pending_batch_index=pending_batch_index,
                saved=len(new_saved),
                save_batches=new_save_batches,
                remaining_after=len(remaining),
                saved_excel_rows=sorted(saved_excel_rows),
            )

            if remaining:
                state = (
                    self.processor.switch_generated_workflow
                    .wait_for_parent_ready_after_voucher_close()
                )
                if state != "parent_ready":
                    raise WorkflowStateError(
                        "分批生成后未回到制单父界面，已停止后续批次: "
                        f"state={state}"
                    )

        return saved_matches, total_save_batches

    def process_full_selection(self, matches, max_save_batches=None):
        rows = [match.nc_row for match in matches]
        self.run_state.set_stage(
            "pending_rows_select",
            excel_rows=[match.item.row for match in matches],
            nc_rows=rows,
        )
        self.perf.event(
            "full_selection_start",
            matches=len(matches),
            excel_rows=[match.item.row for match in matches],
            nc_rows=rows,
        )
        self.record_event(
            "event_pending_rows_select_start",
            rows=len(rows),
            excel_rows=[match.item.row for match in matches],
            nc_rows=rows,
        )
        with self.perf.span("pending_rows_select", rows=len(rows)):
            if not self.jab.select_table_rows(rows):
                raise JABActionError(f"选中 NC 行失败: {rows}")
        self.record_event(
            "event_pending_rows_selected",
            rows=len(rows),
            nc_rows=rows,
        )

        self.run_state.set_stage("front_generate_click", nc_rows=rows)
        self.record_event("event_front_generate_click", rows=len(rows), nc_rows=rows)
        with self.perf.span("front_generate_click", rows=len(rows)):
            self.maximize_batch_main_window()
            if not self.jab.do_generate_front():
                raise JABActionError("点击 生成 -> 前台生成 失败")
        self.record_transition(
            "front_generate_clicked",
            from_state="pending",
            to_state="voucher_open",
            rows=len(rows),
        )

        pending: list[PendingMatch] = list(matches)
        with self.perf.span("voucher_window_open_wait", timeout=self.save_wait):
            voucher_tables = self.processor.voucher_workflow.wait_for_voucher_tables(
                len(pending),
                timeout=self.save_wait,
            )
        saved_matches: list[VoucherSaveMatch] = []
        save_batches = 0

        while pending:
            check_abort()
            self.run_state.set_stage(
                "voucher_match",
                pending_excel_rows=[match.item.row for match in pending],
                save_batch_index=save_batches + 1,
            )
            with self.perf.span(
                "voucher_match",
                pending=len(pending),
                save_batch_index=save_batches + 1,
            ):
                voucher_matches, issues = (
                    self.processor.voucher_workflow.match_voucher_table(
                        pending,
                        tables=voucher_tables,
                    )
                )
                voucher_tables = None
            self.perf.event(
                "voucher_match_done",
                pending=len(pending),
                matches=len(voucher_matches),
                issues=len(issues),
                save_batch_index=save_batches + 1,
            )
            if issues:
                detail = "; ".join(
                    f"Excel行{issue.item.row} {issue.reason}" for issue in issues
                )
                raise TableMatchError(f"制单表匹配失败: {detail}")
            self.ensure_full_voucher_match(pending, voucher_matches)

            new_saved, new_batches = (
                self.processor.voucher_workflow.save_current_voucher_matches(
                    voucher_matches
                )
            )
            saved_matches.extend(new_saved)
            saved_rows = {match.item.row for match in new_saved}
            if max_save_batches and save_batches >= max_save_batches:
                raise ContractViolation(
                    "已达到 max_batches，但全量生成模式不能中途留下制单窗口，"
                    "请不要在全量生成时使用 max_batches"
                )
            pending = [match for match in pending if match.item.row not in saved_rows]
            save_batches += new_batches
            if pending:
                with self.perf.span(
                    "save_batch_wait",
                    wait=self.wait_between_save_batches,
                    save_batch_index=save_batches,
                ):
                    if self.wait_between_save_batches > 0:
                        time.sleep(self.wait_between_save_batches)

        self.run_state.set_stage(
            "voucher_window_close",
            saved_excel_rows=[match.item.row for match in saved_matches],
        )
        with self.perf.span("voucher_window_close", rows=len(saved_matches)):
            self.processor.voucher_workflow.close_voucher_window_after_save(
                saved_matches
            )
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
            pending = [item for item in items if not item.parse_error]
            if not pending:
                log.info("当前没有可用于恢复保存的 Excel 行")
                return 0

            self.run_state.set_stage(
                "resume_read_voucher_window",
                excel_rows=[item.row for item in pending],
            )
            self.require_page_state(
                "voucher_open",
                pending,
                command="resume-voucher",
            )
            tables = self.processor.voucher_workflow.read_voucher_tables(len(pending))
            matches: list[VoucherPendingMatch] = [
                VoucherPendingMatch(
                    item=item,
                    nc_row=None,
                    row_data={},
                )
                for item in pending
            ]
            voucher_matches, issues = (
                self.processor.voucher_workflow.match_voucher_table(
                    matches, tables=tables
                )
            )
            if issues:
                table_rows = sum(table["row_count"] for table in tables)
                if len(voucher_matches) != table_rows:
                    detail = "; ".join(
                        f"Excel行{issue.item.row} {issue.reason}" for issue in issues
                    )
                    raise TableMatchError(f"恢复制单表匹配失败: {detail}")
                log.warning(
                    "恢复保存时忽略未进入当前制单窗口的 Excel 行: "
                    f"excel_rows={[issue.item.row for issue in issues]}"
                )

            saved_matches, save_batches = (
                self.processor.voucher_workflow.save_current_voucher_matches(
                    voucher_matches
                )
            )
            self.run_state.set_stage(
                "resume_voucher_window_close",
                saved_excel_rows=[match.item.row for match in saved_matches],
            )
            with self.perf.span("resume_voucher_window_close", rows=len(saved_matches)):
                self.processor.voucher_workflow.close_voucher_window_after_save(
                    saved_matches
                )
            log.info(
                f"JAB 恢复制单窗口保存完成: batches={save_batches}, saved={len(saved_matches)}"
            )
            self.run_state.update_counts(
                resume_save_batches=save_batches,
                resume_saved=len(saved_matches),
            )
            self.run_state.set_stage("resume_done")
            return len(saved_matches)

    def format_issue_updates(self, issues: list[MatchIssue], prefix=""):
        updates: dict[int, str] = {}
        for issue in issues:
            reason = issue.reason
            if issue.rows:
                reason = f"{reason}-NC行{','.join(str(r) for r in issue.rows[:5])}"
            updates[issue.item.row] = f"{prefix}{reason}"
        return updates

    def maximize_batch_main_window(self):
        open_query = self.batch_cfg.get("open_query", {})
        title = open_query.get("main_title", "")
        if not title:
            return False
        return self.jab.maximize_window_by_title(
            title,
            class_name=open_query.get("main_class"),
            timeout=float(open_query.get("activate_timeout", 1.0)),
        )

    def _log_plan(self, items, parse_errors, matches, issues, batches):
        log.info(
            "JAB dry-run: "
            f"items={len(items)} parse_errors={len(parse_errors)} "
            f"matched={len(matches)} issues={len(issues)} batches={len(batches)}"
        )
        for issue in parse_errors:
            log.warning(f"Excel行{issue.row} 格式错误: {issue.parse_error}")
        for issue in issues:
            log.warning(
                f"Excel行{issue.item.row} {issue.reason}: "
                f"amount={issue.item.amount} partner={issue.item.partner} "
                f"nc_rows={issue.rows}"
            )
        for index, batch in enumerate(batches, start=1):
            log.info(
                f"批次{index}: size={len(batch)} "
                f"excel_rows={[m.item.row for m in batch]} "
                f"nc_rows={[m.nc_row for m in batch]}"
            )
