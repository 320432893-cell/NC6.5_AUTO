from core.errors import WorkflowStateError
from core.logger import log
from core.models import BackfillUpdates, BackfillUpdateValue, ExcelVoucherItem


class NCBackfillWorkflow:
    def __init__(self, processor):
        self.processor = processor

    def __getattr__(self, name):
        return getattr(self.processor, name)

    def backfill_generated_vouchers(
        self,
        limit=None,
        start_row=None,
        end_row=None,
        auto_switch=True,
    ):
        with self.perf.span(
            "backfill_total",
            limit=limit,
            start_row=start_row,
            end_row=end_row,
            auto_switch=auto_switch,
        ):
            self.run_state.set_stage(
                "backfill_load_excel",
                limit=limit,
                start_row=start_row,
                end_row=end_row,
                auto_switch=auto_switch,
            )
            items = self.processor.pending_workflow.load_pending_items(
                skip_filled=False,
                skip_any_status=False,
                limit=limit,
                start_row=start_row,
                end_row=end_row,
            )
            items: list[ExcelVoucherItem] = [
                item
                for item in items
                if not item.get("parse_error")
                and str(item.get("voucher") or "").strip() == self.generated_status
            ]
            with self.perf.span("excel_save_split_columns", rows=len(items)):
                self.data_handler.save_jab_split_columns(items)
            self.perf.event("backfill_items_prepared", rows=len(items))
            self.run_state.update_counts(backfill_pending=len(items))
            if not items:
                self.perf.event(
                    "backfill_done",
                    vouchers=0,
                    issues=0,
                    updates=0,
                    perf_path=str(self.perf.path) if self.perf.path else "",
                )
                self.run_state.set_stage("backfill_done")
                log.info("JAB 回填完成: 没有待回填行")
                return {}
            self.ensure_generated_page_for_backfill(items, auto_switch=auto_switch)
            self.run_state.set_stage(
                "backfill_match_generated_table",
                excel_rows=[item["row"] for item in items],
            )
            self.require_page_state("generated", items, command="backfill")
            matches, issues = self.table_matcher.match_generated_voucher_table(
                items,
                voucher_col=self.voucher_col,
            )
            self.perf.event(
                "backfill_match_done",
                items=len(items),
                matches=len(matches),
                issues=len(issues),
            )

            updates: BackfillUpdates = {}
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

            updates.update(
                self.processor.pending_workflow.format_issue_updates(
                    issues, prefix="回填"
                )
            )
            self.validate_backfill_updates(updates)
            self.run_state.set_stage("backfill_write_excel", rows=len(updates))
            with self.perf.span("backfill_excel_save", rows=len(updates)):
                self.data_handler.save_jab_results(updates)
            self.perf.event(
                "backfill_done",
                vouchers=len(matches),
                issues=len(issues),
                updates=len(updates),
                perf_path=str(self.perf.path) if self.perf.path else "",
            )
            self.run_state.update_counts(
                backfill_matches=len(matches),
                backfill_issues=len(issues),
                backfill_updates=len(updates),
            )
            self.run_state.set_stage("backfill_done")
            log.info(f"JAB 回填完成: vouchers={len(matches)}, issues={len(issues)}")
            return updates

    def ensure_generated_page_for_backfill(self, items, auto_switch=True):
        state = self.detect_page_state(items)
        self.record_event(
            "backfill_preflight_state",
            state=state.name,
            reason=state.reason,
            auto_switch=auto_switch,
        )
        if state.name == "generated":
            return state
        if state.name == "pending" and auto_switch:
            self.record_transition(
                "backfill_auto_switch_generated",
                from_state="pending",
                to_state="generated",
                rows=len(items),
            )
            self.switch_to_generated_list()
            return self.require_page_state("generated", items, command="backfill")

        if state.name == "pending":
            detail = "当前在待生成页，未启用自动切换"
        else:
            detail = "当前页面不适合回填，不能按已生成表列位读取"
        raise WorkflowStateError(
            f"回填前页面状态不正确: state={state.name} reason={state.reason}; {detail}"
        )

    def validate_backfill_updates(self, updates: BackfillUpdates):
        invalid = {
            row: value
            for row, value in updates.items()
            if not self.is_valid_backfill_update_value(value)
        }
        if invalid:
            raise WorkflowStateError(
                "回填更新值不符合契约: "
                f"invalid={invalid} allowed=positive voucher int or non-empty status text"
            )

    def is_valid_backfill_update_value(self, value: BackfillUpdateValue):
        if isinstance(value, int):
            return 1 <= value <= self.generated_voucher_max
        if isinstance(value, str):
            return value.strip() != ""
        return False
