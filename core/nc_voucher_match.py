# 职责：制单表读取与匹配——金额+对手方匹配、多级兜底、汇率一致性求解
# 不做什么：不直接持有 JAB/processor 状态(经 NCVoucherWorkflow 的 self 代理到 processor)
# 允许依赖层：core.errors/models/logger/utils;由 NCVoucherWorkflow 多继承组合
# 谁不应该 import：其它 nc_*_workflow 不应 import(import-linter 独立性约束)

import re
import itertools
from collections import defaultdict

from core.errors import (
    JABControlNotFound,
)
from core.logger import log
from core.models import MatchIssue, VoucherPendingMatch, VoucherSaveMatch


class NCVoucherMatchMixin:
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
            raise JABControlNotFound("未找到制单窗口表格")
        return voucher_tables

    def match_voucher_table(
        self, matches: list[VoucherPendingMatch] | list[VoucherSaveMatch], tables=None
    ) -> tuple[list[VoucherSaveMatch], list[MatchIssue]]:
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
                partner_key = self.normalize_partner_match_text(match.item.partner)
                if partner_key and partner_key in record["partner_key_text"]:
                    partner_index[match.item.row].append(
                        {
                            "table": record["table"],
                            "row": record["row"],
                            "amount": record["amount"],
                            "fallback_reason": "partner_normalized",
                        }
                    )
                if (
                    record["amount"]
                    == self.table_matcher._as_decimal(match.item.amount)
                    and partner_key
                    and partner_key in record["partner_key_text"]
                ):
                    index[match.item.row].append(
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
            rows = index.get(match.item.row, [])
            if len(rows) == 1:
                found = rows[0]
                self._append_voucher_match(voucher_matches, match, found, assigned_rows)
                continue

            partner_rows = partner_index.get(match.item.row, [])
            if len(partner_rows) == 1:
                found = partner_rows[0]
                self._append_voucher_match(voucher_matches, match, found, assigned_rows)
                log.warning(
                    "制单表金额与 Excel 不一致，按唯一对手方匹配: "
                    f"Excel行{match.item.row} expected_amount={match.item.amount} "
                    f"voucher_amount={found['amount']} voucher_row={found['row']['row_index']}"
                )
                continue

            rate_group_match = rate_group_matches.get(match.item.row)
            if rate_group_match:
                self._append_voucher_match(
                    voucher_matches,
                    match,
                    rate_group_match,
                    assigned_rows,
                )
                log.warning(
                    "制单表金额与 Excel 不一致，按归一化对手方+汇率一致性匹配: "
                    f"Excel行{match.item.row} expected_amount={match.item.amount} "
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
                    f"Excel行{match.item.row} expected_amount={match.item.amount} "
                    f"voucher_amount={fallback['amount']} voucher_row={fallback['row']['row_index']}"
                )
            else:
                issues.append(
                    MatchIssue(
                        item=match.item,
                        reason="未找到" if not rows else f"重复{len(rows)}条",
                        rows=[row["row"]["row_index"] for row in rows],
                    )
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
        voucher_match = VoucherSaveMatch(
            item=match.item,
            nc_row=match.nc_row,
            row_data=match.row_data,
            table_index=found["table"]["table_index"],
            table_rows=found["table"]["row_count"],
            voucher_row=found["row"]["row_index"],
            voucher_cells=found["row"]["cells"],
        )
        voucher_match.validate_for_save(context="voucher_match")
        voucher_matches.append(voucher_match)

    def find_voucher_order_fallback(self, match, ordinal, matches, row_records):
        """NC sometimes changes voucher amount during front generation.

        Only trust the fallback when the generated voucher table has the same
        row count as the current batch and the row at the same ordinal contains
        the expected partner name.
        """
        partner = self.jab.normalize_text(match.item.partner)
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
            partner_key = self.normalize_partner_match_text(match.item.partner)
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
                result[match.item.row] = {
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
                excel_amount = self.table_matcher._as_decimal(match.item.amount)
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
