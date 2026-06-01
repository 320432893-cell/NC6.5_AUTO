from collections import defaultdict
from decimal import Decimal


class NCTableMatcher:
    def __init__(self, processor):
        self.processor = processor

    def __getattr__(self, name):
        return getattr(self.processor, name)

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
        self.run_state.event(
            "table_snapshot_loaded",
            rows=len(snapshot),
            voucher_col=voucher_col,
            prefer_generated_date=prefer_generated_date,
        )
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
