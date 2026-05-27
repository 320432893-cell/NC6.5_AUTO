import openpyxl
import re
from decimal import Decimal, InvalidOperation

from core.logger import log


CONCAT_KEY_RE = re.compile(
    r"^\s*([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s*(.+?)\s*$"
)


class DataHandler:
    def __init__(self, config):
        self.cfg = config
        self.excel_path = config["excel_path"]
        self.sheet_my = config.get("sheet_my", "Sheet1")
        self.has_header = config.get("has_header", True)
        batch_cfg = config.get("jab_batch", {})
        self.jab_key_col = batch_cfg.get("key_col", 1)
        self.jab_result_col = batch_cfg.get("result_col", 2)
        self.jab_amount_out_col = batch_cfg.get("amount_out_col", 3)
        self.jab_partner_out_col = batch_cfg.get("partner_out_col", 4)

    def load_jab_batch_data(self, skip_filled=True, skip_any_status=False):
        """读取 Sheet1 的“金额+对手方”拼接列，保持 Excel 行顺序。"""
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb[self.sheet_my]

        start_row = 2 if self.has_header else 1
        data = []

        for row in range(start_row, ws.max_row + 1):
            raw_key = ws.cell(row=row, column=self.jab_key_col).value
            result = ws.cell(row=row, column=self.jab_result_col).value

            if raw_key is None or (isinstance(raw_key, str) and raw_key.strip() == ""):
                continue
            if skip_any_status and self._has_cell_value(result):
                continue
            if skip_filled and self._looks_like_voucher(result):
                continue

            try:
                amount, partner = self.parse_jab_concat_key(raw_key)
            except ValueError as e:
                log.warning(f"行{row} 拼接索引格式错误: {raw_key!r}, {e}")
                data.append({
                    "row": row,
                    "raw_key": raw_key,
                    "amount": None,
                    "partner": "",
                    "voucher": result,
                    "parse_error": str(e),
                })
                continue

            data.append({
                "row": row,
                "raw_key": raw_key,
                "amount": amount,
                "partner": partner,
                "voucher": result,
                "parse_error": "",
            })

        wb.close()
        log.info(f"加载 JAB 批量数据: {len(data)} 条")
        return data

    def parse_jab_concat_key(self, value):
        text = str(value).strip()
        match = CONCAT_KEY_RE.match(text)
        if not match:
            raise ValueError("需要以金额开头，后面紧跟对手方名称")

        amount_text, partner = match.groups()
        partner = "".join(partner.split())
        if not partner:
            raise ValueError("对手方名称为空")

        try:
            amount = Decimal(amount_text.replace(",", "")).quantize(Decimal("0.01"))
        except (InvalidOperation, ValueError) as e:
            raise ValueError(f"金额格式无法识别: {amount_text!r}") from e

        return amount, partner

    def split_jab_keys_to_columns(self, limit=None):
        """把“金额+对手方”拼接列拆成独立金额列和对手方列。"""
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb[self.sheet_my]

        start_row = 2 if self.has_header else 1
        end_row = ws.max_row
        if limit:
            end_row = min(end_row, start_row + limit - 1)

        if self.has_header:
            ws.cell(row=1, column=self.jab_amount_out_col, value="金额")
            ws.cell(row=1, column=self.jab_partner_out_col, value="对手方")

        updates = 0
        errors = {}
        for row in range(start_row, end_row + 1):
            raw_key = ws.cell(row=row, column=self.jab_key_col).value
            if raw_key is None or (isinstance(raw_key, str) and raw_key.strip() == ""):
                continue

            try:
                amount, partner = self.parse_jab_concat_key(raw_key)
            except ValueError as e:
                errors[row] = str(e)
                log.warning(f"行{row} 拼接索引拆分失败: {raw_key!r}, {e}")
                continue

            ws.cell(row=row, column=self.jab_amount_out_col, value=float(amount))
            ws.cell(row=row, column=self.jab_partner_out_col, value=partner)
            updates += 1
            log.info(f"行{row} 拆分索引: amount={amount} partner={partner}")

        wb.save(self.excel_path)
        wb.close()
        log.info(f"JAB 拼接索引拆分完成: updates={updates}, errors={len(errors)}")
        return {
            "updates": updates,
            "errors": errors,
            "amount_col": self.jab_amount_out_col,
            "partner_col": self.jab_partner_out_col,
        }

    def save_jab_results(self, row_values):
        if not row_values:
            return

        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb[self.sheet_my]

        for row, value in row_values.items():
            ws.cell(row=row, column=self.jab_result_col, value=value)
            log.info(f"行{row} 写入结果: {value}")

        wb.save(self.excel_path)
        wb.close()

    def _has_cell_value(self, value):
        return value is not None and str(value).strip() != ""

    def _looks_like_voucher(self, value):
        if not self._has_cell_value(value):
            return False
        if isinstance(value, int):
            return value > 0
        text = str(value).strip()
        if isinstance(value, float) and value.is_integer():
            text = str(int(value))
        return text.isdigit() and int(text) > 0
