import json
import openpyxl
import re
from pathlib import Path
from datetime import datetime
from decimal import Decimal, InvalidOperation
from collections import Counter
from core.logger import log
from core.utils import format_amount


CONCAT_KEY_RE = re.compile(
    r"^\s*([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s*(.+?)\s*$"
)


class DataHandler:
    def __init__(self, config):
        self.cfg = config
        self.excel_path = config["excel_path"]
        self.sheet_my = config["sheet_my"]
        self.sheet_finance = config["sheet_finance"]
        self.my_amount_col = config["my_amount_col"]
        self.my_voucher_col = config["my_voucher_col"]
        self.finance_amount_col = config["finance_amount_col"]
        self.has_header = config["has_header"]
        self.progress_file = Path("logs/progress.json")
        batch_cfg = config.get("jab_batch", {})
        self.jab_key_col = batch_cfg.get("key_col", self.my_amount_col)
        self.jab_result_col = batch_cfg.get("result_col", self.my_voucher_col)
        self.jab_amount_out_col = batch_cfg.get("amount_out_col", 3)
        self.jab_partner_out_col = batch_cfg.get("partner_out_col", 4)

    def load_my_data(self):
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb[self.sheet_my]

        start_row = 2 if self.has_header else 1
        data = []

        for row in range(start_row, ws.max_row + 1):
            amount = ws.cell(row=row, column=self.my_amount_col).value
            voucher = ws.cell(row=row, column=self.my_voucher_col).value

            # 跳过空值和空字符串
            if amount is None or (isinstance(amount, str) and amount.strip() == ""):
                continue

            try:
                data.append({
                    "row": row,
                    "amount": float(amount),
                    "voucher": voucher
                })
            except (ValueError, TypeError) as e:
                log.warning(f"行{row} 金额格式错误: {amount}, 跳过")
                continue

        wb.close()
        log.info(f"加载我的数据: {len(data)} 条")
        return data

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

    def load_finance_data(self):
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb[self.sheet_finance]

        start_row = 2 if self.has_header else 1
        amounts = []

        for row in range(start_row, ws.max_row + 1):
            val = ws.cell(row=row, column=self.finance_amount_col).value

            # 跳过空值和空字符串
            if val is None or (isinstance(val, str) and val.strip() == ""):
                continue

            try:
                amounts.append(float(val))
            except (ValueError, TypeError) as e:
                log.warning(f"财务数据行{row} 金额格式错误: {val}, 跳过")
                continue

        wb.close()
        log.info(f"加载财务数据: {len(amounts)} 条")
        return amounts

    def check_duplicates(self, my_data, finance_amounts):
        finance_count = Counter(finance_amounts)

        queue = []
        duplicates = []
        not_found = []

        for item in my_data:
            if item["voucher"]:
                log.info(f"行{item['row']} 金额{format_amount(item['amount'])} 已有凭证号，跳过")
                continue

            count = finance_count.get(item["amount"], 0)

            if count > 1:
                duplicates.append(item)
                log.warning(f"行{item['row']} 金额{format_amount(item['amount'])} 在财务数据中重复{count}次")
            elif count == 0:
                not_found.append(item)
                log.error(f"行{item['row']} 金额{format_amount(item['amount'])} 在财务数据中未找到")
            else:
                queue.append(item)

        log.info(f"可操作: {len(queue)} 条, 重复: {len(duplicates)} 条, 未找到: {len(not_found)} 条")
        return queue, duplicates, not_found

    def mark_issues(self, items, reason_prefix):
        if not items:
            return

        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb[self.sheet_my]

        for item in items:
            amount = format_amount(item["amount"])
            reason = f"{reason_prefix}-{amount}"
            ws.cell(row=item["row"], column=self.my_voucher_col, value=reason)
            log.info(f"行{item['row']} 标记为: {reason}")

        wb.save(self.excel_path)
        wb.close()
    def save_voucher(self, row, voucher_num, status="success"):
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb[self.sheet_my]

        if status == "success":
            ws.cell(row=row, column=self.my_voucher_col, value=voucher_num)
            log.info(f"行{row} 凭证号 {voucher_num} 已写入")
        else:
            ws.cell(row=row, column=self.my_voucher_col, value=f"失败-{voucher_num}")
            log.error(f"行{row} 标记失败: {voucher_num}")

        wb.save(self.excel_path)
        wb.close()


    def save_progress(self, current_index, total):
        progress = {
            "current": current_index,
            "total": total,
            "timestamp": datetime.now().isoformat()
        }
        self.progress_file.parent.mkdir(parents=True, exist_ok=True)
        with open(self.progress_file, "w", encoding="utf-8") as f:
            json.dump(progress, f, ensure_ascii=False, indent=2)

    def load_progress(self):
        if not self.progress_file.exists():
            return None

        try:
            with open(self.progress_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return None

    def clear_progress(self):
        if self.progress_file.exists():
            self.progress_file.unlink()
            log.info("进度文件已清除")

    def get_previous_voucher(self, current_row):
        """获取上一行的凭证号，用于对比"""
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb[self.sheet_my]

        previous_row = current_row - 1
        start_row = 2 if self.has_header else 1

        # 如果当前行是第一行数据，没有上一行
        if previous_row < start_row:
            wb.close()
            return None

        previous_voucher = ws.cell(row=previous_row, column=self.my_voucher_col).value
        wb.close()

        if previous_voucher and isinstance(previous_voucher, (int, float)):
            previous_voucher = str(int(previous_voucher))
        elif previous_voucher:
            previous_voucher = str(previous_voucher).strip()

        return previous_voucher
