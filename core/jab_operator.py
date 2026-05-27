import ctypes
from decimal import Decimal, InvalidOperation
import os
from ctypes import wintypes
import threading
import time

import pyautogui

from core.logger import log
from core.utils import check_abort
from tools.jab_probe import (
    AccessibleActions,
    AccessibleActionsToDo,
    AccessibleContextInfo,
    AccessibleTableCellInfo,
    AccessibleTableInfo,
    JOBJECT,
    configure_jab,
    enum_windows,
    load_access_bridge,
    run_windows_access_bridge,
)


class JABOperator:
    """Small Java Access Bridge wrapper for stable NC button/menu actions."""

    def __init__(self, config):
        jab_cfg = config.get("jab", {})
        self.dll_path = jab_cfg.get(
            "dll_path",
            r"C:\Users\Queclink\AppData\Local\UClient\share\java1.7.0_51-x64\bin\WindowsAccessBridge-64.dll",
        )
        self.startup_wait = jab_cfg.get("startup_wait", 2.0)
        self.search_timeout = jab_cfg.get("search_timeout", 5.0)
        self.max_depth = jab_cfg.get("max_depth", 25)
        self.max_children = jab_cfg.get("max_children", 1000)
        self.menu_wait = jab_cfg.get("menu_wait", 0.5)
        self.amount_col = jab_cfg.get("amount_col", 4)
        self.partner_col = jab_cfg.get("partner_col", 3)
        self.selection_col = jab_cfg.get("selection_col", 0)
        self.hide_blank_awt_windows_enabled = jab_cfg.get("hide_blank_awt_windows", True)

        self.dll = None
        self.loaded_path = None
        self.stop_pump = None
        self.pump_thread = None

    def ensure_started(self):
        if self.dll:
            return

        if os.name != "nt":
            raise RuntimeError("Java Access Bridge must run under Windows Python.")

        self.dll, self.loaded_path = load_access_bridge(self.dll_path)
        configure_jab(self.dll)

        if hasattr(self.dll, "initializeAccessBridge"):
            if not self.dll.initializeAccessBridge():
                raise RuntimeError("initializeAccessBridge returned false.")
        else:
            self.stop_pump = threading.Event()
            self.pump_thread = threading.Thread(
                target=run_windows_access_bridge,
                args=(self.dll, self.stop_pump),
                daemon=True,
            )
            self.pump_thread.start()

        time.sleep(self.startup_wait)
        self.hide_blank_awt_windows()
        log.info(f"JAB 已加载: {self.loaded_path}")

    def close(self):
        self.hide_blank_awt_windows()
        if self.stop_pump:
            self.stop_pump.set()
        if self.pump_thread:
            self.pump_thread.join(timeout=1)

    def __del__(self):
        try:
            self.close()
        except Exception:
            pass

    def hide_blank_awt_windows(self):
        """Hide small blank AWT helper windows sometimes left visible by JAB/Java."""
        if not self.hide_blank_awt_windows_enabled or os.name != "nt":
            return []
        if not hasattr(ctypes, "windll"):
            return []

        user32 = ctypes.windll.user32
        hidden = []

        class Rect(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def callback(hwnd, _lparam):
            if not user32.IsWindowVisible(hwnd):
                return True

            length = user32.GetWindowTextLengthW(hwnd)
            title = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, title, length + 1)

            class_name = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, class_name, 256)

            rect = Rect()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            width = rect.right - rect.left
            height = rect.bottom - rect.top

            if (
                class_name.value == "SunAwtWindow"
                and title.value == ""
                and 0 < width <= 250
                and 0 < height <= 250
            ):
                user32.ShowWindow(hwnd, 0)
                hidden.append({
                    "hwnd": int(hwnd),
                    "left": rect.left,
                    "top": rect.top,
                    "width": width,
                    "height": height,
                })

            return True

        user32.EnumWindows(enum_proc(callback), 0)
        if hidden:
            log.info(f"JAB 已隐藏空白 AWT 浮窗: {hidden}")
        return hidden

    def close_window_by_title(self, title, class_name=None, wait=None):
        if os.name != "nt":
            return False

        user32 = ctypes.windll.user32
        wm_close = 0x0010
        closed = []
        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def callback(hwnd, _lparam):
            if not user32.IsWindowVisible(hwnd):
                return True

            length = user32.GetWindowTextLengthW(hwnd)
            window_title = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, window_title, length + 1)

            window_class = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, window_class, 256)

            if window_title.value != title:
                return True
            if class_name and window_class.value != class_name:
                return True

            user32.PostMessageW(hwnd, wm_close, 0, 0)
            closed.append({
                "hwnd": int(hwnd),
                "title": window_title.value,
                "class": window_class.value,
            })
            return True

        user32.EnumWindows(enum_proc(callback), 0)
        if closed:
            log.info(f"已关闭窗口: {closed}")
            time.sleep(self.menu_wait if wait is None else wait)
            return True

        log.info(f"未找到需关闭窗口: title={title} class={class_name}")
        return False

    def activate_window_by_title(self, title, class_name=None, timeout=None):
        if os.name != "nt":
            return False

        deadline = time.time() + (timeout or self.search_timeout)
        while time.time() < deadline:
            hwnd = self.find_window_handle(title, class_name=class_name, visible_only=False)
            if hwnd:
                user32 = ctypes.windll.user32
                user32.ShowWindow(hwnd, 9)
                user32.SetForegroundWindow(hwnd)
                time.sleep(0.2)
                return True
            time.sleep(0.2)
        return False

    def wait_window_by_title(self, title, class_name=None, timeout=None):
        deadline = time.time() + (timeout or self.search_timeout)
        while time.time() < deadline:
            hwnd = self.find_window_handle(title, class_name=class_name, visible_only=True)
            if hwnd:
                return hwnd
            time.sleep(0.2)
        return None

    def window_exists(self, title, class_name=None):
        return bool(self.find_window_handle(title, class_name=class_name, visible_only=True))

    def find_window_handle(self, title, class_name=None, visible_only=True):
        if os.name != "nt":
            return None

        user32 = ctypes.windll.user32
        found = []
        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def callback(hwnd, _lparam):
            if visible_only and not user32.IsWindowVisible(hwnd):
                return True

            length = user32.GetWindowTextLengthW(hwnd)
            window_title = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, window_title, length + 1)

            window_class = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, window_class, 256)

            if window_title.value != title:
                return True
            if class_name and window_class.value != class_name:
                return True

            found.append(hwnd)
            return False

        user32.EnumWindows(enum_proc(callback), 0)
        return found[0] if found else None

    def press_key(self, key, wait=None):
        pyautogui.press(key)
        time.sleep(self.menu_wait if wait is None else wait)

    def do_generate_front(self):
        self.ensure_started()

        log.debug("JAB 点击生成按钮")
        if not self.click_control("生成", roles=("push button",), timeout=self.search_timeout):
            return False

        time.sleep(self.menu_wait)
        check_abort()

        log.debug("JAB 点击前台生成菜单项")
        return self.click_control("前台生成", roles=("menu item",), timeout=self.search_timeout)

    def run_named_steps(self, steps):
        self.ensure_started()
        for index, step in enumerate(steps, start=1):
            check_abort()
            name = step["name"] if isinstance(step, dict) else str(step)
            role = step.get("role") if isinstance(step, dict) else None
            wait = float(step.get("wait", self.menu_wait)) if isinstance(step, dict) else self.menu_wait
            timeout = float(step.get("timeout", self.search_timeout)) if isinstance(step, dict) else self.search_timeout
            action = step.get("action") if isinstance(step, dict) else None
            require_showing = step.get("require_showing", True) if isinstance(step, dict) else True
            roles = (role,) if role else ()

            log.info(f"JAB 执行步骤 {index}/{len(steps)}: name={name} role={role}")
            if not self.click_control(
                name,
                roles=roles,
                timeout=timeout,
                action_name=action,
                require_showing=require_showing,
            ):
                raise RuntimeError(f"JAB 步骤失败: {name}")
            time.sleep(wait)

        return True

    def click_control(self, name, roles=(), timeout=None, action_name=None, require_showing=False):
        self.ensure_started()
        context, vm_id, owned_contexts = self.find_context(
            name,
            roles=roles,
            timeout=timeout,
            require_showing=require_showing,
        )
        if not context:
            log.warning(f"JAB 未找到控件: {name}")
            return False

        try:
            return self.do_action(vm_id, context, action_name=action_name)
        finally:
            self.release_contexts(vm_id, owned_contexts)

    def find_amount_rows(self, amount, amount_col=None, timeout=None):
        self.ensure_started()
        table_context, vm_id, owned_contexts, table_info = self.find_main_table(timeout=timeout)
        if not table_context:
            log.warning("JAB 未找到业务表格")
            return []

        try:
            amount_col = self.resolve_amount_col(amount_col)
            target = self.normalize_amount(amount)
            if target is None:
                raise ValueError(f"金额格式无法识别: {amount!r}")

            matched_rows = []
            for row in range(table_info.rowCount):
                check_abort()
                cell_text = self.get_table_cell_text(vm_id, table_context, row, amount_col)
                if self.normalize_amount(cell_text) == target:
                    matched_rows.append(row)

            log.info(f"JAB 表格金额匹配: amount={amount} col={amount_col} rows={matched_rows}")
            return matched_rows
        finally:
            self.release_contexts(vm_id, owned_contexts)

    def select_amount_row(self, amount, amount_col=None, selection_col=None, timeout=None):
        matched_rows = self.find_amount_rows(amount, amount_col=amount_col, timeout=timeout)
        if len(matched_rows) != 1:
            log.warning(f"JAB 金额匹配行数不是 1: amount={amount} rows={matched_rows}")
            return False
        return self.select_table_rows(matched_rows, selection_col=selection_col, timeout=timeout)

    def find_amount_partner_rows(
        self,
        amount,
        partner_name,
        amount_col=None,
        partner_col=None,
        timeout=None,
        match_mode="exact",
    ):
        self.ensure_started()
        table_context, vm_id, owned_contexts, table_info = self.find_main_table(timeout=timeout)
        if not table_context:
            log.warning("JAB 未找到业务表格")
            return []

        try:
            amount_col = self.resolve_amount_col(amount_col)
            partner_col = self.resolve_partner_col(partner_col)
            target_amount = self.normalize_amount(amount)
            target_partner = self.normalize_text(partner_name)
            if target_amount is None:
                raise ValueError(f"金额格式无法识别: {amount!r}")
            if not target_partner:
                raise ValueError("业务关联方名称为空")

            matched_rows = []
            amount_rows = []
            for row in range(table_info.rowCount):
                check_abort()
                amount_text = self.get_table_cell_text(vm_id, table_context, row, amount_col)
                if self.normalize_amount(amount_text) != target_amount:
                    continue

                amount_rows.append(row)
                partner_text = self.get_table_cell_text(vm_id, table_context, row, partner_col)
                if self.text_matches(partner_text, target_partner, match_mode):
                    matched_rows.append(row)

            log.info(
                "JAB 金额+关联方匹配: "
                f"amount={amount} amount_col={amount_col} amount_rows={amount_rows} "
                f"partner={partner_name} partner_col={partner_col} rows={matched_rows}"
            )
            return matched_rows
        finally:
            self.release_contexts(vm_id, owned_contexts)

    def select_amount_partner_row(
        self,
        amount,
        partner_name,
        amount_col=None,
        partner_col=None,
        selection_col=None,
        timeout=None,
        match_mode="exact",
    ):
        matched_rows = self.find_amount_partner_rows(
            amount,
            partner_name,
            amount_col=amount_col,
            partner_col=partner_col,
            timeout=timeout,
            match_mode=match_mode,
        )
        if len(matched_rows) != 1:
            log.warning(
                f"JAB 金额+关联方匹配行数不是 1: "
                f"amount={amount} partner={partner_name} rows={matched_rows}"
            )
            return False
        return self.select_table_rows(matched_rows, selection_col=selection_col, timeout=timeout)

    def select_table_rows(self, rows, selection_col=None, clear=True, timeout=None):
        self.ensure_started()
        table_context, vm_id, owned_contexts, table_info = self.find_main_table(timeout=timeout)
        if not table_context:
            log.warning("JAB 未找到业务表格")
            return False

        try:
            selection_col = self.resolve_selection_col(selection_col)
            if not self.has_selection_api():
                log.warning("当前 JAB DLL 不支持 selection API")
                return False

            if clear:
                self.dll.clearAccessibleSelectionFromContext(vm_id, table_context)

            for row in rows:
                if row < 0 or row >= table_info.rowCount:
                    raise ValueError(f"行号越界: {row}, rowCount={table_info.rowCount}")
                if selection_col < 0 or selection_col >= table_info.columnCount:
                    raise ValueError(f"列号越界: {selection_col}, colCount={table_info.columnCount}")
                child_index = row * table_info.columnCount + selection_col
                self.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)

            selected_indexes = self.get_selected_child_indexes(
                vm_id,
                table_context,
                table_info.rowCount * table_info.columnCount,
            )
            expected_indexes = [
                row * table_info.columnCount + selection_col
                for row in rows
            ]
            ok = all(index in selected_indexes for index in expected_indexes)
            log.info(
                f"JAB 选中表格行: rows={rows} col={selection_col} "
                f"expected={expected_indexes} selected={selected_indexes[:20]}"
            )
            return ok
        finally:
            self.release_contexts(vm_id, owned_contexts)

    def clear_table_selection(self, timeout=None):
        self.ensure_started()
        table_context, vm_id, owned_contexts, _ = self.find_main_table(timeout=timeout)
        if not table_context:
            log.warning("JAB 未找到业务表格")
            return False

        try:
            if not hasattr(self.dll, "clearAccessibleSelectionFromContext"):
                return False
            self.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
            return True
        finally:
            self.release_contexts(vm_id, owned_contexts)

    def resolve_selection_col(self, selection_col):
        if selection_col is not None:
            return selection_col
        return self.selection_col

    def has_selection_api(self):
        return (
            hasattr(self.dll, "clearAccessibleSelectionFromContext")
            and hasattr(self.dll, "addAccessibleSelectionFromContext")
            and hasattr(self.dll, "isAccessibleChildSelectedFromContext")
        )

    def get_selected_child_indexes(self, vm_id, context, child_count):
        if not hasattr(self.dll, "isAccessibleChildSelectedFromContext"):
            return []
        selected = []
        for index in range(child_count):
            if self.dll.isAccessibleChildSelectedFromContext(vm_id, context, index):
                selected.append(index)
        return selected

    def read_amount_column(self, amount_col=None, limit=None, timeout=None):
        self.ensure_started()
        table_context, vm_id, owned_contexts, table_info = self.find_main_table(timeout=timeout)
        if not table_context:
            log.warning("JAB 未找到业务表格")
            return []

        try:
            amount_col = self.resolve_amount_col(amount_col)
            row_count = table_info.rowCount if limit is None else min(table_info.rowCount, limit)
            values = []
            for row in range(row_count):
                check_abort()
                values.append(self.get_table_cell_text(vm_id, table_context, row, amount_col))
            return values
        finally:
            self.release_contexts(vm_id, owned_contexts)

    def read_table_snapshot(
        self,
        amount_col=None,
        partner_col=None,
        voucher_col=None,
        extra_cols=None,
        limit=None,
        timeout=None,
    ):
        self.ensure_started()
        table_context, vm_id, owned_contexts, table_info = self.find_main_table(timeout=timeout)
        if not table_context:
            log.warning("JAB 未找到业务表格")
            return []

        try:
            amount_col = self.resolve_amount_col(amount_col)
            partner_col = self.resolve_partner_col(partner_col)
            extra_cols = extra_cols or []
            row_count = table_info.rowCount if limit is None else min(table_info.rowCount, limit)
            rows = []
            for row in range(row_count):
                check_abort()
                amount_text = self.get_table_cell_text(vm_id, table_context, row, amount_col)
                partner_text = self.get_table_cell_text(vm_id, table_context, row, partner_col)
                item = {
                    "row_index": row,
                    "amount_text": amount_text,
                    "amount": self.normalize_amount(amount_text),
                    "partner_text": partner_text,
                    "partner": self.normalize_text(partner_text),
                }
                if voucher_col is not None:
                    item["voucher_text"] = self.get_table_cell_text(
                        vm_id,
                        table_context,
                        row,
                        voucher_col,
                    ).strip()
                if extra_cols:
                    item["extra_text"] = {
                        col: self.get_table_cell_text(
                            vm_id,
                            table_context,
                            row,
                            col,
                        ).strip()
                        for col in extra_cols
                    }
                rows.append(item)

            log.info(
                "JAB 读取表格快照: "
                f"rows={len(rows)} amount_col={amount_col} partner_col={partner_col} "
                f"voucher_col={voucher_col}"
            )
            return rows
        finally:
            self.release_contexts(vm_id, owned_contexts)

    def read_all_table_cells(self, max_rows=None, max_cols=None, timeout=None):
        self.ensure_started()
        tables = self.find_tables_once()
        result = []

        for table_index, (table_context, vm_id, owned_contexts, table_info, window_info) in enumerate(tables):
            try:
                row_count = table_info.rowCount
                col_count = table_info.columnCount
                row_limit = row_count if max_rows is None else min(row_count, max_rows)
                col_limit = col_count if max_cols is None else min(col_count, max_cols)
                rows = []

                for row in range(row_limit):
                    check_abort()
                    cells = []
                    selected = False
                    for col in range(col_limit):
                        text, is_selected = self.get_table_cell_text_and_selection(
                            vm_id,
                            table_context,
                            row,
                            col,
                        )
                        cells.append(text)
                        selected = selected or is_selected
                    rows.append({
                        "row_index": row,
                        "cells": cells,
                        "selected": selected,
                    })

                result.append({
                    "table_index": table_index,
                    "window_title": window_info["title"],
                    "window_class": window_info["class_name"],
                    "window_visible": window_info["visible"],
                    "row_count": row_count,
                    "col_count": col_count,
                    "rows": rows,
                })
            finally:
                self.release_contexts(vm_id, owned_contexts)

        log.debug(f"JAB 读取所有表格: count={len(result)}")
        return result

    def read_window_table_cells(self, window_title, max_rows=None, max_cols=None):
        tables = self.read_all_table_cells(max_rows=max_rows, max_cols=max_cols)
        return [table for table in tables if table.get("window_title") == window_title]

    def select_visible_table_rows(
        self,
        table_index,
        rows,
        window_title=None,
        selection_col=0,
        timeout=None,
    ):
        self.ensure_started()
        deadline = time.time() + (timeout or self.search_timeout)

        while time.time() < deadline:
            check_abort()
            ok = self.select_visible_table_rows_once(
                table_index,
                rows,
                window_title=window_title,
                selection_col=selection_col,
            )
            if ok:
                return True
            time.sleep(0.2)
        return False

    def select_visible_table_rows_once(
        self,
        table_index,
        rows,
        window_title=None,
        selection_col=0,
    ):
        tables = self.find_tables_once()
        for current_index, (table_context, vm_id, owned_contexts, table_info, window_info) in enumerate(tables):
            try:
                if current_index != table_index:
                    continue
                if window_title is not None and window_info.get("title") != window_title:
                    return False
                if not self.has_selection_api():
                    log.warning("当前 JAB DLL 不支持 selection API")
                    return False
                if selection_col < 0 or selection_col >= table_info.columnCount:
                    raise ValueError(f"列号越界: {selection_col}, colCount={table_info.columnCount}")

                self.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
                expected_indexes = []
                for row in rows:
                    if row < 0 or row >= table_info.rowCount:
                        raise ValueError(f"行号越界: {row}, rowCount={table_info.rowCount}")
                    child_index = row * table_info.columnCount + selection_col
                    expected_indexes.append(child_index)
                    self.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)

                selected_indexes = self.get_selected_child_indexes(
                    vm_id,
                    table_context,
                    table_info.rowCount * table_info.columnCount,
                )
                ok = all(index in selected_indexes for index in expected_indexes)
                log.info(
                    f"JAB 选中可见表格行: table={table_index} window={window_title} "
                    f"rows={rows} expected={expected_indexes} selected={selected_indexes[:40]}"
                )
                return ok
            finally:
                self.release_contexts(vm_id, owned_contexts)

        return False

    def wait_for_record_visible(
        self,
        amount,
        partner_name,
        timeout=None,
        selected_first=True,
        max_rows=200,
        max_cols=50,
        window_title=None,
    ):
        deadline = time.time() + (timeout or self.search_timeout)
        while time.time() < deadline:
            check_abort()
            found = self.find_record_in_visible_tables(
                amount,
                partner_name,
                selected_first=selected_first,
                max_rows=max_rows,
                max_cols=max_cols,
                window_title=window_title,
            )
            if found:
                return found
            time.sleep(0.2)
        return None

    def find_record_in_visible_tables(
        self,
        amount,
        partner_name,
        selected_first=True,
        max_rows=200,
        max_cols=50,
        window_title=None,
    ):
        target_amount = self.normalize_amount(amount)
        target_partner = self.normalize_text(partner_name)
        if target_amount is None or not target_partner:
            return None

        tables = self.read_all_table_cells(max_rows=max_rows, max_cols=max_cols)
        candidates = []
        fallback = []

        for table in tables:
            if window_title is not None and table.get("window_title") != window_title:
                continue
            for row in table["rows"]:
                normalized_cells = [self.normalize_text(cell) for cell in row["cells"]]
                row_text = "".join(normalized_cells)
                amount_match = any(
                    self.normalize_amount(cell) == target_amount
                    for cell in row["cells"]
                )
                partner_match = target_partner in row_text
                if amount_match and partner_match:
                    item = {
                        "table_index": table["table_index"],
                        "window_title": table.get("window_title"),
                        "window_class": table.get("window_class"),
                        "table_rows": table["row_count"],
                        "table_cols": table["col_count"],
                        "row_index": row["row_index"],
                        "selected": row["selected"],
                        "cells": row["cells"],
                    }
                    if row["selected"]:
                        candidates.append(item)
                    else:
                        fallback.append(item)

        if selected_first and candidates:
            log.debug(f"JAB 找到选中当前记录: {candidates[0]}")
            return candidates[0]
        if fallback:
            log.debug(f"JAB 找到可见记录: {fallback[0]}")
            return fallback[0]
        return None

    def select_record_in_visible_tables(
        self,
        amount,
        partner_name,
        window_title=None,
        selection_col=0,
        timeout=None,
        max_rows=200,
        max_cols=50,
    ):
        self.ensure_started()
        deadline = time.time() + (timeout or self.search_timeout)

        while time.time() < deadline:
            check_abort()
            result = self.select_record_in_visible_tables_once(
                amount,
                partner_name,
                window_title=window_title,
                selection_col=selection_col,
                max_rows=max_rows,
                max_cols=max_cols,
            )
            if result:
                return result
            time.sleep(0.2)

        return None

    def select_record_in_visible_tables_once(
        self,
        amount,
        partner_name,
        window_title=None,
        selection_col=0,
        max_rows=200,
        max_cols=50,
    ):
        target_amount = self.normalize_amount(amount)
        target_partner = self.normalize_text(partner_name)
        if target_amount is None or not target_partner:
            return None

        tables = self.find_tables_once()
        for table_index, (table_context, vm_id, owned_contexts, table_info, window_info) in enumerate(tables):
            try:
                if window_title is not None and window_info.get("title") != window_title:
                    continue

                row_limit = min(table_info.rowCount, max_rows)
                col_limit = min(table_info.columnCount, max_cols)
                for row in range(row_limit):
                    cells = []
                    amount_match = False
                    for col in range(col_limit):
                        text = self.get_table_cell_text(vm_id, table_context, row, col)
                        cells.append(text)
                        if self.normalize_amount(text) == target_amount:
                            amount_match = True

                    row_text = "".join(self.normalize_text(cell) for cell in cells)
                    if not amount_match or target_partner not in row_text:
                        continue

                    if not self.has_selection_api():
                        log.warning("当前 JAB DLL 不支持 selection API")
                        return None
                    if selection_col < 0 or selection_col >= table_info.columnCount:
                        raise ValueError(
                            f"列号越界: {selection_col}, colCount={table_info.columnCount}"
                        )

                    self.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
                    child_index = row * table_info.columnCount + selection_col
                    self.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)
                    selected_indexes = self.get_selected_child_indexes(
                        vm_id,
                        table_context,
                        table_info.rowCount * table_info.columnCount,
                    )
                    result = {
                        "ok": child_index in selected_indexes,
                        "table_index": table_index,
                        "window_title": window_info.get("title"),
                        "window_class": window_info.get("class_name"),
                        "table_rows": table_info.rowCount,
                        "table_cols": table_info.columnCount,
                        "row_index": row,
                        "child_index": child_index,
                        "selected_indexes": selected_indexes[:20],
                        "cells": cells,
                    }
                    log.info(f"JAB 选择可见表格记录: {result}")
                    return result
            finally:
                self.release_contexts(vm_id, owned_contexts)

        return None

    def click_save(self, timeout=None):
        self.ensure_started()
        return self.click_control(
            "保存(Ctrl+S)",
            roles=("push button",),
            timeout=timeout or self.search_timeout,
        )

    def wait_for_control(self, name, roles=(), timeout=None, require_showing=False):
        self.ensure_started()
        context, vm_id, owned_contexts = self.find_context(
            name,
            roles=roles,
            timeout=timeout,
            require_showing=require_showing,
        )
        if not context:
            return False
        self.release_contexts(vm_id, owned_contexts)
        return True

    def wait_save_success(self, timeout=None):
        return self.wait_for_control("保存成功", timeout=timeout or self.search_timeout)

    def resolve_amount_col(self, amount_col):
        if amount_col is not None:
            return amount_col
        return self.amount_col

    def resolve_partner_col(self, partner_col):
        if partner_col is not None:
            return partner_col
        return self.partner_col

    def find_main_table(self, timeout=None):
        deadline = time.time() + (timeout or self.search_timeout)

        while time.time() < deadline:
            check_abort()
            tables = self.find_tables_once()
            if tables:
                tables.sort(key=lambda item: item[3].rowCount * item[3].columnCount, reverse=True)
                table_context, vm_id, owned_contexts, table_info, _window_info = tables[0]

                log.debug(
                    f"JAB 找到业务表格: rows={table_info.rowCount} cols={table_info.columnCount}"
                )
                return table_context, vm_id, owned_contexts, table_info
            time.sleep(0.2)

        return None, None, [], None

    def find_tables_once(self):
        tables = []
        windows = enum_windows(include_children=True)

        for hwnd, title, class_name, pid, visible in windows:
            if not self.dll.isJavaWindow(hwnd):
                continue

            vm_id = ctypes.c_long()
            root_context = JOBJECT()
            if not self.dll.getAccessibleContextFromHWND(
                hwnd,
                ctypes.byref(vm_id),
                ctypes.byref(root_context),
            ):
                continue

            tables.extend(
                self.find_tables_in_tree(
                    vm_id.value,
                    root_context.value,
                    depth=0,
                    owned_path=[],
                    window_info={
                        "hwnd": int(hwnd),
                        "title": title,
                        "class_name": class_name,
                        "pid": pid,
                        "visible": visible,
                    },
                )
            )

        return tables

    def find_tables_in_tree(self, vm_id, context, depth, owned_path, window_info=None):
        info = self.get_context_info(vm_id, context)
        if not info:
            return []

        role = (info.role_en_US.strip() or info.role.strip()).lower()
        if role == "table":
            table_info = self.get_table_info(vm_id, context)
            if table_info and table_info.rowCount > 0 and table_info.columnCount > 0:
                return [(context, vm_id, list(owned_path), table_info, window_info or {})]
            return []

        if depth >= self.max_depth:
            return []

        tables = []
        child_count = min(info.childrenCount, self.max_children)
        for index in range(child_count):
            child = self.dll.getAccessibleChildFromContext(vm_id, context, index)
            if not child:
                continue

            child_tables = self.find_tables_in_tree(
                vm_id,
                child,
                depth + 1,
                owned_path + [child],
                window_info=window_info,
            )
            if child_tables:
                tables.extend(child_tables)
            else:
                self.release_contexts(vm_id, [child])

        return tables

    def get_table_info(self, vm_id, context):
        if not hasattr(self.dll, "getAccessibleTableInfo"):
            return None

        table_info = AccessibleTableInfo()
        if not self.dll.getAccessibleTableInfo(vm_id, context, ctypes.byref(table_info)):
            return None
        return table_info

    def get_table_cell_text(self, vm_id, table_context, row, col):
        if not hasattr(self.dll, "getAccessibleTableCellInfo"):
            return ""

        cell_info = AccessibleTableCellInfo()
        ok = self.dll.getAccessibleTableCellInfo(
            vm_id,
            table_context,
            row,
            col,
            ctypes.byref(cell_info),
        )
        if not ok or not cell_info.accessibleContext:
            return ""

        info = self.get_context_info(vm_id, cell_info.accessibleContext)
        if not info:
            return ""

        return info.name.strip() or info.description.strip()

    def get_table_cell_text_and_selection(self, vm_id, table_context, row, col):
        if not hasattr(self.dll, "getAccessibleTableCellInfo"):
            return "", False

        cell_info = AccessibleTableCellInfo()
        ok = self.dll.getAccessibleTableCellInfo(
            vm_id,
            table_context,
            row,
            col,
            ctypes.byref(cell_info),
        )
        if not ok or not cell_info.accessibleContext:
            return "", bool(cell_info.isSelected)

        info = self.get_context_info(vm_id, cell_info.accessibleContext)
        if not info:
            return "", bool(cell_info.isSelected)

        return info.name.strip() or info.description.strip(), bool(cell_info.isSelected)

    def normalize_amount(self, value):
        if value is None:
            return None
        text = str(value).strip().replace(",", "")
        if not text:
            return None
        try:
            return Decimal(text).quantize(Decimal("0.01"))
        except (InvalidOperation, ValueError):
            return None

    def normalize_text(self, value):
        if value is None:
            return ""
        return "".join(str(value).split())

    def text_matches(self, value, target, match_mode):
        text = self.normalize_text(value)
        if match_mode == "contains":
            return target in text
        return text == target

    def find_context(self, name, roles=(), timeout=None, require_showing=False):
        deadline = time.time() + (timeout or self.search_timeout)
        normalized_roles = {role.lower() for role in roles}

        while time.time() < deadline:
            check_abort()
            result = self.find_context_once(
                name,
                normalized_roles,
                require_showing=require_showing,
            )
            if result[0]:
                return result
            time.sleep(0.2)

        return None, None, []

    def find_context_once(self, name, normalized_roles, require_showing=False):
        windows = enum_windows(include_children=True)

        for hwnd, title, class_name, pid, visible in windows:
            if not self.dll.isJavaWindow(hwnd):
                continue

            vm_id = ctypes.c_long()
            root_context = JOBJECT()
            if not self.dll.getAccessibleContextFromHWND(
                hwnd,
                ctypes.byref(vm_id),
                ctypes.byref(root_context),
            ):
                continue

            context, owned_contexts = self.find_in_tree(
                vm_id.value,
                root_context.value,
                name,
                normalized_roles,
                require_showing,
                depth=0,
                owned_path=[],
            )
            if context:
                log.debug(
                    f"JAB 找到控件 {name}: hwnd={int(hwnd)} pid={pid} "
                    f"class={class_name!r} title={title!r} visible={visible}"
                )
                return context, vm_id.value, owned_contexts

        return None, None, []

    def find_in_tree(self, vm_id, context, name, normalized_roles, require_showing, depth, owned_path):
        info = self.get_context_info(vm_id, context)
        if not info:
            return None, []

        role = (info.role_en_US.strip() or info.role.strip()).lower()
        control_name = info.name.strip()
        desc = info.description.strip()
        states = (info.states_en_US.strip() or info.states.strip()).lower()

        if self.matches_control(
            control_name,
            desc,
            role,
            states,
            name,
            normalized_roles,
            require_showing,
        ):
            return context, list(owned_path)

        if depth >= self.max_depth:
            return None, []
        if role == "table" and "table" not in normalized_roles:
            return None, []

        child_count = min(info.childrenCount, self.max_children)
        for index in range(child_count):
            child = self.dll.getAccessibleChildFromContext(vm_id, context, index)
            if not child:
                continue

            found, found_owned = self.find_in_tree(
                vm_id,
                child,
                name,
                normalized_roles,
                require_showing,
                depth + 1,
                owned_path + [child],
            )
            if found:
                return found, found_owned

            self.release_contexts(vm_id, [child])

        return None, []

    def get_context_info(self, vm_id, context):
        info = AccessibleContextInfo()
        if not self.dll.getAccessibleContextInfo(vm_id, context, ctypes.byref(info)):
            return None
        return info

    def matches_control(
        self,
        control_name,
        desc,
        role,
        states,
        expected_name,
        normalized_roles,
        require_showing,
    ):
        if normalized_roles and role not in normalized_roles:
            return False
        if require_showing and ("visible" not in states or "showing" not in states):
            return False
        return control_name == expected_name or desc == expected_name

    def get_action_names(self, vm_id, context):
        actions = AccessibleActions()
        if not self.dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
            return []
        return [actions.actionInfo[index].name.strip() for index in range(actions.actionsCount)]

    def do_action(self, vm_id, context, action_name=None):
        if not hasattr(self.dll, "getAccessibleActions") or not hasattr(self.dll, "doAccessibleActions"):
            log.warning("当前 JAB DLL 不支持 AccessibleActions")
            return False

        action_names = self.get_action_names(vm_id, context)
        if not action_names:
            log.warning("JAB 控件没有可执行动作")
            return False

        chosen_action = action_name or action_names[0]
        if chosen_action not in action_names:
            log.warning(f"JAB 控件不支持动作 {chosen_action!r}，可用动作: {action_names}")
            return False

        todo = AccessibleActionsToDo()
        todo.actionsCount = 1
        todo.actions[0].name = chosen_action
        failure = ctypes.c_int(-1)
        ok = self.dll.doAccessibleActions(
            vm_id,
            context,
            ctypes.byref(todo),
            ctypes.byref(failure),
        )
        log.debug(f"JAB 执行动作 {chosen_action!r}: ok={bool(ok)} failure={failure.value}")
        return bool(ok)

    def release_contexts(self, vm_id, contexts):
        if not vm_id or not hasattr(self.dll, "releaseJavaObject"):
            return
        for context in reversed(contexts):
            try:
                self.dll.releaseJavaObject(vm_id, context)
            except Exception:
                pass
