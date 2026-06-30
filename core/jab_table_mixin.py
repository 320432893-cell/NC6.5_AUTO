# 职责：承载 JABOperator 的表格相关兼容方法代理
# 不做什么：不实现表格读取/选择算法，不直接访问 Excel/Sheet，不发送键盘鼠标输入
# 允许依赖层：core.jab_table_ops、core.jab_table_reader
# 谁不应该 import：业务填报工具和收款匹配模块不应直接 import

from core import jab_table_ops, jab_table_reader


class JABTableMixin:
    def select_table_rows(self, rows, selection_col=None, clear=True, timeout=None):
        return jab_table_ops.select_table_rows(
            self,
            rows,
            selection_col=selection_col,
            clear=clear,
            timeout=timeout,
        )

    def resolve_selection_col(self, selection_col):
        return jab_table_ops.resolve_selection_col(self, selection_col)

    def has_selection_api(self):
        return jab_table_ops.has_selection_api(self)

    def get_selected_child_indexes(self, vm_id, context, child_count):
        return jab_table_ops.get_selected_child_indexes(
            self, vm_id, context, child_count
        )

    def read_table_snapshot(
        self,
        amount_col=None,
        partner_col=None,
        voucher_col=None,
        extra_cols=None,
        limit=None,
        timeout=None,
    ):
        return jab_table_ops.read_table_snapshot(
            self,
            amount_col=amount_col,
            partner_col=partner_col,
            voucher_col=voucher_col,
            extra_cols=extra_cols,
            limit=limit,
            timeout=timeout,
        )

    def read_all_table_cells(
        self,
        max_rows=None,
        max_cols=None,
        timeout=None,
        scope_hwnd=None,
        exact_cols=None,
    ):
        return jab_table_ops.read_all_table_cells(
            self,
            max_rows=max_rows,
            max_cols=max_cols,
            timeout=timeout,
            scope_hwnd=scope_hwnd,
            exact_cols=exact_cols,
        )

    def read_table_summaries(
        self, min_rows=1, min_cols=None, scope_hwnd=None, exact_cols=None
    ):
        return jab_table_ops.read_table_summaries(
            self,
            min_rows=min_rows,
            min_cols=min_cols,
            scope_hwnd=scope_hwnd,
            exact_cols=exact_cols,
        )

    def read_table_selected_columns_from_context(
        self,
        table_index,
        table_context,
        vm_id,
        table_info,
        window_info,
        columns,
        max_rows=None,
    ):
        return jab_table_reader.read_table_selected_columns_from_context(
            self,
            table_index,
            table_context,
            vm_id,
            table_info,
            window_info,
            columns,
            max_rows=max_rows,
        )

    def read_all_table_selected_columns(
        self,
        columns,
        max_rows=None,
        min_rows=1,
        min_cols=None,
        scope_hwnd=None,
        exact_cols=None,
    ):
        return jab_table_ops.read_all_table_selected_columns(
            self,
            columns,
            max_rows=max_rows,
            min_rows=min_rows,
            min_cols=min_cols,
            scope_hwnd=scope_hwnd,
            exact_cols=exact_cols,
        )

    def read_window_table_cells(self, window_title, max_rows=None, max_cols=None):
        return jab_table_ops.read_window_table_cells(
            self, window_title, max_rows=max_rows, max_cols=max_cols
        )

    def read_window_table_counts(self, window_title):
        return jab_table_ops.read_window_table_counts(self, window_title)

    def read_table_cells_from_context(
        self,
        table_index,
        table_context,
        vm_id,
        table_info,
        window_info,
        max_rows=None,
        max_cols=None,
    ):
        return jab_table_reader.read_table_cells_from_context(
            self,
            table_index,
            table_context,
            vm_id,
            table_info,
            window_info,
            max_rows=max_rows,
            max_cols=max_cols,
        )

    def clear_table_cache(self, window_title=None):
        return jab_table_ops.clear_table_cache(self, window_title=window_title)

    def select_visible_table_rows(
        self,
        table_index,
        rows,
        window_title=None,
        selection_col=0,
        timeout=None,
    ):
        return jab_table_ops.select_visible_table_rows(
            self,
            table_index,
            rows,
            window_title=window_title,
            selection_col=selection_col,
            timeout=timeout,
        )

    def find_tables_once(self, scope_hwnd=None):
        return jab_table_ops.find_tables_once(self, scope_hwnd=scope_hwnd)

    def get_table_info(self, vm_id, context):
        return jab_table_reader.get_table_info(self, vm_id, context)

    def get_table_cell_text(self, vm_id, table_context, row, col):
        return jab_table_reader.get_table_cell_text(
            self,
            vm_id,
            table_context,
            row,
            col,
        )

    def get_table_cell_text_and_selection(self, vm_id, table_context, row, col):
        return jab_table_reader.get_table_cell_text_and_selection(
            self,
            vm_id,
            table_context,
            row,
            col,
        )
