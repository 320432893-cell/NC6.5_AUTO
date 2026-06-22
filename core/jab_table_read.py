# 职责：JAB 表格单元格/快照读取、金额/关联方列解析、可见表记录匹配
# 不做什么：不读取单元格底层 ctypes 结构，不做 context path 动作，不发送全局键盘输入
# 允许依赖层：标准库 ctypes/os/time、JABOperator 暴露能力、tools.jab_probe、core.jab_table_* 同层
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import

from core.logger import log
from core.utils import check_abort

from core.jab_table_find import find_main_table, find_tables_once


def read_table_snapshot(
    jab,
    amount_col=None,
    partner_col=None,
    voucher_col=None,
    extra_cols=None,
    limit=None,
    timeout=None,
):
    jab.ensure_started()
    table_context, vm_id, owned_contexts, table_info = find_main_table(
        jab, timeout=timeout
    )
    if not table_context:
        log.warning("JAB 未找到业务表格")
        return []

    try:
        amount_col = resolve_amount_col(jab, amount_col)
        partner_col = resolve_partner_col(jab, partner_col)
        extra_cols = extra_cols or []
        row_count = (
            table_info.rowCount if limit is None else min(table_info.rowCount, limit)
        )
        rows = []
        for row in range(row_count):
            check_abort()
            amount_text = jab.get_table_cell_text(vm_id, table_context, row, amount_col)
            partner_text = jab.get_table_cell_text(
                vm_id, table_context, row, partner_col
            )
            item = {
                "row_index": row,
                "amount_text": amount_text,
                "amount": jab.normalize_amount(amount_text),
                "partner_text": partner_text,
                "partner": jab.normalize_text(partner_text),
            }
            if voucher_col is not None:
                item["voucher_text"] = jab.get_table_cell_text(
                    vm_id, table_context, row, voucher_col
                ).strip()
            if extra_cols:
                item["extra_text"] = {
                    col: jab.get_table_cell_text(vm_id, table_context, row, col).strip()
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
        jab.release_contexts(vm_id, owned_contexts)



def read_all_table_cells(
    jab, max_rows=None, max_cols=None, timeout=None, scope_hwnd=None, exact_cols=None
):
    jab.ensure_started()
    tables = find_tables_once(jab, scope_hwnd=scope_hwnd)
    result = []

    for table_index, (
        table_context,
        vm_id,
        owned_contexts,
        table_info,
        window_info,
    ) in enumerate(tables):
        try:
            if exact_cols is not None and table_info.columnCount != int(exact_cols):
                continue
            result.append(
                jab.read_table_cells_from_context(
                    table_index,
                    table_context,
                    vm_id,
                    table_info,
                    window_info,
                    max_rows=max_rows,
                    max_cols=max_cols,
                )
            )
        finally:
            jab.release_contexts(vm_id, owned_contexts)

    log.debug(f"JAB 读取所有表格: count={len(result)}")
    return result



def read_table_summaries(
    jab, min_rows=1, min_cols=None, scope_hwnd=None, exact_cols=None
):
    jab.ensure_started()
    tables = find_tables_once(jab, scope_hwnd=scope_hwnd)
    result = []

    for table_index, (
        _table_context,
        vm_id,
        owned_contexts,
        table_info,
        window_info,
    ) in enumerate(tables):
        try:
            if table_info.rowCount < min_rows:
                continue
            if min_cols is not None and table_info.columnCount < min_cols:
                continue
            if exact_cols is not None and table_info.columnCount != int(exact_cols):
                continue
            result.append(
                {
                    "table_index": table_index,
                    "window_title": window_info.get("title"),
                    "window_class": window_info.get("class_name"),
                    "window_visible": window_info.get("visible"),
                    "row_count": table_info.rowCount,
                    "col_count": table_info.columnCount,
                }
            )
        finally:
            jab.release_contexts(vm_id, owned_contexts)

    log.debug(f"JAB 读取表格摘要: count={len(result)}")
    return result



def read_all_table_selected_columns(
    jab,
    columns,
    max_rows=None,
    min_rows=1,
    min_cols=None,
    scope_hwnd=None,
    exact_cols=None,
):
    jab.ensure_started()
    selected_columns = sorted({int(column) for column in columns})
    if not selected_columns:
        return []
    min_col_count = max(selected_columns) + 1 if min_cols is None else int(min_cols)
    tables = find_tables_once(jab, scope_hwnd=scope_hwnd)
    result = []

    for table_index, (
        table_context,
        vm_id,
        owned_contexts,
        table_info,
        window_info,
    ) in enumerate(tables):
        try:
            if table_info.rowCount < min_rows or table_info.columnCount < min_col_count:
                continue
            if exact_cols is not None and table_info.columnCount != int(exact_cols):
                continue
            result.append(
                jab.read_table_selected_columns_from_context(
                    table_index,
                    table_context,
                    vm_id,
                    table_info,
                    window_info,
                    selected_columns,
                    max_rows=max_rows,
                )
            )
        finally:
            jab.release_contexts(vm_id, owned_contexts)

    log.debug(
        "JAB 读取表格指定列: "
        f"count={len(result)} columns={selected_columns} max_rows={max_rows}"
    )
    return result



def resolve_amount_col(jab, amount_col):
    if amount_col is not None:
        return amount_col
    return jab.amount_col



def resolve_partner_col(jab, partner_col):
    if partner_col is not None:
        return partner_col
    return jab.partner_col
