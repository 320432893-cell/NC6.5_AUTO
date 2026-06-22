# 职责：读取 JAB table 元信息、单元格文本，并投影为调用方使用的表格 dict
# 不做什么：不枚举窗口，不缓存 context，不选择行，不匹配业务金额/关联方
# 允许依赖层：标准库 ctypes、JABOperator 暴露的 dll/get_context_info 能力、tools.jab_probe 表格结构
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import

import ctypes

from core.jab_context_tree import release_contexts
from core.utils import check_abort
from tools.jab_probe import AccessibleTableCellInfo, AccessibleTableInfo


def read_table_selected_columns_from_context(
    jab,
    table_index,
    table_context,
    vm_id,
    table_info,
    window_info,
    columns,
    max_rows=None,
):
    row_count = table_info.rowCount
    col_count = table_info.columnCount
    selected_columns = sorted(
        {int(column) for column in columns if 0 <= int(column) < col_count}
    )
    row_limit = row_count if max_rows is None else min(row_count, max_rows)
    cell_count = max(selected_columns, default=-1) + 1
    rows = []

    for row in range(row_limit):
        check_abort()
        cells = [""] * cell_count
        selected = False
        for col in selected_columns:
            text, is_selected = get_table_cell_text_and_selection(
                jab,
                vm_id,
                table_context,
                row,
                col,
            )
            cells[col] = text
            selected = selected or is_selected
        rows.append(
            {
                "row_index": row,
                "cells": cells,
                "selected": selected,
            }
        )

    return {
        "table_index": table_index,
        "window_title": window_info.get("title"),
        "window_class": window_info.get("class_name"),
        "window_visible": window_info.get("visible"),
        "row_count": row_count,
        "col_count": col_count,
        "read_columns": selected_columns,
        "rows": rows,
    }


def read_table_cells_from_context(
    jab,
    table_index,
    table_context,
    vm_id,
    table_info,
    window_info,
    max_rows=None,
    max_cols=None,
):
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
            text, is_selected = get_table_cell_text_and_selection(
                jab,
                vm_id,
                table_context,
                row,
                col,
            )
            cells.append(text)
            selected = selected or is_selected
        rows.append(
            {
                "row_index": row,
                "cells": cells,
                "selected": selected,
            }
        )

    return {
        "table_index": table_index,
        "window_title": window_info.get("title"),
        "window_class": window_info.get("class_name"),
        "window_visible": window_info.get("visible"),
        "row_count": row_count,
        "col_count": col_count,
        "rows": rows,
    }


def get_table_info(jab, vm_id, context):
    if not hasattr(jab.dll, "getAccessibleTableInfo"):
        return None

    table_info = AccessibleTableInfo()
    if not jab.dll.getAccessibleTableInfo(vm_id, context, ctypes.byref(table_info)):
        return None
    return table_info


def get_table_cell_text(jab, vm_id, table_context, row, col):
    if not hasattr(jab.dll, "getAccessibleTableCellInfo"):
        return ""

    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    if not ok or not cell_info.accessibleContext:
        return ""

    info = jab.get_context_info(vm_id, cell_info.accessibleContext)
    # 读完即释放该单元格 JAB 句柄,否则整批制单(50-100 单不重启 NC)会在
    # NC 客户端侧累积数十万未回收句柄拖垮读控件;释放门面见 jab_context_tree
    release_contexts(jab, vm_id, [cell_info.accessibleContext])
    if not info:
        return ""

    return info.name.strip() or info.description.strip()


def get_table_cell_text_and_selection(jab, vm_id, table_context, row, col):
    if not hasattr(jab.dll, "getAccessibleTableCellInfo"):
        return "", False

    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    if not ok or not cell_info.accessibleContext:
        return "", bool(cell_info.isSelected)

    info = jab.get_context_info(vm_id, cell_info.accessibleContext)
    # 读完即释放该单元格 JAB 句柄(同 get_table_cell_text)
    release_contexts(jab, vm_id, [cell_info.accessibleContext])
    if not info:
        return "", bool(cell_info.isSelected)

    return info.name.strip() or info.description.strip(), bool(cell_info.isSelected)
