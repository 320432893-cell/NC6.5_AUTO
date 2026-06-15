# 职责：读取和筛选收款单查询结果表，提供结果表路径枚举能力。
# 不做什么：不处理翻页按钮、不设置分页大小、不生成匹配报告、不解析 CLI 参数。
# 允许依赖层：JABOperator 暴露的窗口/上下文/表格读取接口、tools.jab_probe 的 JOBJECT 类型。
# 谁不应该 import：与收款单查询无关的业务流程、临时探针脚本。

import ctypes

from tools.jab_probe import JOBJECT


def summarize_receipt_tables(jab, query_cfg, scope_hwnd=None):
    indexes = receipt_result_read_columns(query_cfg)
    min_cols = max(indexes) + 1 if indexes else None
    if hasattr(jab, "read_table_summaries"):
        return [
            table
            for table in jab.read_table_summaries(
                min_rows=1,
                min_cols=min_cols,
                scope_hwnd=scope_hwnd,
            )
            if is_receipt_result_table_candidate(table, query_cfg)
        ]
    tables = read_receipt_tables(
        jab,
        query_cfg,
        max_rows=0,
        max_cols=0,
        read_columns=[],
        scope_hwnd=scope_hwnd,
    )
    return [
        {
            "table_index": table.get("table_index"),
            "row_count": table.get("row_count"),
            "col_count": table.get("col_count"),
        }
        for table in tables
        if is_receipt_result_table_candidate(table, query_cfg)
    ]


def receipt_result_read_columns(query_cfg, include_amount_candidates=False):
    indexes = query_cfg.get("result_column_indexes") or {}
    columns = {
        int(column)
        for column in indexes.values()
        if isinstance(column, int) and column >= 0
    }
    if include_amount_candidates:
        columns.update({6, 7, 8})
    return sorted(columns)


def read_receipt_tables(
    jab,
    query_cfg,
    max_rows=500,
    max_cols=80,
    read_columns=None,
    scope_hwnd=None,
):
    if read_columns and hasattr(jab, "read_all_table_selected_columns"):
        tables = jab.read_all_table_selected_columns(
            read_columns,
            max_rows=max_rows,
            min_rows=1,
            min_cols=max(read_columns) + 1,
            scope_hwnd=scope_hwnd,
        )
    else:
        try:
            tables = jab.read_all_table_cells(
                max_rows=max_rows,
                max_cols=max_cols,
                scope_hwnd=scope_hwnd,
            )
        except TypeError:
            tables = jab.read_all_table_cells(max_rows=max_rows, max_cols=max_cols)
    return [
        table for table in tables if is_receipt_result_table_candidate(table, query_cfg)
    ]


def read_receipt_result_table_by_path(
    jab,
    table_path,
    table_hwnd,
    query_cfg,
    max_rows=500,
    max_cols=80,
    read_columns=None,
    table_index=0,
):
    if not table_path or not table_hwnd:
        return []
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        table_path,
        class_name=(query_cfg.get("pagination") or {}).get(
            "window_class", "SunAwtCanvas"
        ),
        scope_hwnd=table_hwnd,
        role="table",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return []
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return []
        min_cols = receipt_result_min_required_columns(query_cfg, read_columns)
        if min_cols is not None and table_info.columnCount < min_cols:
            return []
        if read_columns:
            table = jab.read_table_selected_columns_from_context(
                table_index,
                context,
                vm_id,
                table_info,
                window_info,
                sorted({int(column) for column in read_columns}),
                max_rows=max_rows,
            )
        else:
            table = jab.read_table_cells_from_context(
                table_index,
                context,
                vm_id,
                table_info,
                window_info,
                max_rows=max_rows,
                max_cols=max_cols,
            )
        table["read_method"] = "result-table-path"
        table["path"] = table_path
        return [table] if is_receipt_result_table_candidate(table, query_cfg) else []
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def read_receipt_result_tables_runtime(
    jab,
    query_cfg,
    setup_report,
    max_rows=500,
    max_cols=80,
    read_columns=None,
):
    tables = read_receipt_result_table_by_path(
        jab,
        setup_report.get("result_table_path"),
        setup_report.get("pager_hwnd"),
        query_cfg,
        max_rows=max_rows,
        max_cols=max_cols,
        read_columns=read_columns,
    )
    if tables:
        return tables
    return read_receipt_tables(
        jab,
        query_cfg,
        max_rows=max_rows,
        max_cols=max_cols,
        read_columns=read_columns,
        scope_hwnd=setup_report.get("pager_hwnd"),
    )


def is_receipt_result_table_candidate(table, query_cfg):
    if table.get("row_count", 0) <= 0:
        return False
    min_cols = receipt_result_min_required_columns(query_cfg)
    if min_cols is None:
        return int(table.get("col_count", 0)) > 0
    if int(table.get("col_count", 0)) < min_cols:
        return False
    return receipt_result_table_rows_match(table, query_cfg)


def enumerate_receipt_result_table_paths(jab, query_cfg, window_class):
    return [
        candidate
        for candidate in enumerate_visible_table_paths(jab, window_class)
        if is_receipt_result_table_candidate(candidate, query_cfg)
    ]


def enumerate_visible_table_paths(jab, window_class):
    if not hasattr(jab, "dll"):
        return []
    jab.ensure_started()
    if jab.dll is None:
        return []
    result = []
    table_index = 0
    for hwnd, title, class_name, pid, visible in jab.get_scoped_windows(
        include_children=True
    ):
        if class_name != window_class:
            continue
        if not visible:
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vm_id), ctypes.byref(root_context)
        ):
            continue

        tables = find_table_paths_in_context(
            jab,
            vm_id.value,
            root_context.value,
            depth=0,
            index_path=[],
            owned_contexts=[],
        )
        contexts_to_release = []
        for table in tables:
            table_info = table["table_info"]
            candidate = {
                "table_index": table_index,
                "path": "0" + "".join(f".{index}" for index in table["index_path"]),
                "hwnd": int(hwnd),
                "window_title": title,
                "window_class": class_name,
                "window_visible": visible,
                "pid": pid,
                "row_count": int(table_info.rowCount),
                "col_count": int(table_info.columnCount),
            }
            table_index += 1
            result.append(candidate)
            contexts_to_release.extend(table["owned_contexts"])
        if contexts_to_release:
            unique_contexts = list(dict.fromkeys(contexts_to_release))
            jab.release_contexts(vm_id.value, unique_contexts)
    return result


def receipt_result_min_required_columns(query_cfg, read_columns=None):
    columns = set(receipt_result_read_columns(query_cfg))
    if read_columns:
        columns.update(
            int(column)
            for column in read_columns
            if isinstance(column, int) and column >= 0
        )
    if not columns:
        return None
    return max(columns) + 1


def receipt_result_table_rows_match(table, query_cfg):
    rows = table.get("rows") or []
    if not rows:
        return True
    guard_cfg = query_cfg.get("result_guard") or {}
    document_type_col = int(guard_cfg.get("document_type_column", 2))
    document_type = str(guard_cfg.get("document_type", "收款单")).strip()
    if not document_type:
        return True
    sampled_values = []
    for row in rows:
        cells = row.get("cells") or []
        if len(cells) <= document_type_col:
            continue
        value = str(cells[document_type_col] or "").strip()
        if value:
            sampled_values.append(value)
    if not sampled_values:
        return True
    return any(document_type in value for value in sampled_values)


def find_table_paths_in_context(
    jab,
    vm_id,
    context,
    depth,
    index_path,
    owned_contexts,
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return []

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if role == "table":
        table_info = jab.get_table_info(vm_id, context)
        if table_info and table_info.rowCount > 0 and table_info.columnCount > 0:
            return [
                {
                    "index_path": list(index_path),
                    "owned_contexts": list(owned_contexts),
                    "table_info": table_info,
                }
            ]
        return []

    if depth >= jab.max_depth:
        return []

    tables = []
    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_tables = find_table_paths_in_context(
            jab,
            vm_id,
            child,
            depth + 1,
            index_path + [index],
            owned_contexts + [child],
        )
        if child_tables:
            tables.extend(child_tables)
        else:
            jab.release_contexts(vm_id, [child])
    return tables


def first_non_empty_cell(cells):
    for cell in cells:
        text = str(cell or "").strip()
        if text:
            return text
    return ""


def first_non_empty_cell_at(cells, column):
    if column >= len(cells):
        return ""
    return str(cells[column] or "").strip()
