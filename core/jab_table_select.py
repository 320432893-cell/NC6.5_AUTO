# 职责：JAB 表格行选择——选行/可见表选择/按记录定位选择
# 不做什么：不读取单元格底层 ctypes 结构，不做 context path 动作，不发送全局键盘输入
# 允许依赖层：标准库 ctypes/os/time、JABOperator 暴露能力、tools.jab_probe、core.jab_table_* 同层
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import

import time

from core.logger import log
from core.utils import check_abort

from core.jab_table_cache import clear_table_cache, get_window_table_cache
from core.jab_table_find import find_main_table, find_tables_once


def select_table_rows(jab, rows, selection_col=None, clear=True, timeout=None):
    jab.ensure_started()
    table_context, vm_id, owned_contexts, table_info = find_main_table(
        jab, timeout=timeout
    )
    if not table_context:
        log.warning("JAB 未找到业务表格")
        return False

    try:
        selection_col = resolve_selection_col(jab, selection_col)
        if not has_selection_api(jab):
            log.warning("当前 JAB DLL 不支持 selection API")
            return False

        if clear:
            jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)

        for row in rows:
            if row < 0 or row >= table_info.rowCount:
                raise ValueError(f"行号越界: {row}, rowCount={table_info.rowCount}")
            if selection_col < 0 or selection_col >= table_info.columnCount:
                raise ValueError(
                    f"列号越界: {selection_col}, colCount={table_info.columnCount}"
                )
            child_index = row * table_info.columnCount + selection_col
            jab.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)

        selected_indexes = get_selected_child_indexes(
            jab, vm_id, table_context, table_info.rowCount * table_info.columnCount
        )
        expected_indexes = [
            row * table_info.columnCount + selection_col for row in rows
        ]
        ok = all(index in selected_indexes for index in expected_indexes)
        log.info(
            f"JAB 选中表格行: rows={rows} col={selection_col} "
            f"expected={expected_indexes} selected={selected_indexes[:20]}"
        )
        return ok
    finally:
        jab.release_contexts(vm_id, owned_contexts)



def resolve_selection_col(jab, selection_col):
    if selection_col is not None:
        return selection_col
    return jab.selection_col



def has_selection_api(jab):
    return (
        hasattr(jab.dll, "clearAccessibleSelectionFromContext")
        and hasattr(jab.dll, "addAccessibleSelectionFromContext")
        and hasattr(jab.dll, "isAccessibleChildSelectedFromContext")
    )



def get_selected_child_indexes(jab, vm_id, context, child_count):
    if not hasattr(jab.dll, "isAccessibleChildSelectedFromContext"):
        return []
    selected = []
    for index in range(child_count):
        if jab.dll.isAccessibleChildSelectedFromContext(vm_id, context, index):
            selected.append(index)
    return selected



def select_visible_table_rows(
    jab, table_index, rows, window_title=None, selection_col=0, timeout=None
):
    jab.ensure_started()
    deadline = time.time() + (timeout or jab.search_timeout)

    while time.time() < deadline:
        check_abort()
        ok = select_visible_table_rows_once(
            jab,
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
    jab, table_index, rows, window_title=None, selection_col=0
):
    if window_title is not None:
        cached = get_window_table_cache(jab, window_title)
        if cached is not None:
            for table in cached:
                if table["table_index"] != table_index:
                    continue
                try:
                    table_info = jab.get_table_info(table["vm_id"], table["context"])
                    if not table_info:
                        raise RuntimeError("cached table is no longer readable")
                    return select_table_rows_from_context(
                        jab,
                        table_index,
                        rows,
                        table["context"],
                        table["vm_id"],
                        table_info,
                        window_title=window_title,
                        selection_col=selection_col,
                    )
                except Exception as exc:
                    log.warning(f"JAB 选行缓存失效，回退全量查找: {exc}")
                    clear_table_cache(jab, window_title)
                    break

    tables = find_tables_once(jab)
    for current_index, (
        table_context,
        vm_id,
        owned_contexts,
        table_info,
        window_info,
    ) in enumerate(tables):
        try:
            if current_index != table_index:
                continue
            if window_title is not None and window_info.get("title") != window_title:
                return False
            if not has_selection_api(jab):
                log.warning("当前 JAB DLL 不支持 selection API")
                return False
            if selection_col < 0 or selection_col >= table_info.columnCount:
                raise ValueError(
                    f"列号越界: {selection_col}, colCount={table_info.columnCount}"
                )

            return select_table_rows_from_context(
                jab,
                table_index,
                rows,
                table_context,
                vm_id,
                table_info,
                window_title=window_title,
                selection_col=selection_col,
            )
        finally:
            jab.release_contexts(vm_id, owned_contexts)

    return False



def select_table_rows_from_context(
    jab,
    table_index,
    rows,
    table_context,
    vm_id,
    table_info,
    window_title=None,
    selection_col=0,
):
    if not has_selection_api(jab):
        log.warning("当前 JAB DLL 不支持 selection API")
        return False
    if selection_col < 0 or selection_col >= table_info.columnCount:
        raise ValueError(
            f"列号越界: {selection_col}, colCount={table_info.columnCount}"
        )

    jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
    expected_indexes = []
    for row in rows:
        if row < 0 or row >= table_info.rowCount:
            raise ValueError(f"行号越界: {row}, rowCount={table_info.rowCount}")
        child_index = row * table_info.columnCount + selection_col
        expected_indexes.append(child_index)
        jab.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)

    selected_indexes = get_selected_child_indexes(
        jab, vm_id, table_context, table_info.rowCount * table_info.columnCount
    )
    ok = all(index in selected_indexes for index in expected_indexes)
    log.info(
        f"JAB 选中可见表格行: table={table_index} window={window_title} "
        f"rows={rows} expected={expected_indexes} selected={selected_indexes[:40]}"
    )
    return ok
