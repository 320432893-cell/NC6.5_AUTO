# 职责：封装 JAB 表格发现、缓存、选择、快照读取和金额/关联方匹配
# 不做什么：不读取单元格底层 ctypes 结构，不做 context path 动作，不发送全局键盘输入
# 允许依赖层：标准库 ctypes/os/time、JABOperator 暴露的 dll/table/context 能力、tools.jab_probe 窗口枚举
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import

import ctypes
import os
from ctypes import wintypes
import time

from core.logger import log
from core.utils import check_abort
from tools.jab_probe import JOBJECT, enum_windows


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


def read_window_table_cells(jab, window_title, max_rows=None, max_cols=None):
    jab.ensure_started()
    cached = get_window_table_cache(jab, window_title)
    if cached is None:
        tables = read_all_table_cells(jab, max_rows=max_rows, max_cols=max_cols)
        return [table for table in tables if table.get("window_title") == window_title]

    result = []
    try:
        for table in cached:
            table_info = jab.get_table_info(table["vm_id"], table["context"])
            if (
                not table_info
                or table_info.rowCount <= 0
                or table_info.columnCount <= 0
            ):
                raise RuntimeError("cached table is no longer readable")
            result.append(
                jab.read_table_cells_from_context(
                    table["table_index"],
                    table["context"],
                    table["vm_id"],
                    table_info,
                    table["window_info"],
                    max_rows=max_rows,
                    max_cols=max_cols,
                )
            )
        log.debug(f"JAB 读取缓存窗口表格: window={window_title} count={len(result)}")
        return result
    except Exception as exc:
        log.warning(f"JAB 窗口表缓存失效，回退全量查找: {exc}")
        clear_table_cache(jab, window_title)
        tables = read_all_table_cells(jab, max_rows=max_rows, max_cols=max_cols)
        return [table for table in tables if table.get("window_title") == window_title]


def read_window_table_counts(jab, window_title):
    jab.ensure_started()
    cached = get_window_table_cache(jab, window_title)
    if cached is None:
        return []

    try:
        result = []
        for table in cached:
            table_info = jab.get_table_info(table["vm_id"], table["context"])
            if not table_info:
                raise RuntimeError("cached table is no longer readable")
            result.append(
                {
                    "table_index": table["table_index"],
                    "window_title": table["window_info"].get("title"),
                    "window_class": table["window_info"].get("class_name"),
                    "row_count": table_info.rowCount,
                    "col_count": table_info.columnCount,
                }
            )
        log.debug(
            f"JAB 读取缓存窗口表格行数: window={window_title} count={len(result)}"
        )
        return result
    except Exception as exc:
        log.warning(f"JAB 窗口表行数缓存失效，回退全量查找: {exc}")
        clear_table_cache(jab, window_title)
        cached = get_window_table_cache(jab, window_title)
        if cached is None:
            return []
        result = []
        for table in cached:
            table_info = jab.get_table_info(table["vm_id"], table["context"])
            if not table_info:
                continue
            result.append(
                {
                    "table_index": table["table_index"],
                    "window_title": table["window_info"].get("title"),
                    "window_class": table["window_info"].get("class_name"),
                    "row_count": table_info.rowCount,
                    "col_count": table_info.columnCount,
                }
            )
        return result


def get_window_table_cache(jab, window_title):
    cached = jab.table_cache.get(window_title)
    if cached is not None:
        return cached

    tables = find_tables_once(jab)
    matches = []
    for table_index, (
        table_context,
        vm_id,
        owned_contexts,
        _table_info,
        window_info,
    ) in enumerate(tables):
        if window_info.get("title") == window_title:
            matches.append(
                {
                    "table_index": table_index,
                    "context": table_context,
                    "vm_id": vm_id,
                    "owned_contexts": owned_contexts,
                    "window_info": window_info,
                }
            )
        else:
            jab.release_contexts(vm_id, owned_contexts)

    if not matches:
        return None

    jab.table_cache[window_title] = matches
    log.debug(f"JAB 缓存窗口表格: window={window_title} count={len(matches)}")
    return matches


def clear_table_cache(jab, window_title=None):
    if not jab.table_cache:
        return
    keys = [window_title] if window_title is not None else list(jab.table_cache)
    for key in keys:
        cached = jab.table_cache.pop(key, [])
        for table in cached:
            jab.release_contexts(
                table.get("vm_id"),
                table.get("owned_contexts", []),
            )


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


def wait_for_record_visible(
    jab,
    amount,
    partner_name,
    timeout=None,
    selected_first=True,
    max_rows=200,
    max_cols=50,
    window_title=None,
):
    deadline = time.time() + (timeout or jab.search_timeout)
    while time.time() < deadline:
        check_abort()
        found = find_record_in_visible_tables(
            jab,
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
    jab,
    amount,
    partner_name,
    selected_first=True,
    max_rows=200,
    max_cols=50,
    window_title=None,
):
    target_amount = jab.normalize_amount(amount)
    target_partner = jab.normalize_text(partner_name)
    if target_amount is None or not target_partner:
        return None

    tables = read_all_table_cells(jab, max_rows=max_rows, max_cols=max_cols)
    candidates = []
    fallback = []

    for table in tables:
        if window_title is not None and table.get("window_title") != window_title:
            continue
        for row in table["rows"]:
            normalized_cells = [jab.normalize_text(cell) for cell in row["cells"]]
            row_text = "".join(normalized_cells)
            amount_match = any(
                jab.normalize_amount(cell) == target_amount for cell in row["cells"]
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


def select_record_in_visible_tables_once(
    jab,
    amount,
    partner_name,
    window_title=None,
    selection_col=0,
    max_rows=200,
    max_cols=50,
):
    target_amount = jab.normalize_amount(amount)
    target_partner = jab.normalize_text(partner_name)
    if target_amount is None or not target_partner:
        return None

    tables = find_tables_once(jab)
    for table_index, (
        table_context,
        vm_id,
        owned_contexts,
        table_info,
        window_info,
    ) in enumerate(tables):
        try:
            if window_title is not None and window_info.get("title") != window_title:
                continue

            row_limit = min(table_info.rowCount, max_rows)
            col_limit = min(table_info.columnCount, max_cols)
            for row in range(row_limit):
                cells = []
                amount_match = False
                for col in range(col_limit):
                    text = jab.get_table_cell_text(vm_id, table_context, row, col)
                    cells.append(text)
                    if jab.normalize_amount(text) == target_amount:
                        amount_match = True

                row_text = "".join(jab.normalize_text(cell) for cell in cells)
                if not amount_match or target_partner not in row_text:
                    continue

                if not has_selection_api(jab):
                    log.warning("当前 JAB DLL 不支持 selection API")
                    return None
                if selection_col < 0 or selection_col >= table_info.columnCount:
                    raise ValueError(
                        f"列号越界: {selection_col}, colCount={table_info.columnCount}"
                    )

                jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
                child_index = row * table_info.columnCount + selection_col
                jab.dll.addAccessibleSelectionFromContext(
                    vm_id, table_context, child_index
                )
                selected_indexes = get_selected_child_indexes(
                    jab,
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
            jab.release_contexts(vm_id, owned_contexts)

    return None


def resolve_amount_col(jab, amount_col):
    if amount_col is not None:
        return amount_col
    return jab.amount_col


def resolve_partner_col(jab, partner_col):
    if partner_col is not None:
        return partner_col
    return jab.partner_col


def find_main_table(jab, timeout=None):
    deadline = time.time() + (timeout or jab.search_timeout)

    while time.time() < deadline:
        check_abort()
        tables = find_tables_once(jab)
        if tables:
            tables.sort(
                key=lambda item: item[3].rowCount * item[3].columnCount,
                reverse=True,
            )
            table_context, vm_id, owned_contexts, table_info, _window_info = tables[0]

            log.debug(
                f"JAB 找到业务表格: rows={table_info.rowCount} cols={table_info.columnCount}"
            )
            return table_context, vm_id, owned_contexts, table_info
        time.sleep(0.2)

    return None, None, [], None


def find_tables_once(jab, scope_hwnd=None):
    tables = []
    windows = enum_windows(include_children=True)
    scoped_hwnds = window_descendant_hwnds(scope_hwnd)

    for hwnd, title, class_name, pid, visible in windows:
        if scoped_hwnds is not None and int(hwnd) not in scoped_hwnds:
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            continue

        tables.extend(
            find_tables_in_tree(
                jab,
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


def window_descendant_hwnds(scope_hwnd):
    if scope_hwnd is None or os.name != "nt":
        return None

    scoped = {int(scope_hwnd)}
    if not hasattr(ctypes, "windll"):
        return scoped

    user32 = ctypes.windll.user32
    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def child_callback(hwnd, _lparam):
        scoped.add(int(hwnd))
        return True

    user32.EnumChildWindows(wintypes.HWND(scope_hwnd), enum_proc(child_callback), 0)
    return scoped


def find_tables_in_tree(jab, vm_id, context, depth, owned_path, window_info=None):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return []

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if role == "table":
        table_info = jab.get_table_info(vm_id, context)
        if table_info and table_info.rowCount > 0 and table_info.columnCount > 0:
            return [(context, vm_id, list(owned_path), table_info, window_info or {})]
        return []

    if depth >= jab.max_depth:
        return []

    tables = []
    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue

        child_tables = find_tables_in_tree(
            jab,
            vm_id,
            child,
            depth + 1,
            owned_path + [child],
            window_info=window_info,
        )
        if child_tables:
            tables.extend(child_tables)
        else:
            jab.release_contexts(vm_id, [child])

    return tables
