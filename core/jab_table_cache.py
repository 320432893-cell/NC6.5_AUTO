# 职责：按窗口标题缓存表格上下文，并按窗读取单元格/计数
# 不做什么：不读取单元格底层 ctypes 结构，不做 context path 动作，不发送全局键盘输入
# 允许依赖层：标准库 ctypes/os/time、JABOperator 暴露能力、tools.jab_probe、core.jab_table_* 同层
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import


from core.logger import log

from core.jab_table_find import find_tables_once
from core.jab_table_read import read_all_table_cells


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
