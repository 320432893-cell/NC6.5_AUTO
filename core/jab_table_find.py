# 职责：JAB 表格发现——查找窗口内表格、主表、窗口子句柄遍历
# 不做什么：不读取单元格底层 ctypes 结构，不做 context path 动作，不发送全局键盘输入
# 允许依赖层：标准库 ctypes/os/time、JABOperator 暴露能力、tools.jab_probe、core.jab_table_* 同层
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import

import ctypes
import os
from ctypes import wintypes
import time

from core.logger import log
from core.utils import check_abort
from tools.jab_probe import JOBJECT, enum_windows


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
