import argparse
import ctypes
import json
import os
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready  # noqa: E402
from tools.jab_probe import JOBJECT, enum_windows  # noqa: E402
from tools.receipt_account_reference_try import STOP_HOTKEY, is_stop_hotkey_pressed  # noqa: E402
from tools.receipt_body_table_locator import KEY_COLUMNS, table_bounds  # noqa: E402

START_DELAY_SECONDS = 2
MAX_ROWS = 3
WATCH_CLASSES = {"YonyouUWnd", "SunAwtFrame", "SunAwtCanvas", "SunAwtDialog"}
PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010
LIST_MODULES_ALL = 0x03


def print_header():
    print("测试功能：收款单界面 JAB 只读探测")
    print()
    print("目标：")
    print("1. 打印当前前台窗口和同进程 NC/Java 窗口")
    print("2. 列出可连接的 JAB 根节点")
    print("3. 列出 JAB table 或带 table interface 的节点")
    print("4. 只标注结构特征，不评分、不自动判定明细表")
    print()
    print("不会做：写入、选择单元格、点击、保存、暂存、关闭窗口")
    print(f"紧急停止：按 {STOP_HOTKEY}")
    print(f"启动后等待：{START_DELAY_SECONDS} 秒，用来切到 NC 窗口")
    print("=" * 60)


def collect_all_tables(jab, max_rows=MAX_ROWS):
    jab.ensure_started()
    foreground = get_foreground_window_info()
    foreground_pid = foreground.get("pid") if foreground else None
    scoped_windows = describe_scoped_windows(jab, foreground_pid)
    jab_status = describe_jab_status(jab, scoped_windows)
    jab_health = check_jab_ready(jab)
    tables = []
    roots = []
    stats = {"max_depth_seen": 0, "depth_limit_hits": 0}
    for item in scoped_windows:
        hwnd = item["hwnd"]
        title = item["title"]
        class_name = item["class_name"]
        pid = item["pid"]
        visible = item["visible"]
        if is_stop_hotkey_pressed():
            return {
                "stopped_by_hotkey": True,
                "tables": tables,
                "roots": roots,
                "stats": stats,
                "foreground": foreground,
                "windows": scoped_windows,
                "jab_status": jab_status,
                "jab_health": jab_health,
            }
        if not jab.dll.isJavaWindow(hwnd):
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vm_id), ctypes.byref(root_context)
        ):
            continue

        window = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
        }
        root_info = jab.get_context_info(vm_id.value, root_context.value)
        roots.append(
            {
                "window": window,
                "root_role": (
                    root_info.role_en_US.strip() or root_info.role.strip()
                    if root_info
                    else None
                ),
                "root_name": root_info.name.strip() if root_info else None,
                "root_description": (
                    root_info.description.strip() if root_info else None
                ),
                "root_children": root_info.childrenCount if root_info else None,
            }
        )
        try:
            collect_tables_in_tree(
                jab,
                vm_id.value,
                root_context.value,
                depth=0,
                path_indexes=[],
                owned_contexts=[],
                window=window,
                result=tables,
                max_rows=max_rows,
                stats=stats,
            )
        finally:
            jab.release_contexts(vm_id.value, [root_context.value])
    return {
        "stopped_by_hotkey": False,
        "tables": tables,
        "roots": roots,
        "stats": stats,
        "max_depth_config": jab.max_depth,
        "max_children_config": jab.max_children,
        "foreground": foreground,
        "windows": scoped_windows,
        "jab_status": jab_status,
        "jab_health": jab_health,
    }


def describe_jab_status(jab, windows):
    java_pids = sorted(
        {
            int(item["pid"])
            for item in windows
            if str(item.get("class_name") or "").startswith("SunAwt")
        }
    )
    direct_contexts = []
    for item in windows:
        if item.get("class_name") not in (
            "SunAwtFrame",
            "SunAwtCanvas",
            "SunAwtDialog",
        ):
            continue
        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        ok = bool(
            jab.dll.getAccessibleContextFromHWND(
                item["hwnd"], ctypes.byref(vm_id), ctypes.byref(root_context)
            )
        )
        root_info = None
        if ok:
            info = jab.get_context_info(vm_id.value, root_context.value)
            root_info = {
                "role": (info.role_en_US.strip() or info.role.strip())
                if info
                else None,
                "name": info.name.strip() if info else None,
                "description": info.description.strip() if info else None,
                "children": info.childrenCount if info else None,
            }
            jab.release_contexts(vm_id.value, [root_context.value])
        direct_contexts.append(
            {
                "hwnd": item["hwnd"],
                "title": item["title"],
                "class_name": item["class_name"],
                "pid": item["pid"],
                "visible": item["visible"],
                "is_java": item["is_java"],
                "get_context_ok": ok,
                "vm_id": vm_id.value if ok else None,
                "root": root_info,
            }
        )

    return {
        "dll_path": str(jab.loaded_path),
        "has_initializeAccessBridge": bool(hasattr(jab.dll, "initializeAccessBridge")),
        "has_Windows_run": bool(hasattr(jab.dll, "Windows_run")),
        "accessibility_files": read_accessibility_files(),
        "java_processes": [describe_process(pid) for pid in java_pids],
        "java_pid_windows": describe_windows_for_pids(jab, java_pids),
        "access_bridge_status_windows": read_access_bridge_status_windows(java_pids),
        "direct_contexts": direct_contexts,
    }


def read_access_bridge_status_windows(pids):
    wanted = {int(pid) for pid in pids}
    if not wanted:
        return []
    result = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if int(pid) not in wanted:
            continue
        if title != "Access Bridge status" and class_name != "#32770":
            continue
        result.append(
            {
                "hwnd": int(hwnd),
                "pid": int(pid),
                "title": title,
                "class_name": class_name,
                "visible": visible,
                "children": read_child_window_texts(hwnd),
            }
        )
    return result


def read_child_window_texts(parent_hwnd):
    user32 = ctypes.windll.user32
    enum_proc = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
    children = []

    def callback(hwnd, _lparam):
        title, class_name, pid, visible = read_window_info(hwnd)
        children.append(
            {
                "hwnd": int(hwnd),
                "pid": pid,
                "class_name": class_name,
                "title": title,
                "visible": visible,
                "text": read_window_text(hwnd),
            }
        )
        return 1

    user32.EnumChildWindows(parent_hwnd, enum_proc(callback), 0)
    children.sort(key=lambda item: (item["class_name"], item["hwnd"]))
    return children


def read_window_text(hwnd):
    user32 = ctypes.windll.user32
    length = user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(max(length + 1, 1024))
    user32.GetWindowTextW(hwnd, buffer, len(buffer))
    return buffer.value


def describe_windows_for_pids(jab, pids):
    wanted = {int(pid) for pid in pids}
    if not wanted:
        return []
    user32 = ctypes.windll.user32
    result = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if int(pid) not in wanted:
            continue
        parent = user32.GetParent(hwnd)
        root = user32.GetAncestor(hwnd, 2)
        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        is_java = bool(jab.dll.isJavaWindow(hwnd))
        get_context_ok = bool(
            jab.dll.getAccessibleContextFromHWND(
                hwnd, ctypes.byref(vm_id), ctypes.byref(root_context)
            )
        )
        root_summary = None
        if get_context_ok:
            info = jab.get_context_info(vm_id.value, root_context.value)
            root_summary = {
                "role": (info.role_en_US.strip() or info.role.strip())
                if info
                else None,
                "name": info.name.strip() if info else None,
                "children": info.childrenCount if info else None,
            }
            jab.release_contexts(vm_id.value, [root_context.value])
        result.append(
            {
                "hwnd": int(hwnd),
                "title": title,
                "class_name": class_name,
                "pid": int(pid),
                "visible": visible,
                "parent_hwnd": int(parent) if parent else None,
                "root_hwnd": int(root) if root else None,
                "is_java": is_java,
                "get_context_ok": get_context_ok,
                "root": root_summary,
            }
        )
    result.sort(
        key=lambda item: (
            item["pid"],
            0 if item["visible"] else 1,
            item["class_name"],
            item["title"],
            item["hwnd"],
        )
    )
    return result


def describe_process(pid):
    modules = list_process_modules(pid)
    bridge_modules = [
        item
        for item in modules
        if any(
            needle in item["name"].lower()
            for needle in ("access", "bridge", "jawt", "java")
        )
    ]
    return {
        "pid": int(pid),
        "modules_read_ok": bool(modules),
        "module_count": len(modules),
        "bridge_modules": bridge_modules,
        "has_java_access_bridge_dll": any(
            item["name"].lower() == "javaaccessbridge-64.dll" for item in modules
        ),
        "has_jawt_access_bridge_dll": any(
            item["name"].lower() == "jawtaccessbridge-64.dll" for item in modules
        ),
    }


def list_process_modules(pid):
    if sys.platform != "win32":
        return []
    kernel32 = ctypes.windll.kernel32
    psapi = ctypes.windll.psapi
    kernel32.OpenProcess.argtypes = [ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong]
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    kernel32.CloseHandle.restype = ctypes.c_int
    psapi.EnumProcessModulesEx.argtypes = [
        ctypes.c_void_p,
        ctypes.POINTER(ctypes.c_void_p),
        ctypes.c_ulong,
        ctypes.POINTER(ctypes.c_ulong),
        ctypes.c_ulong,
    ]
    psapi.EnumProcessModulesEx.restype = ctypes.c_int
    psapi.GetModuleBaseNameW.argtypes = [
        ctypes.c_void_p,
        ctypes.c_void_p,
        ctypes.c_wchar_p,
        ctypes.c_ulong,
    ]
    psapi.GetModuleBaseNameW.restype = ctypes.c_ulong
    psapi.GetModuleFileNameExW.argtypes = [
        ctypes.c_void_p,
        ctypes.c_void_p,
        ctypes.c_wchar_p,
        ctypes.c_ulong,
    ]
    psapi.GetModuleFileNameExW.restype = ctypes.c_ulong
    handle = kernel32.OpenProcess(
        PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, int(pid)
    )
    if not handle:
        return []
    try:
        needed = ctypes.c_ulong()
        module_array = (ctypes.c_void_p * 2048)()
        ok = psapi.EnumProcessModulesEx(
            handle,
            module_array,
            ctypes.sizeof(module_array),
            ctypes.byref(needed),
            LIST_MODULES_ALL,
        )
        if not ok:
            return []
        count = min(int(needed.value // ctypes.sizeof(ctypes.c_void_p)), 2048)
        result = []
        for index in range(count):
            module = module_array[index]
            name_buffer = ctypes.create_unicode_buffer(260)
            path_buffer = ctypes.create_unicode_buffer(1024)
            psapi.GetModuleBaseNameW(handle, module, name_buffer, len(name_buffer))
            psapi.GetModuleFileNameExW(handle, module, path_buffer, len(path_buffer))
            result.append({"name": name_buffer.value, "path": path_buffer.value})
        return result
    finally:
        kernel32.CloseHandle(handle)


def read_accessibility_files():
    paths = [
        Path(os.environ.get("USERPROFILE", "")) / ".accessibility.properties",
        Path(os.environ.get("LOCALAPPDATA", ""))
        / "UClient"
        / "share"
        / "java1.7.0_51-x64"
        / "lib"
        / "accessibility.properties",
    ]
    result = []
    for path in paths:
        if not str(path):
            continue
        item = {"path": str(path), "exists": path.exists()}
        if path.exists():
            try:
                text = path.read_text(encoding="utf-8", errors="replace")
            except OSError as exc:
                item["error"] = str(exc)
            else:
                item["has_access_bridge_line"] = (
                    "assistive_technologies=com.sun.java.accessibility.AccessBridge"
                    in text
                )
        result.append(item)
    return result


def get_foreground_window_info():
    if sys.platform != "win32":
        return None
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None
    title, class_name, pid, visible = read_window_info(hwnd)
    parent = user32.GetParent(hwnd)
    root = user32.GetAncestor(hwnd, 2)
    return {
        "hwnd": int(hwnd),
        "title": title,
        "class_name": class_name,
        "pid": pid,
        "visible": visible,
        "parent_hwnd": int(parent) if parent else None,
        "root_hwnd": int(root) if root else None,
    }


def read_window_info(hwnd):
    user32 = ctypes.windll.user32
    length = user32.GetWindowTextLengthW(hwnd)
    title = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, title, length + 1)
    class_name = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_name, 256)
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return (
        title.value,
        class_name.value,
        int(pid.value),
        bool(user32.IsWindowVisible(hwnd)),
    )


def describe_scoped_windows(jab, foreground_pid):
    windows = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        is_watch_class = class_name in WATCH_CLASSES
        is_java_process_window = class_name.startswith("SunAwt")
        if (
            foreground_pid
            and pid != foreground_pid
            and not (is_watch_class or is_java_process_window)
        ):
            continue
        if not (is_watch_class or is_java_process_window) and not title:
            continue
        windows.append(
            {
                "hwnd": int(hwnd),
                "title": title,
                "class_name": class_name,
                "pid": pid,
                "visible": visible,
                "is_java": bool(jab.dll.isJavaWindow(hwnd)),
            }
        )
    windows.sort(
        key=lambda item: (
            0 if item["pid"] == foreground_pid else 1,
            0 if item["visible"] else 1,
            item["class_name"],
            item["title"],
            item["hwnd"],
        )
    )
    return windows


def collect_tables_in_tree(
    jab,
    vm_id,
    context,
    depth,
    path_indexes,
    owned_contexts,
    window,
    result,
    max_rows,
    stats,
):
    stats["max_depth_seen"] = max(stats.get("max_depth_seen", 0), depth)
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    table_info = jab.get_table_info(vm_id, context)
    if table_info and table_info.rowCount > 0 and table_info.columnCount > 0:
        result.append(
            describe_table(
                jab,
                vm_id,
                context,
                table_info,
                window,
                "0" + "".join(f".{index}" for index in path_indexes),
                max_rows=max_rows,
                depth=depth,
                role=role,
            )
        )
        return

    if depth >= jab.max_depth:
        stats["depth_limit_hits"] = stats.get("depth_limit_hits", 0) + 1
        return

    for index in range(min(info.childrenCount, jab.max_children)):
        if is_stop_hotkey_pressed():
            return
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        try:
            collect_tables_in_tree(
                jab,
                vm_id,
                child,
                depth + 1,
                path_indexes + [index],
                owned_contexts + [child],
                window,
                result,
                max_rows,
                stats,
            )
        finally:
            jab.release_contexts(vm_id, [child])


def describe_table(
    jab, vm_id, context, table_info, window, path, max_rows, depth, role
):
    rows = []
    for row in range(min(table_info.rowCount, max_rows)):
        cells = {}
        selected = False
        for col in KEY_COLUMNS:
            if col >= table_info.columnCount:
                continue
            text, is_selected = jab.get_table_cell_text_and_selection(
                vm_id, context, row, col
            )
            cells[str(col)] = text
            selected = selected or is_selected
        rows.append({"row_index": row, "selected": selected, "cells": cells})

    selected_indexes = jab.get_selected_child_indexes(
        vm_id, context, table_info.rowCount * table_info.columnCount
    )
    return {
        "path": path,
        "path_note": "窗口 root context 内相对 path，不是全局唯一 path",
        "depth": depth,
        "role": role,
        "window": window,
        "row_count": table_info.rowCount,
        "col_count": table_info.columnCount,
        "bounds": table_bounds(jab, vm_id, context),
        "selected_indexes": selected_indexes,
        "rows": rows,
        "features": infer_features(table_info, window, rows, selected_indexes),
    }


def infer_features(table_info, window, rows, selected_indexes):
    features = []
    if table_info.columnCount == 25:
        features.append("25列候选")
    if table_info.rowCount == 1:
        features.append("单行")
    if table_info.rowCount >= 2:
        features.append("多行")
    if window.get("class_name") == "SunAwtCanvas":
        features.append("SunAwtCanvas")
    if selected_indexes:
        features.append(f"有选中项={selected_indexes}")

    first = rows[0]["cells"] if rows else {}
    if any(first.values()):
        features.append("首行关键列有文本")
    if first.get("1") in ("货款", "手续费"):
        features.append("col1像收款业务类型")
    if first.get("3") in ("人民币", "美元"):
        features.append("col3像币种")
    if first.get("4"):
        features.append("col4有账户文本")
    if first.get("5"):
        features.append("col5有科目文本")
    if first.get("7"):
        features.append("col7有金额文本")
    if first.get("11"):
        features.append("col11有结算方式文本")
    return features


def print_report(report):
    if report.get("stopped_by_hotkey"):
        print(f"已停止：检测到紧急停止键 {STOP_HOTKEY}。")
        return

    tables = report.get("tables") or []
    roots = report.get("roots") or []
    foreground = report.get("foreground")
    print()
    print("当前前台窗口：")
    if foreground:
        print(
            f"  title={foreground.get('title') or '<无标题>'} "
            f"class={foreground.get('class_name')} pid={foreground.get('pid')} "
            f"hwnd={foreground.get('hwnd')} root={foreground.get('root_hwnd')}"
        )
    else:
        print("  未读到前台窗口。")
    print()
    windows = report.get("windows") or []
    print(f"同进程/NC Java 相关窗口数量：{len(windows)}")
    for item in windows[:30]:
        print(
            f"  {'JAB' if item.get('is_java') else '---'} "
            f"title={item.get('title') or '<无标题>'} "
            f"class={item.get('class_name')} visible={item.get('visible')} "
            f"pid={item.get('pid')} hwnd={item.get('hwnd')}"
        )
    if len(windows) > 30:
        print(f"  ... 已省略 {len(windows) - 30} 个窗口")
    print()
    jab_status = report.get("jab_status") or {}
    jab_health = report.get("jab_health") or {}
    print("JAB 状态：")
    print(f"  健康检查：ok={jab_health.get('ok')} reason={jab_health.get('reason')}")
    print(f"  DLL={jab_status.get('dll_path')}")
    print(
        "  初始化接口："
        f"initializeAccessBridge={jab_status.get('has_initializeAccessBridge')} "
        f"Windows_run={jab_status.get('has_Windows_run')}"
    )
    print("  accessibility 配置：")
    for item in jab_status.get("accessibility_files") or []:
        print(
            f"    {item.get('path')}: exists={item.get('exists')} "
            f"accessBridgeLine={item.get('has_access_bridge_line')}"
        )
    print("  Java 进程模块：")
    java_processes = jab_status.get("java_processes") or []
    if not java_processes:
        print("    未发现 SunAwt 对应 Java 进程。")
    for process in java_processes:
        print(
            f"    pid={process.get('pid')} modulesRead={process.get('modules_read_ok')} "
            f"moduleCount={process.get('module_count')} "
            f"JavaAccessBridge={process.get('has_java_access_bridge_dll')} "
            f"JAWTAccessBridge={process.get('has_jawt_access_bridge_dll')}"
        )
        for module in process.get("bridge_modules") or []:
            print(f"      {module.get('name')} -> {module.get('path')}")
    print("  Java 进程窗口逐个取 context：")
    pid_windows = jab_status.get("java_pid_windows") or []
    if not pid_windows:
        print("    未发现 Java PID 下的窗口。")
    for item in pid_windows:
        root = item.get("root") or {}
        print(
            f"    hwnd={item.get('hwnd')} class={item.get('class_name')} "
            f"title={item.get('title') or '<无标题>'} visible={item.get('visible')} "
            f"parent={item.get('parent_hwnd')} rootHwnd={item.get('root_hwnd')} "
            f"isJava={item.get('is_java')} getContext={item.get('get_context_ok')} "
            f"rootRole={root.get('role')} children={root.get('children')}"
        )
    print("  Access Bridge status 窗口文本：")
    status_windows = jab_status.get("access_bridge_status_windows") or []
    if not status_windows:
        print("    未发现 Access Bridge status 窗口。")
    for window in status_windows:
        print(
            f"    hwnd={window.get('hwnd')} visible={window.get('visible')} "
            f"class={window.get('class_name')}"
        )
        for child in window.get("children") or []:
            print(
                f"      {child.get('class_name')} hwnd={child.get('hwnd')} "
                f"title={child.get('title')!r} text={child.get('text')!r}"
            )
    print("  SunAwt 直接取 JAB context：")
    direct_contexts = jab_status.get("direct_contexts") or []
    if not direct_contexts:
        print("    未发现 SunAwtFrame/SunAwtCanvas/SunAwtDialog 窗口。")
    for item in direct_contexts:
        root = item.get("root") or {}
        print(
            f"    hwnd={item.get('hwnd')} class={item.get('class_name')} "
            f"visible={item.get('visible')} isJava={item.get('is_java')} "
            f"getContext={item.get('get_context_ok')} "
            f"rootRole={root.get('role')} children={root.get('children')}"
        )
    print()
    print(
        f"探索配置：max_depth={report.get('max_depth_config')} "
        f"max_children={report.get('max_children_config')}"
    )
    stats = report.get("stats") or {}
    print(
        f"探索统计：max_depth_seen={stats.get('max_depth_seen')} "
        f"depth_limit_hits={stats.get('depth_limit_hits')}"
    )
    print(f"Java 根窗口数量：{len(roots)}")
    for index, root in enumerate(roots, start=1):
        window = root.get("window") or {}
        print(
            f"  根 {index}: title={window.get('title') or '<无标题>'} "
            f"class={window.get('class_name')} visible={window.get('visible')} "
            f"hwnd={window.get('hwnd')} root_role={root.get('root_role')} "
            f"children={root.get('root_children')}"
        )
    print()
    print(f"发现 JAB table/interface 数量：{len(tables)}")
    if not tables:
        print("没有发现任何 JAB table/interface。")
        print("这只能说明当前探测入口没读到表，不能直接等同于 JAB 总开关失效。")
        return

    for index, table in enumerate(tables, start=1):
        window = table["window"]
        title = window.get("title") or "<无标题>"
        print("-" * 60)
        print(f"表 {index}")
        print(
            f"窗口：title={title} class={window.get('class_name')} "
            f"visible={window.get('visible')} hwnd={window.get('hwnd')}"
        )
        print(
            f"相对path={table.get('path')} role={table.get('role')} "
            f"depth={table.get('depth')} "
            f"行列={table.get('row_count')} x {table.get('col_count')} "
            f"bounds={table.get('bounds')}"
        )
        print(f"selected={table.get('selected_indexes')}")
        print(f"结构特征：{table.get('features')}")
        for row in table.get("rows") or []:
            print(
                f"  row {row.get('row_index')} selected={row.get('selected')}: "
                f"{row.get('cells')}"
            )


def wait_exit():
    try:
        input("按回车退出...")
    except KeyboardInterrupt:
        print()
        print("已退出。")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default=str(ROOT / "config.json"))
    parser.add_argument("--no-wait", action="store_true")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    print_header()
    if not args.no_wait:
        print()
        print(f"请在 {START_DELAY_SECONDS} 秒内切到 NC 收款单窗口...")
        time.sleep(START_DELAY_SECONDS)
    print("开始探测。")

    try:
        config = load_config(args.config)
        jab = JABOperator(config)
        jab.hide_blank_awt_windows_enabled = False
        try:
            report = collect_all_tables(jab, max_rows=MAX_ROWS)
        finally:
            jab.close()
    except Exception as exc:
        report = {
            "stopped_by_hotkey": False,
            "exception": type(exc).__name__,
            "reason": str(exc),
            "traceback": traceback.format_exc(),
            "tables": [],
        }

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    elif report.get("exception"):
        print()
        print(f"脚本异常：{report.get('exception')}")
        print(f"原因：{report.get('reason')}")
    else:
        print_report(report)
    print()
    if not args.no_wait:
        wait_exit()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
