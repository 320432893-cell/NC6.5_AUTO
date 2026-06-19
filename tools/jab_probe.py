import argparse
import ctypes
import glob
import os
import platform
import threading
import time
from ctypes import wintypes

from tools.jab_environment import uclient_access_bridge_dll_patterns


JOBJECT64 = ctypes.c_longlong
JOBJECT32 = ctypes.c_int
JOBJECT = JOBJECT64 if platform.architecture()[0] == "64bit" else JOBJECT32


class AccessibleContextInfo(ctypes.Structure):
    _fields_ = [
        ("name", ctypes.c_wchar * 1024),
        ("description", ctypes.c_wchar * 1024),
        ("role", ctypes.c_wchar * 256),
        ("role_en_US", ctypes.c_wchar * 256),
        ("states", ctypes.c_wchar * 256),
        ("states_en_US", ctypes.c_wchar * 256),
        ("indexInParent", ctypes.c_int),
        ("childrenCount", ctypes.c_int),
        ("x", ctypes.c_int),
        ("y", ctypes.c_int),
        ("width", ctypes.c_int),
        ("height", ctypes.c_int),
        ("accessibleComponent", wintypes.BOOL),
        ("accessibleAction", wintypes.BOOL),
        ("accessibleSelection", wintypes.BOOL),
        ("accessibleText", wintypes.BOOL),
        ("accessibleInterfaces", wintypes.BOOL),
    ]


class AccessibleTableInfo(ctypes.Structure):
    _fields_ = [
        ("caption", JOBJECT),
        ("summary", JOBJECT),
        ("rowCount", ctypes.c_int),
        ("columnCount", ctypes.c_int),
        ("accessibleContext", JOBJECT),
        ("accessibleTable", JOBJECT),
    ]


class AccessibleTableCellInfo(ctypes.Structure):
    _fields_ = [
        ("accessibleContext", JOBJECT),
        ("index", ctypes.c_int),
        ("row", ctypes.c_int),
        ("column", ctypes.c_int),
        ("rowExtent", ctypes.c_int),
        ("columnExtent", ctypes.c_int),
        ("isSelected", wintypes.BOOL),
    ]


class AccessibleTextInfo(ctypes.Structure):
    _fields_ = [
        ("charCount", ctypes.c_int),
        ("caretIndex", ctypes.c_int),
        ("indexAtPoint", ctypes.c_int),
    ]


class AccessibleActionInfo(ctypes.Structure):
    _fields_ = [
        ("name", ctypes.c_wchar * 256),
    ]


class AccessibleActions(ctypes.Structure):
    _fields_ = [
        ("actionsCount", ctypes.c_int),
        ("actionInfo", AccessibleActionInfo * 256),
    ]


class AccessibleActionsToDo(ctypes.Structure):
    _fields_ = [
        ("actionsCount", ctypes.c_int),
        ("actions", AccessibleActionInfo * 32),
    ]


def iter_candidate_dlls():
    dll_name = (
        "WindowsAccessBridge-64.dll"
        if platform.architecture()[0] == "64bit"
        else "WindowsAccessBridge-32.dll"
    )

    seen = set()

    for item in os.environ.get("PATH", "").split(os.pathsep):
        if not item:
            continue
        path = os.path.join(item, dll_name)
        if path not in seen:
            seen.add(path)
            yield path

    java_home = os.environ.get("JAVA_HOME")
    if java_home:
        path = os.path.join(java_home, "bin", dll_name)
        if path not in seen:
            seen.add(path)
            yield path

    patterns = [
        *uclient_access_bridge_dll_patterns(),
        rf"C:\Program Files\Java\*\bin\{dll_name}",
        rf"C:\Program Files\Eclipse Adoptium\*\bin\{dll_name}",
        rf"C:\Program Files\Microsoft\jdk-*\bin\{dll_name}",
        rf"C:\Program Files (x86)\Java\*\bin\{dll_name}",
        rf"C:\Windows\System32\{dll_name}",
        rf"C:\Windows\SysWOW64\{dll_name}",
    ]
    for pattern in patterns:
        for path in glob.glob(pattern):
            if path not in seen:
                seen.add(path)
                yield path


def load_access_bridge(dll_path=None):
    if dll_path:
        return ctypes.WinDLL(dll_path), dll_path

    errors = []
    for path in iter_candidate_dlls():
        if not os.path.exists(path):
            continue
        try:
            dll = ctypes.WinDLL(path)
            return dll, path
        except OSError as exc:
            errors.append(f"{path}: {exc}")

    raise RuntimeError(
        "Could not load WindowsAccessBridge DLL. Enable Java Access Bridge first "
        "with `jabswitch -enable`, then run this script with Windows Python.\n"
        + "\n".join(errors)
    )


def configure_jab(dll):
    if hasattr(dll, "initializeAccessBridge"):
        dll.initializeAccessBridge.argtypes = []
        dll.initializeAccessBridge.restype = wintypes.BOOL
    elif hasattr(dll, "Windows_run"):
        dll.Windows_run.argtypes = []
        dll.Windows_run.restype = None
    else:
        raise RuntimeError(
            "Access Bridge DLL has no initializeAccessBridge or Windows_run export."
        )

    dll.isJavaWindow.argtypes = [wintypes.HWND]
    dll.isJavaWindow.restype = wintypes.BOOL

    dll.getAccessibleContextFromHWND.argtypes = [
        wintypes.HWND,
        ctypes.POINTER(ctypes.c_long),
        ctypes.POINTER(JOBJECT),
    ]
    dll.getAccessibleContextFromHWND.restype = wintypes.BOOL

    dll.getAccessibleContextInfo.argtypes = [
        ctypes.c_long,
        JOBJECT,
        ctypes.POINTER(AccessibleContextInfo),
    ]
    dll.getAccessibleContextInfo.restype = wintypes.BOOL

    dll.getAccessibleChildFromContext.argtypes = [ctypes.c_long, JOBJECT, ctypes.c_int]
    dll.getAccessibleChildFromContext.restype = JOBJECT

    if hasattr(dll, "getAccessibleTableInfo"):
        dll.getAccessibleTableInfo.argtypes = [
            ctypes.c_long,
            JOBJECT,
            ctypes.POINTER(AccessibleTableInfo),
        ]
        dll.getAccessibleTableInfo.restype = wintypes.BOOL

    if hasattr(dll, "getAccessibleTableCellInfo"):
        dll.getAccessibleTableCellInfo.argtypes = [
            ctypes.c_long,
            JOBJECT,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.POINTER(AccessibleTableCellInfo),
        ]
        dll.getAccessibleTableCellInfo.restype = wintypes.BOOL

    if hasattr(dll, "getAccessibleActions"):
        dll.getAccessibleActions.argtypes = [
            ctypes.c_long,
            JOBJECT,
            ctypes.POINTER(AccessibleActions),
        ]
        dll.getAccessibleActions.restype = wintypes.BOOL

    if hasattr(dll, "doAccessibleActions"):
        dll.doAccessibleActions.argtypes = [
            ctypes.c_long,
            JOBJECT,
            ctypes.POINTER(AccessibleActionsToDo),
            ctypes.POINTER(ctypes.c_int),
        ]
        dll.doAccessibleActions.restype = wintypes.BOOL

    if hasattr(dll, "requestFocus"):
        dll.requestFocus.argtypes = [ctypes.c_long, JOBJECT]
        dll.requestFocus.restype = wintypes.BOOL

    if hasattr(dll, "setTextContents"):
        dll.setTextContents.argtypes = [ctypes.c_long, JOBJECT, ctypes.c_wchar_p]
        dll.setTextContents.restype = wintypes.BOOL

    if hasattr(dll, "getAccessibleTextInfo"):
        dll.getAccessibleTextInfo.argtypes = [
            ctypes.c_long,
            JOBJECT,
            ctypes.POINTER(AccessibleTextInfo),
            ctypes.c_int,
            ctypes.c_int,
        ]
        dll.getAccessibleTextInfo.restype = wintypes.BOOL

    if hasattr(dll, "getAccessibleTextRange"):
        dll.getAccessibleTextRange.argtypes = [
            ctypes.c_long,
            JOBJECT,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_wchar_p,
            ctypes.c_short,
        ]
        dll.getAccessibleTextRange.restype = wintypes.BOOL

    if hasattr(dll, "getAccessibleSelectionCountFromContext"):
        dll.getAccessibleSelectionCountFromContext.argtypes = [ctypes.c_long, JOBJECT]
        dll.getAccessibleSelectionCountFromContext.restype = ctypes.c_int

    if hasattr(dll, "isAccessibleChildSelectedFromContext"):
        dll.isAccessibleChildSelectedFromContext.argtypes = [
            ctypes.c_long,
            JOBJECT,
            ctypes.c_int,
        ]
        dll.isAccessibleChildSelectedFromContext.restype = wintypes.BOOL

    if hasattr(dll, "addAccessibleSelectionFromContext"):
        dll.addAccessibleSelectionFromContext.argtypes = [
            ctypes.c_long,
            JOBJECT,
            ctypes.c_int,
        ]
        dll.addAccessibleSelectionFromContext.restype = None

    if hasattr(dll, "removeAccessibleSelectionFromContext"):
        dll.removeAccessibleSelectionFromContext.argtypes = [
            ctypes.c_long,
            JOBJECT,
            ctypes.c_int,
        ]
        dll.removeAccessibleSelectionFromContext.restype = None

    if hasattr(dll, "clearAccessibleSelectionFromContext"):
        dll.clearAccessibleSelectionFromContext.argtypes = [ctypes.c_long, JOBJECT]
        dll.clearAccessibleSelectionFromContext.restype = None

    if hasattr(dll, "selectAllAccessibleSelectionFromContext"):
        dll.selectAllAccessibleSelectionFromContext.argtypes = [ctypes.c_long, JOBJECT]
        dll.selectAllAccessibleSelectionFromContext.restype = None

    if hasattr(dll, "releaseJavaObject"):
        dll.releaseJavaObject.argtypes = [ctypes.c_long, JOBJECT]
        dll.releaseJavaObject.restype = None


def enum_windows(include_children=False):
    user32 = ctypes.WinDLL("user32", use_last_error=True)

    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    windows = []
    seen = set()

    def add_window(hwnd):
        if hwnd_int(hwnd) in seen:
            return
        seen.add(hwnd_int(hwnd))

        length = user32.GetWindowTextLengthW(hwnd)
        title_buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, title_buffer, length + 1)

        class_buffer = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_buffer, 256)

        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

        title = title_buffer.value
        class_name = class_buffer.value
        if title or class_name:
            windows.append(
                (hwnd, title, class_name, pid.value, bool(user32.IsWindowVisible(hwnd)))
            )

    def child_callback(hwnd, _):
        add_window(hwnd)
        return True

    def callback(hwnd, _):
        add_window(hwnd)

        if include_children:
            user32.EnumChildWindows(hwnd, EnumWindowsProc(child_callback), 0)

        return True

    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return windows


def hwnd_int(hwnd):
    return int(getattr(hwnd, "value", hwnd))


def run_message_pump(stop_event):
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    msg = wintypes.MSG()
    while not stop_event.is_set():
        while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        time.sleep(0.01)


def run_windows_access_bridge(dll, stop_event):
    dll.Windows_run()
    run_message_pump(stop_event)


def print_context_tree(
    dll, vm_id, context, depth=0, max_depth=4, max_children=40, query=None, path="0"
):
    info = AccessibleContextInfo()
    if not dll.getAccessibleContextInfo(vm_id, context, ctypes.byref(info)):
        print("  " * depth + "<failed to read context info>")
        return

    indent = "  " * depth
    bounds = f"{info.x},{info.y},{info.width},{info.height}"
    name = info.name.strip()
    desc = info.description.strip()
    role = info.role_en_US.strip() or info.role.strip()
    states = info.states_en_US.strip() or info.states.strip()
    line = f"{indent}- path={path} role={role!r} name={name!r} desc={desc!r} states={states!r} bounds={bounds} children={info.childrenCount}"
    if not query or query.lower() in line.lower():
        print(line)

    if depth >= max_depth:
        return

    count = min(info.childrenCount, max_children)
    for index in range(count):
        child = dll.getAccessibleChildFromContext(vm_id, context, index)
        if child:
            print_context_tree(
                dll,
                vm_id,
                child,
                depth + 1,
                max_depth,
                max_children,
                query=query,
                path=f"{path}.{index}",
            )
            if hasattr(dll, "releaseJavaObject"):
                dll.releaseJavaObject(vm_id, child)

    if info.childrenCount > max_children:
        print(f"{indent}  ... {info.childrenCount - max_children} more children")


def get_context_by_path(dll, vm_id, root_context, path):
    parts = [part for part in path.split(".") if part]
    if not parts or parts[0] != "0":
        raise ValueError("Path must start with 0")

    context = root_context
    owned = []
    for part in parts[1:]:
        context = dll.getAccessibleChildFromContext(vm_id, context, int(part))
        if not context:
            return None, owned
        owned.append(context)
    return context, owned


def print_table_info(dll, vm_id, context, max_rows=5, max_cols=10):
    if not hasattr(dll, "getAccessibleTableInfo"):
        print("This Access Bridge DLL does not expose getAccessibleTableInfo.")
        return

    table_info = AccessibleTableInfo()
    if not dll.getAccessibleTableInfo(vm_id, context, ctypes.byref(table_info)):
        print("getAccessibleTableInfo failed.")
        return

    print(
        "TABLE "
        f"rows={table_info.rowCount} cols={table_info.columnCount} "
        f"accessibleContext={table_info.accessibleContext} accessibleTable={table_info.accessibleTable}"
    )

    if not hasattr(dll, "getAccessibleTableCellInfo"):
        return

    row_limit = min(table_info.rowCount, max_rows)
    col_limit = min(table_info.columnCount, max_cols)
    for row in range(row_limit):
        for col in range(col_limit):
            cell_info = AccessibleTableCellInfo()
            if not dll.getAccessibleTableCellInfo(
                vm_id, context, row, col, ctypes.byref(cell_info)
            ):
                print(f"  cell[{row},{col}] <failed>")
                continue

            cell_context = cell_info.accessibleContext
            if cell_context:
                info = AccessibleContextInfo()
                if dll.getAccessibleContextInfo(
                    vm_id, cell_context, ctypes.byref(info)
                ):
                    name = info.name.strip()
                    desc = info.description.strip()
                    role = info.role_en_US.strip() or info.role.strip()
                    bounds = f"{info.x},{info.y},{info.width},{info.height}"
                    print(
                        f"  cell[{row},{col}] index={cell_info.index} role={role!r} "
                        f"name={name!r} desc={desc!r} bounds={bounds} selected={bool(cell_info.isSelected)}"
                    )
                else:
                    print(
                        f"  cell[{row},{col}] context={cell_context} <context info failed>"
                    )
            else:
                print(f"  cell[{row},{col}] <no accessible context>")


def print_actions(dll, vm_id, context):
    if not hasattr(dll, "getAccessibleActions"):
        print("This Access Bridge DLL does not expose getAccessibleActions.")
        return []

    actions = AccessibleActions()
    if not dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
        print("getAccessibleActions failed.")
        return []

    names = [
        actions.actionInfo[index].name.strip() for index in range(actions.actionsCount)
    ]
    print(f"ACTIONS count={actions.actionsCount} names={names!r}")
    return names


def iter_action_contexts(
    dll,
    vm_id,
    context,
    depth=0,
    max_depth=4,
    max_children=40,
    path="0",
):
    info = AccessibleContextInfo()
    if not dll.getAccessibleContextInfo(vm_id, context, ctypes.byref(info)):
        return

    role = info.role_en_US.strip() or info.role.strip()
    states = info.states_en_US.strip() or info.states.strip()
    name = info.name.strip()
    desc = info.description.strip()
    bounds = f"{info.x},{info.y},{info.width},{info.height}"
    action_names = []
    if hasattr(dll, "getAccessibleActions") and info.accessibleAction:
        actions = AccessibleActions()
        if dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
            action_names = [
                actions.actionInfo[index].name.strip()
                for index in range(actions.actionsCount)
            ]

    if action_names:
        yield {
            "path": path,
            "role": role,
            "name": name,
            "desc": desc,
            "states": states,
            "bounds": bounds,
            "children": info.childrenCount,
            "actions": action_names,
        }

    if depth >= max_depth:
        return

    count = min(info.childrenCount, max_children)
    for index in range(count):
        child = dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        yield from iter_action_contexts(
            dll,
            vm_id,
            child,
            depth=depth + 1,
            max_depth=max_depth,
            max_children=max_children,
            path=f"{path}.{index}",
        )
        if hasattr(dll, "releaseJavaObject"):
            dll.releaseJavaObject(vm_id, child)


def print_action_contexts(dll, vm_id, context, max_depth=4, max_children=40):
    count = 0
    for item in iter_action_contexts(
        dll,
        vm_id,
        context,
        max_depth=max_depth,
        max_children=max_children,
    ):
        count += 1
        print(
            "ACTIONABLE "
            f"path={item['path']} role={item['role']!r} name={item['name']!r} "
            f"desc={item['desc']!r} states={item['states']!r} "
            f"bounds={item['bounds']} children={item['children']} "
            f"actions={item['actions']!r}"
        )
    print(f"ACTIONABLE total={count}")


def do_action(dll, vm_id, context, action_name=None):
    if not context_is_interactable(dll, vm_id, context):
        print("DO_ACTION refused: target is not showing or has invalid bounds.")
        return False

    return do_action_raw(dll, vm_id, context, action_name)


def do_action_raw(dll, vm_id, context, action_name=None):
    names = print_actions(dll, vm_id, context)
    if not names:
        return False

    chosen = action_name or names[0]
    if chosen not in names:
        print(f"Requested action {chosen!r} not found.")
        return False

    if not hasattr(dll, "doAccessibleActions"):
        print("This Access Bridge DLL does not expose doAccessibleActions.")
        return False

    todo = AccessibleActionsToDo()
    todo.actionsCount = 1
    todo.actions[0].name = chosen
    failure = ctypes.c_int(-1)
    ok = dll.doAccessibleActions(
        vm_id, context, ctypes.byref(todo), ctypes.byref(failure)
    )
    print(f"DO_ACTION action={chosen!r} ok={bool(ok)} failure={failure.value}")
    return bool(ok)


def context_is_interactable(dll, vm_id, context):
    info = AccessibleContextInfo()
    if not dll.getAccessibleContextInfo(vm_id, context, ctypes.byref(info)):
        return False
    states = (info.states_en_US.strip() or info.states.strip()).lower()
    return (
        "visible" in states
        and "showing" in states
        and info.x >= 0
        and info.y >= 0
        and info.width > 0
        and info.height > 0
    )


def print_selection_info(dll, vm_id, context, max_children=20):
    info = AccessibleContextInfo()
    if not dll.getAccessibleContextInfo(vm_id, context, ctypes.byref(info)):
        print("getAccessibleContextInfo failed.")
        return

    role = info.role_en_US.strip() or info.role.strip()
    print(
        "SELECTION "
        f"role={role!r} name={info.name.strip()!r} "
        f"accessibleSelection={bool(info.accessibleSelection)} children={info.childrenCount}"
    )

    if not hasattr(dll, "getAccessibleSelectionCountFromContext"):
        print("This Access Bridge DLL does not expose selection APIs.")
        return

    count = dll.getAccessibleSelectionCountFromContext(vm_id, context)
    print(f"SELECTION selected_count={count}")

    if hasattr(dll, "isAccessibleChildSelectedFromContext"):
        limit = min(info.childrenCount, max_children)
        selected = []
        for index in range(limit):
            if dll.isAccessibleChildSelectedFromContext(vm_id, context, index):
                selected.append(index)
        print(f"SELECTION selected_child_indexes_first_{limit}={selected}")


def do_selection_change(dll, vm_id, context, args):
    if args.clear_selection_path:
        if not hasattr(dll, "clearAccessibleSelectionFromContext"):
            print(
                "This Access Bridge DLL does not expose clearAccessibleSelectionFromContext."
            )
            return False
        dll.clearAccessibleSelectionFromContext(vm_id, context)
        print("CLEAR_SELECTION done")
        return True

    if args.select_all_path:
        if not hasattr(dll, "selectAllAccessibleSelectionFromContext"):
            print(
                "This Access Bridge DLL does not expose selectAllAccessibleSelectionFromContext."
            )
            return False
        dll.selectAllAccessibleSelectionFromContext(vm_id, context)
        print("SELECT_ALL done")
        return True

    if args.select_child_index is not None:
        if not hasattr(dll, "addAccessibleSelectionFromContext"):
            print(
                "This Access Bridge DLL does not expose addAccessibleSelectionFromContext."
            )
            return False
        dll.addAccessibleSelectionFromContext(vm_id, context, args.select_child_index)
        print(f"ADD_SELECTION child_index={args.select_child_index} done")
        return True

    if args.remove_child_index is not None:
        if not hasattr(dll, "removeAccessibleSelectionFromContext"):
            print(
                "This Access Bridge DLL does not expose removeAccessibleSelectionFromContext."
            )
            return False
        dll.removeAccessibleSelectionFromContext(
            vm_id, context, args.remove_child_index
        )
        print(f"REMOVE_SELECTION child_index={args.remove_child_index} done")
        return True

    return False


def _build_parser():
    parser = argparse.ArgumentParser(
        description="Probe Java Access Bridge windows and controls."
    )
    parser.add_argument(
        "--title",
        default="NC",
        help="Only inspect windows whose title contains this text.",
    )
    parser.add_argument(
        "--all", action="store_true", help="Inspect all visible Java windows."
    )
    parser.add_argument(
        "--children", action="store_true", help="Also inspect child and hidden windows."
    )
    parser.add_argument(
        "--hwnd", type=int, action="append", help="Inspect a specific window handle."
    )
    parser.add_argument(
        "--dll",
        help="Full path to WindowsAccessBridge-64.dll or WindowsAccessBridge-32.dll.",
    )
    parser.add_argument("--depth", type=int, default=4, help="Control tree depth.")
    parser.add_argument(
        "--max-children", type=int, default=40, help="Max children per node."
    )
    parser.add_argument(
        "--startup-wait",
        type=float,
        default=3.0,
        help="Seconds to wait after initializing Access Bridge.",
    )
    parser.add_argument(
        "--query", help="Only print control lines containing this text."
    )
    parser.add_argument(
        "--inspect-path",
        help="Inspect one context path from each matching Java window.",
    )
    parser.add_argument(
        "--actions-path",
        help="List actions for one context path from each matching Java window.",
    )
    parser.add_argument(
        "--dump-actions",
        action="store_true",
        help="List every context with exposed actions without executing anything.",
    )
    parser.add_argument(
        "--do-action-path",
        help="Execute an action for one context path from each matching Java window.",
    )
    parser.add_argument(
        "--allow-hidden-action",
        action="store_true",
        help="Allow --do-action-path on non-showing or invalid-bounds contexts.",
    )
    parser.add_argument(
        "--action", help="Action name to execute. Defaults to the first exposed action."
    )
    parser.add_argument(
        "--selection-path",
        help="Inspect selection state for one context path from each matching Java window.",
    )
    parser.add_argument(
        "--select-child-index",
        type=int,
        help="Add one child index to the selection for --selection-path.",
    )
    parser.add_argument(
        "--remove-child-index",
        type=int,
        help="Remove one child index from the selection for --selection-path.",
    )
    parser.add_argument(
        "--clear-selection-path",
        help="Clear selection for one context path from each matching Java window.",
    )
    parser.add_argument(
        "--select-all-path",
        help="Select all children for one context path from each matching Java window.",
    )
    return parser


def _start_access_bridge(dll, stop_pump):
    if hasattr(dll, "initializeAccessBridge"):
        if not dll.initializeAccessBridge():
            print(
                "initializeAccessBridge returned false. Try `jabswitch -enable`, then restart NC and rerun."
            )
            return None, False
        return None, True

    pump_thread = threading.Thread(
        target=run_windows_access_bridge, args=(dll, stop_pump), daemon=True
    )
    pump_thread.start()
    return pump_thread, True


def _describe_window(value):
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    hwnd = wintypes.HWND(value)
    length = user32.GetWindowTextLengthW(hwnd)
    title_buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, title_buffer, length + 1)
    class_buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_buffer, 256)
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return (
        hwnd,
        title_buffer.value,
        class_buffer.value,
        pid.value,
        bool(user32.IsWindowVisible(hwnd)),
    )


def _select_windows(args):
    windows = enum_windows(include_children=args.children)
    if args.hwnd:
        known = {hwnd_int(hwnd) for hwnd, *_ in windows}
        for value in args.hwnd:
            if value not in known:
                windows.append(_describe_window(value))
    return windows


def _dispatch_path_action(dll, vm_id, context, args):
    action_path = args.actions_path or args.do_action_path
    selection_path = (
        args.selection_path or args.clear_selection_path or args.select_all_path
    )
    if not (args.inspect_path or action_path or selection_path):
        return False

    target_path = args.inspect_path or action_path or selection_path
    target_context, owned_contexts = get_context_by_path(
        dll, vm_id, context, target_path
    )
    if not target_context:
        print(f"Path not found: {target_path}")
        return True

    print(f"Inspecting path {target_path}:")
    print_context_tree(
        dll,
        vm_id,
        target_context,
        max_depth=args.depth,
        max_children=args.max_children,
    )
    print_table_info(dll, vm_id, target_context)
    if args.actions_path:
        print_actions(dll, vm_id, target_context)
    if args.do_action_path:
        if args.allow_hidden_action:
            print_actions(dll, vm_id, target_context)
            do_action_raw(dll, vm_id, target_context, args.action)
        else:
            do_action(dll, vm_id, target_context, args.action)
    if selection_path:
        print_selection_info(
            dll, vm_id, target_context, max_children=args.max_children
        )
        if (
            args.select_child_index is not None
            or args.remove_child_index is not None
            or args.clear_selection_path
            or args.select_all_path
        ):
            do_selection_change(dll, vm_id, target_context, args)
            print_selection_info(
                dll,
                vm_id,
                target_context,
                max_children=args.max_children,
            )
    for owned_context in owned_contexts:
        if hasattr(dll, "releaseJavaObject"):
            dll.releaseJavaObject(vm_id, owned_context)
    return True


def _inspect_window(dll, hwnd, title, class_name, pid, visible, args):
    print("=" * 80)
    print(
        f"hwnd={hwnd_int(hwnd)} pid={pid} visible={visible} class={class_name!r} title={title!r}"
    )

    vm_id = ctypes.c_long()
    context = JOBJECT()
    if not dll.getAccessibleContextFromHWND(
        hwnd, ctypes.byref(vm_id), ctypes.byref(context)
    ):
        print("Could not get accessible context from this window.")
        return

    if _dispatch_path_action(dll, vm_id.value, context.value, args):
        return

    if args.dump_actions:
        print_action_contexts(
            dll,
            vm_id.value,
            context.value,
            max_depth=args.depth,
            max_children=args.max_children,
        )
        return

    print_context_tree(
        dll,
        vm_id.value,
        context.value,
        max_depth=args.depth,
        max_children=args.max_children,
        query=args.query,
    )


def _inspect_windows(dll, windows, args):
    matched = 0
    total_windows = 0
    java_candidates = []
    hwnd_filter = {int(value) for value in args.hwnd or []}
    for hwnd, title, class_name, pid, visible in windows:
        total_windows += 1
        is_java = bool(dll.isJavaWindow(hwnd))
        title_match = args.title.lower() in title.lower() if args.title else True
        hwnd_match = hwnd_int(hwnd) in hwnd_filter
        if is_java:
            java_candidates.append((hwnd, title, class_name, pid, visible))
        if not is_java:
            continue
        if not args.all and not title_match and not hwnd_match:
            continue

        matched += 1
        _inspect_window(dll, hwnd, title, class_name, pid, visible, args)

    if matched == 0:
        print("No matching Java windows found.")
        print(f"Scanned windows: {total_windows}")
        if java_candidates:
            print("Java windows detected but filtered out:")
            for hwnd, title, class_name, pid, visible in java_candidates[:50]:
                print(
                    f"  hwnd={hwnd_int(hwnd)} pid={pid} visible={visible} class={class_name!r} title={title!r}"
                )
        print(
            "Checks: NC is open, Java Access Bridge is enabled, and this script uses the same bitness as Java."
        )


def main():
    args = _build_parser().parse_args()

    if os.name != "nt":
        print("This script must run with Windows Python, not WSL/Linux Python.")
        return 2

    dll, path = load_access_bridge(args.dll)
    configure_jab(dll)

    stop_pump = threading.Event()
    pump_thread, started = _start_access_bridge(dll, stop_pump)
    if not started:
        return 2

    time.sleep(args.startup_wait)
    print(f"Loaded: {path}")
    print("Visible Java windows:")

    _inspect_windows(dll, _select_windows(args), args)

    stop_pump.set()
    if pump_thread:
        pump_thread.join(timeout=1)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
