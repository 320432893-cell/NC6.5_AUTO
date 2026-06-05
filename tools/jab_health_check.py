# T0 temporary helper. Delete after NC/JAB startup health checks move into core flow.

import ctypes

from tools.jab_probe import JOBJECT, enum_windows


def check_jab_ready(jab):
    """Read-only guard: at least one visible SunAwt window must expose JAB context."""
    jab.ensure_started()
    windows = []
    ready_windows = []
    visible_sunawt = []

    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if not class_name.startswith("SunAwt"):
            continue
        item = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": visible,
            "is_java": bool(jab.dll.isJavaWindow(hwnd)),
            "get_context_ok": False,
            "root_role": None,
            "root_children": None,
        }
        if visible:
            visible_sunawt.append(item)

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if jab.dll.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vm_id), ctypes.byref(root_context)
        ):
            item["get_context_ok"] = True
            info = jab.get_context_info(vm_id.value, root_context.value)
            if info:
                item["root_role"] = info.role_en_US.strip() or info.role.strip()
                item["root_children"] = info.childrenCount
            jab.release_contexts(vm_id.value, [root_context.value])

        if visible and item["is_java"] and item["get_context_ok"]:
            ready_windows.append(item)
        windows.append(item)

    ok = bool(ready_windows)
    return {
        "ok": ok,
        "reason": None if ok else build_health_reason(visible_sunawt, windows),
        "ready_windows": ready_windows,
        "visible_sunawt": visible_sunawt,
        "sunawt_count": len(windows),
    }


def build_health_reason(visible_sunawt, windows):
    if not windows:
        return "未发现 SunAwt Java 窗口，NC Java 界面可能未打开。"
    if not visible_sunawt:
        return "发现 SunAwt 窗口但没有可见 Java 窗口，当前 NC 页面可能未显示。"
    return (
        "当前 NC Java 窗口未注册到 Java Access Bridge："
        "SunAwt 窗口存在，但 isJava/getContext 均未就绪。"
    )


def print_jab_health_failure(health):
    print("JAB 健康检查：失败")
    print(f"原因：{health.get('reason')}")
    visible = health.get("visible_sunawt") or []
    if visible:
        print("可见 SunAwt 窗口：")
        for item in visible[:5]:
            print(
                "  "
                f"hwnd={item.get('hwnd')} class={item.get('class_name')} "
                f"title={item.get('title') or '<无标题>'} "
                f"isJava={item.get('is_java')} getContext={item.get('get_context_ok')}"
            )
    print("本次不会继续输入、点击、写明细、保存或暂存。")
