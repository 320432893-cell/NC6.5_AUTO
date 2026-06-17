# 职责：自制菜单选择与录入态检测(含自制/录入态名集常量)
# 不做什么：不做 CLI 解析/不起 JABOperator(那是 receipt_new_probe 主入口)
# 允许依赖层：标准库、core JAB(经 jab 参数)、tools.jab_probe、tools.receipt_new_* 同层
# 谁不应该 import：core 层模块不应 import

import sys




class _ProbeNamespace:
    # 调用时从已加载的 receipt_new_probe 读顶层函数,使测试对
    # tools.receipt_new_probe.<name> 的 monkeypatch 与拆分前一致生效,且不在加载期 import 成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_new_probe"], name)


_probe = _ProbeNamespace()


SELF_MADE_NAMES = {"自制"}
ENTRY_STATE_NAMES = {"保存(Ctrl+S)", "暂存", "取消(Ctrl+Q)"}


def choose_self_made_menu_item(jab, windows, fallback_index, popup_hwnd=None):
    candidates = []
    for window in windows:
        if not window.get("is_java"):
            continue
        if not window.get("visible"):
            continue
        if popup_hwnd is not None and window.get("hwnd") != popup_hwnd:
            continue
        for control in window.get("all_controls", []):
            if not _probe.is_current_visible_control(control):
                continue
            if not control.get("accessibleAction"):
                continue
            if (
                control["role"].lower() == "menu item"
                or control.get("name") in SELF_MADE_NAMES
            ):
                candidates.append({"window": window, "control": control})

    named = [
        item for item in candidates if item["control"].get("name") in SELF_MADE_NAMES
    ]
    if named:
        target = named[0]
    elif fallback_index is not None and 0 <= fallback_index < len(candidates):
        target = candidates[fallback_index]
    else:
        return {
            "ok": False,
            "reason": "self-made menu item not found",
            "candidate_count": len(candidates),
            "candidates": _probe.summarize_candidates(candidates),
        }

    ok = _probe.do_action_by_window_path(
        jab,
        target["window"]["hwnd"],
        target["control"]["path"],
        action_name=_probe.choose_click_action(target["control"].get("actions", [])),
    )
    return {
        "ok": bool(ok),
        "target": {
            "window": {
                key: target["window"].get(key)
                for key in ("hwnd", "title", "class_name", "visible")
            },
            "control": target["control"],
        },
        "candidate_count": len(candidates),
        "candidates": _probe.summarize_candidates(candidates),
    }



def is_current_visible_control(control):
    states = control.get("states", "")
    bounds = control.get("bounds") or []
    normalized_states = states.lower()
    if "visible" not in normalized_states or "showing" not in normalized_states:
        return False
    return (
        isinstance(bounds, list)
        and len(bounds) == 4
        and bounds[2] > 0
        and bounds[3] > 0
    )



def choose_click_action(actions):
    if not actions:
        return None
    for preferred in ("单击", "click", "press"):
        if preferred in actions:
            return preferred
    return actions[0]



def do_action_by_window_path(jab, hwnd, path, action_name=None):
    result = jab.find_context_by_path_once(
        path,
        scope_hwnd=hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    context, vm_id, owned, _window_info = result
    if not context:
        return False
    try:
        return jab.do_action(
            vm_id,
            context,
            action_name=action_name,
            cleanup_blank_awt=False,
        )
    finally:
        jab.release_contexts(vm_id, owned)



def detect_self_made_entry_state(windows):
    names = set()
    hits = []
    for window in windows or []:
        for control in window.get("controls", []):
            matched_names = _probe.normalize_entry_state_names(control)
            if matched_names:
                names.update(matched_names)
                hits.append(
                    {
                        "window": {
                            key: window.get(key)
                            for key in ("hwnd", "class_name", "title", "visible")
                        },
                        "control": control,
                    }
                )
    return {
        "ok": ENTRY_STATE_NAMES.issubset(names),
        "partial_ok": bool(names),
        "names": sorted(names),
        "hits": hits,
    }



def normalize_entry_state_names(control):
    texts = {
        str(control.get("name") or "").strip(),
        str(control.get("description") or "").strip(),
    }
    matched = set()
    if "暂存" in texts:
        matched.add("暂存")
    if "保存(Ctrl+S)" in texts or "保存" in texts:
        matched.add("保存(Ctrl+S)")
    if "取消(Ctrl+Q)" in texts or "取消" in texts:
        matched.add("取消(Ctrl+Q)")
    return matched
