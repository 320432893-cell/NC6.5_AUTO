# 职责：收款单开单探针的报告/摘要格式化与控制台打印
# 不做什么：不做 CLI 解析/不起 JABOperator(那是 receipt_new_probe 主入口)
# 允许依赖层：标准库、core JAB(经 jab 参数)、tools.jab_probe、tools.receipt_new_* 同层
# 谁不应该 import：core 层模块不应 import

import ctypes
import json
import sys

from tools.jab_probe import AccessibleActions



class _ProbeNamespace:
    # 调用时从已加载的 receipt_new_probe 读顶层函数,使测试对
    # tools.receipt_new_probe.<name> 的 monkeypatch 与拆分前一致生效,且不在加载期 import 成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_new_probe"], name)


_probe = _ProbeNamespace()


def summarize_context(jab, vm_id, context, path):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    return _probe.summarize_info(jab, vm_id, context, info, path)



def summarize_info(jab, vm_id, context, info, path):
    role = info.role_en_US.strip() or info.role.strip()
    item = {
        "path": path,
        "role": role,
        "name": info.name.strip(),
        "description": info.description.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "bounds": [info.x, info.y, info.width, info.height],
        "children": info.childrenCount,
        "accessibleAction": bool(info.accessibleAction),
        "actions": [],
    }
    if info.accessibleAction:
        item["actions"] = _probe.get_action_names(jab, vm_id, context)
    return item



def get_action_names(jab, vm_id, context):
    if not hasattr(jab.dll, "getAccessibleActions"):
        return []
    actions = AccessibleActions()
    if not jab.dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
        return []
    return [
        actions.actionInfo[index].name.strip() for index in range(actions.actionsCount)
    ]



def summarize_candidates(candidates):
    result = []
    for index, item in enumerate(candidates):
        result.append(
            {
                "index": index,
                "window": {
                    key: item["window"].get(key)
                    for key in (
                        "hwnd",
                        "class_name",
                        "title",
                        "visible",
                        "root_hwnd",
                        "is_foreground_root",
                    )
                },
                "control": item["control"],
            }
        )
    return result



def summarize_report(report):
    return {
        "foreground": report.get("foreground"),
        "matches": [_probe.summarize_target(item) for item in report.get("matches", [])[:20]],
        "buttons": [_probe.summarize_target(item) for item in report.get("buttons", [])[:20]],
        "usable_buttons": [
            _probe.summarize_target(item) for item in report.get("usable_buttons", [])[:20]
        ],
        "open": _probe.summarize_action_report(report.get("open")),
        "tracked_popup": report.get("tracked_popup"),
        "popup_cleanup": report.get("popup_cleanup"),
        "changed_windows": [
            {
                "hwnd": item.get("hwnd"),
                "class_name": item.get("class_name"),
                "title": item.get("title"),
                "visible": item.get("visible"),
                "root": item.get("root"),
                "controls": [
                    _probe.summarize_control(control)
                    for control in item.get("controls", [])[:30]
                ],
            }
            for item in report.get("new_or_changed_after_open", [])[:10]
        ],
        "choose_self_made": _probe.summarize_action_report(report.get("choose_self_made")),
        "entry_state": report.get("entry_state"),
        "timings": report.get("timings") or [],
    }



def summarize_action_report(action_report):
    if not isinstance(action_report, dict):
        return action_report
    result = {
        key: value
        for key, value in action_report.items()
        if key
        in {
            "ok",
            "method",
            "reason",
            "path",
            "candidate_count",
            "action_returned_within_timeout",
            "action_status",
            "rejected_count",
        }
    }
    if "target" in action_report:
        result["target"] = _probe.summarize_target(action_report["target"])
    if "candidates" in action_report:
        result["candidates"] = [
            _probe.summarize_target(item) for item in action_report.get("candidates", [])[:20]
        ]
    return result



def summarize_target(item):
    if not isinstance(item, dict):
        return item
    return {
        "window": {
            key: item.get("window", {}).get(key)
            for key in (
                "hwnd",
                "class_name",
                "title",
                "visible",
                "root_hwnd",
                "is_foreground_root",
            )
        },
        "control": _probe.summarize_control(item.get("control", {})),
    }



def summarize_control(control):
    return {
        key: control.get(key)
        for key in (
            "path",
            "role",
            "name",
            "description",
            "states",
            "bounds",
            "accessibleAction",
            "actions",
        )
    }



def print_text(report):
    print("open:", json.dumps(report["open"], ensure_ascii=False))
    print("new_or_changed_after_open:", len(report["new_or_changed_after_open"]))
    for window in report["new_or_changed_after_open"]:
        print(
            f"  window hwnd={window['hwnd']} class={window['class_name']!r} "
            f"title={window['title']!r} visible={window['visible']} root={window['root']}"
        )
        for control in window["controls"][:80]:
            print(
                f"    path={control['path']} role={control['role']!r} "
                f"name={control['name']!r} desc={control['description']!r} "
                f"states={control['states']!r} actions={control['actions']} "
                f"bounds={control['bounds']}"
            )
    print(
        "choose_self_made:", json.dumps(report["choose_self_made"], ensure_ascii=False)
    )
    print("entry_state:", json.dumps(report["entry_state"], ensure_ascii=False))
