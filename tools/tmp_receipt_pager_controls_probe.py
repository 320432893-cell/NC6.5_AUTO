# 职责: T0 只读扫描当前 NC 页面中疑似分页控件（页码标签、每页行数文本框、下一页按钮）的 JAB path
# 不做什么: 不点击、不写文本、不改分页、不保存、不写 Excel
# 允许依赖层: core JAB/config、tools.jab_probe
# 谁不应该 import: 正式流程、core 模块和测试不应 import 本临时探针
# 生命周期: T0 临时探针（删除条件：当前查询结果页分页控件 path 后缀确认并沉淀到正式模块/文档）

import argparse
import ctypes
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_probe import JOBJECT  # noqa: E402


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--max-items", type=int, default=120)
    args = parser.parse_args(argv)

    config = load_config(args.config)
    jab = JABOperator(config)
    try:
        jab.ensure_started()
        report = scan_pager_controls(jab, args.max_items)
    finally:
        jab.close()
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    return 0


def scan_pager_controls(jab, max_items):
    items = []
    for hwnd, title, class_name, pid, visible in jab.get_scoped_windows(
        include_children=True
    ):
        if (
            class_name != "SunAwtCanvas"
            or not visible
            or not jab.dll.isJavaWindow(hwnd)
        ):
            continue
        vm_id_ref = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id_ref),
            ctypes.byref(root_context),
        ):
            continue
        owned = [root_context.value]
        try:
            collect_controls(
                jab,
                vm_id_ref.value,
                root_context.value,
                path=[0],
                owned=owned,
                result=items,
                max_items=max_items,
            )
        finally:
            jab.release_contexts(vm_id_ref.value, list(dict.fromkeys(owned)))
    return {"count": len(items), "items": items[:max_items]}


def collect_controls(jab, vm_id, context, path, owned, result, max_items, depth=0):
    if len(result) >= max_items or depth > jab.max_depth:
        return
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    name = info.name.strip()
    desc = info.description.strip()
    role = (info.role_en_US.strip() or info.role.strip()).lower()
    text = desc or name
    path_text = ".".join(str(part) for part in path)
    if is_pager_candidate(role, text, path_text):
        result.append(
            {
                "path": path_text,
                "role": role,
                "name": name,
                "description": desc,
                "x": info.x,
                "y": info.y,
                "width": info.width,
                "height": info.height,
                "states": info.states_en_US.strip() or info.states.strip(),
            }
        )
    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        owned.append(child)
        collect_controls(
            jab,
            vm_id,
            child,
            path + [index],
            owned,
            result,
            max_items,
            depth + 1,
        )


def is_pager_candidate(role, text, path):
    if "页" in text and ("每页" in text or "记录" in text):
        return True
    if role in {"text", "push button", "label"} and path.startswith("0.0.1.0.0.0.0"):
        if text in {"10", "50", "100", "500", "下一页", "下页", ">"}:
            return True
    return False


if __name__ == "__main__":
    raise SystemExit(main())
