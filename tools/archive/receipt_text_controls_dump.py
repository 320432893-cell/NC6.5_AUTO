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
from tools.jab_probe import JOBJECT, enum_windows  # noqa: E402


def main():
    parser = argparse.ArgumentParser(
        description="Dump current visible text controls via configured JAB."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--class-name", default="SunAwtCanvas")
    parser.add_argument("--query", default=None)
    parser.add_argument("--all-controls", action="store_true")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        jab.ensure_started()
        controls = collect_text_controls(
            jab, args.class_name, args.query, args.all_controls
        )
    finally:
        jab.close()

    if args.json:
        print(json.dumps(controls, ensure_ascii=False, indent=2))
    else:
        for item in controls:
            print(
                f"path={item['path']} role={item['role']!r} name={item['name']!r} "
                f"desc={item['description']!r} states={item['states']!r} "
                f"value={item['value']!r} bounds={item['bounds']} window={item['window']}"
            )
    return 0


def collect_text_controls(jab, class_name, query, all_controls=False):
    result = []
    query_l = str(query).lower() if query else None
    for hwnd, title, window_class, pid, visible in enum_windows(include_children=True):
        if class_name and window_class != class_name:
            continue
        if not visible or not jab.dll.isJavaWindow(hwnd):
            continue
        vm_id = ctypes.c_long()
        root = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vm_id), ctypes.byref(root)
        ):
            continue
        collect_tree(
            jab,
            vm_id.value,
            root.value,
            "0",
            {"hwnd": int(hwnd), "title": title, "class": window_class, "pid": pid},
            result,
            query_l,
            all_controls,
            depth=0,
        )
    return result


def collect_tree(
    jab, vm_id, context, path, window, result, query_l, all_controls, depth
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    role = info.role_en_US.strip() or info.role.strip()
    states = info.states_en_US.strip() or info.states.strip()
    text = f"{info.name} {info.description} {role} {states}".lower()
    if (all_controls or role.lower() == "text" or info.accessibleText) and (
        not query_l or query_l in text
    ):
        result.append(
            {
                "path": path,
                "role": role,
                "name": info.name.strip(),
                "description": info.description.strip(),
                "states": states,
                "value": jab.get_text_context_value(vm_id, context),
                "bounds": [info.x, info.y, info.width, info.height],
                "window": window,
            }
        )
    if depth >= jab.max_depth or role.lower() == "table":
        return
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        collect_tree(
            jab,
            vm_id,
            child,
            f"{path}.{index}",
            window,
            result,
            query_l,
            all_controls,
            depth + 1,
        )
        jab.release_contexts(vm_id, [child])


if __name__ == "__main__":
    raise SystemExit(main())
