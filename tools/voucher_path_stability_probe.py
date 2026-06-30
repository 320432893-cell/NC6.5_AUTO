import argparse
import json
import sys
import time
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.nc_page_probe import NCPageProbe  # noqa: E402
from core.paths import logs_dir  # noqa: E402
from core.utils import load_config  # noqa: E402


TRACK_NAMES = (
    "单据生成",
    "删除",
    "查询",
    "刷新",
    "选择",
    "生成",
    "正式单据",
    "确定",
    "保存(Ctrl+S)",
)


def main():
    parser = argparse.ArgumentParser(description="只读采样凭证关键控件 path 稳定性")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--samples", type=int, default=4)
    parser.add_argument("--interval", type=float, default=0.8)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    probe = NCPageProbe(jab, cfg.get("jab_batch", {}))
    samples = []
    try:
        for index in range(args.samples):
            jab.ensure_started()
            controls = [
                control
                for control in probe.collect_controls()
                if control.get("name") in TRACK_NAMES
            ]
            toolbar = probe.detect_pending_toolbar(controls)
            samples.append(
                {
                    "index": index + 1,
                    "controls": controls,
                    "pending_toolbar": toolbar,
                }
            )
            if index + 1 < args.samples:
                time.sleep(args.interval)
    finally:
        jab.close()

    report = {
        "samples": samples,
        "summary": summarize(samples),
    }
    output_path = (
        logs_dir() / f"voucher_path_stability_{time.strftime('%Y%m%d_%H%M%S')}.json"
    )
    output_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report, output_path)


def summarize(samples):
    by_name = defaultdict(list)
    for sample in samples:
        controls = sample.get("controls") or []
        for name in TRACK_NAMES:
            matches = [control for control in controls if control.get("name") == name]
            by_name[name].append(matches)

    result = []
    for name in TRACK_NAMES:
        per_sample = by_name[name]
        primary = [pick_primary(matches) for matches in per_sample]
        found = [item for item in primary if item]
        paths = [item.get("path") for item in found]
        bounds = [tuple(item.get("bounds") or []) for item in found]
        parent_paths = [parent_path(path) for path in paths if path]
        result.append(
            {
                "name": name,
                "found": len(found),
                "samples": len(samples),
                "same_path": len(set(paths)) == 1 if paths else False,
                "paths": unique(paths),
                "same_parent": len(set(parent_paths)) == 1 if parent_paths else False,
                "parent_paths": unique(parent_paths),
                "same_bounds": len(set(bounds)) == 1 if bounds else False,
                "bounds": [list(item) for item in unique(bounds)],
                "counts": [len(matches) for matches in per_sample],
            }
        )

    toolbar_ok = [
        bool((sample.get("pending_toolbar") or {}).get("ok")) for sample in samples
    ]
    return {
        "controls": result,
        "pending_toolbar_ok": toolbar_ok,
        "pending_toolbar_stable": all(toolbar_ok) if toolbar_ok else False,
        "pending_toolbar_buttons": summarize_toolbar_buttons(samples),
    }


def pick_primary(matches):
    showing = [item for item in matches if item.get("showing")]
    items = showing or matches
    if not items:
        return None
    return min(
        items, key=lambda item: (item.get("window_hwnd") or 0, item.get("path") or "")
    )


def parent_path(path):
    if not path or "." not in path:
        return ""
    return str(path).rsplit(".", 1)[0]


def summarize_toolbar_buttons(samples):
    by_name = defaultdict(list)
    for sample in samples:
        toolbar = sample.get("pending_toolbar") or {}
        window_hwnd = toolbar.get("window_hwnd")
        for button in toolbar.get("buttons") or []:
            by_name[button.get("name")].append(
                {
                    **button,
                    "window_hwnd": window_hwnd,
                }
            )

    result = []
    for name in ("删除", "查询", "刷新", "选择", "生成"):
        items = by_name.get(name) or []
        paths = [item.get("path") for item in items if item.get("path")]
        parents = [parent_path(path) for path in paths]
        bounds = [tuple(item.get("bounds") or []) for item in items]
        result.append(
            {
                "name": name,
                "found": len(items),
                "samples": len(samples),
                "same_path": len(set(paths)) == 1 if paths else False,
                "paths": unique(paths),
                "same_parent": len(set(parents)) == 1 if parents else False,
                "parent_paths": unique(parents),
                "same_bounds": len(set(bounds)) == 1 if bounds else False,
                "bounds": [list(item) for item in unique(bounds)],
                "window_hwnds": unique(
                    [
                        item.get("window_hwnd")
                        for item in items
                        if item.get("window_hwnd")
                    ]
                ),
            }
        )
    return result


def unique(values):
    result = []
    seen = set()
    for value in values:
        key = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
        if key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


def print_text(report, output_path):
    summary = report["summary"]
    print(f"采样详情: {output_path}")
    print(f"待生成工具栏通过: {summary['pending_toolbar_ok']}")
    print(f"待生成工具栏稳定: {summary['pending_toolbar_stable']}")
    print("待生成工具栏选中按钮:")
    for item in summary.get("pending_toolbar_buttons") or []:
        print(
            "- {name}: found={found}/{samples} same_path={same_path} "
            "same_parent={same_parent} same_bounds={same_bounds}".format(**item)
        )
        print(f"  paths={item['paths']}")
    print("控件稳定性:")
    for item in summary["controls"]:
        print(
            "- {name}: found={found}/{samples} same_path={same_path} "
            "same_parent={same_parent} same_bounds={same_bounds} counts={counts}".format(
                **item
            )
        )
        print(f"  paths={item['paths']}")
        if not item["same_bounds"]:
            print(f"  bounds={item['bounds']}")


if __name__ == "__main__":
    main()
