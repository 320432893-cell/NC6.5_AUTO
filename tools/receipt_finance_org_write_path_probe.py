import argparse
import ctypes
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.paths import logs_dir  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    HEADER_DYNAMIC_PREFIX_BASE,
    HEADER_SCOPE_ANCHOR_TEXT,
    find_context_by_path_readonly,
    find_context_with_window,
    find_finance_org_header_scope_by_paths,
)


WRITE_PREFIX_BASE = "0.0.0.0.1.0.0.0.0"


def main():
    parser = argparse.ArgumentParser(
        description="只读探测收款单财务组织可写控件 path 稳定性"
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--hwnd", type=int, default=None)
    parser.add_argument("--min-index", type=int, default=1)
    parser.add_argument("--max-index", type=int, default=10)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        jab.ensure_started()
        report = probe(jab, args)
    finally:
        jab.close()

    output_path = (
        logs_dir()
        / f"receipt_finance_org_write_path_probe_{time.strftime('%Y%m%d_%H%M%S')}.json"
    )
    output_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"探测详情: {output_path}")
        print_summary(report)


def probe(jab, args):
    scope = find_finance_org_header_scope_by_paths(
        jab,
        args.hwnd,
        min_index=args.min_index,
        max_index=args.max_index,
    )
    scope_hwnd = (
        scope.get("scope_hwnd")
        or ((scope.get("window") or {}).get("hwnd"))
        or args.hwnd
    )
    legacy = find_legacy_finance_org_text_path(jab, scope_hwnd=scope_hwnd)
    suffix_info = infer_write_path_suffix(legacy.get("path"))
    candidates = []
    suffix = suffix_info.get("suffix")
    if suffix:
        for index in range(int(args.min_index), int(args.max_index) + 1):
            path = build_write_path(index, suffix)
            readonly = find_context_by_path_readonly(
                jab,
                path,
                scope_hwnd=scope_hwnd,
                role="text",
            )
            candidates.append(
                {
                    "dynamic_index": index,
                    "path": path,
                    "ok": bool(readonly.get("ok")),
                    "name": readonly.get("name"),
                    "description": readonly.get("description"),
                    "text": readonly.get("text"),
                    "role": readonly.get("role"),
                    "states": readonly.get("states"),
                }
            )
    return {
        "scope": redact_scope(scope),
        "scope_hwnd": scope_hwnd,
        "legacy_write": legacy,
        "write_suffix": suffix_info,
        "write_candidates": candidates,
    }


def find_legacy_finance_org_text_path(jab, scope_hwnd=None):
    started_at = time.perf_counter()
    context, vm_id, owned_contexts, owned_indexes, window_info = (
        find_context_with_window(
            jab,
            HEADER_SCOPE_ANCHOR_TEXT,
            roles=("text",),
            timeout=1.5,
            require_showing=True,
            window_class="SunAwtCanvas",
            visible_only=True,
            scope_hwnd=scope_hwnd,
        )
    )
    if not context:
        return {
            "ok": False,
            "reason": "legacy control-name text not found",
            "scope_hwnd": scope_hwnd,
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    path = "0" + "".join(f".{index}" for index in owned_indexes)
    try:
        info = jab.get_context_info(vm_id, context)
        text = jab.get_text_context_value(vm_id, context)
        return {
            "ok": True,
            "path": path,
            "owned_indexes": list(owned_indexes),
            "window": window_info,
            "scope_hwnd": scope_hwnd,
            "name": str(getattr(info, "name", "") or "").strip() if info else "",
            "description": str(getattr(info, "description", "") or "").strip()
            if info
            else "",
            "text": text,
            "role": (
                str(getattr(info, "role_en_US", "") or "").strip()
                or str(getattr(info, "role", "") or "").strip()
            )
            if info
            else "",
            "states": (
                str(getattr(info, "states_en_US", "") or "").strip()
                or str(getattr(info, "states", "") or "").strip()
            )
            if info
            else "",
            "seconds": round(time.perf_counter() - started_at, 3),
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def infer_write_path_suffix(path):
    text = str(path or "").strip()
    prefix = f"{WRITE_PREFIX_BASE}."
    if not text.startswith(prefix):
        return {
            "ok": False,
            "reason": "write path does not match expected prefix",
            "prefix_base": WRITE_PREFIX_BASE,
            "path": text,
        }
    rest = text[len(prefix) :]
    index, sep, suffix = rest.partition(".")
    if not sep:
        return {
            "ok": False,
            "reason": "write path missing suffix after dynamic index",
            "prefix_base": WRITE_PREFIX_BASE,
            "path": text,
        }
    try:
        dynamic_index = int(index)
    except ValueError:
        return {
            "ok": False,
            "reason": "write dynamic index is not integer",
            "prefix_base": WRITE_PREFIX_BASE,
            "path": text,
        }
    return {
        "ok": True,
        "prefix_base": WRITE_PREFIX_BASE,
        "dynamic_index": dynamic_index,
        "suffix": suffix,
        "path": text,
        "header_prefix_base": HEADER_DYNAMIC_PREFIX_BASE,
    }


def build_write_path(dynamic_index, suffix):
    return f"{WRITE_PREFIX_BASE}.{int(dynamic_index)}.{suffix}"


def redact_scope(scope):
    if not isinstance(scope, dict):
        return scope
    return {
        key: value
        for key, value in scope.items()
        if key not in {"context", "owned_contexts", "vm_id"}
    }


def print_summary(report):
    scope = report.get("scope") or {}
    legacy = report.get("legacy_write") or {}
    suffix = report.get("write_suffix") or {}
    print(
        "scope: "
        f"ok={scope.get('ok')} dynamic_index={scope.get('dynamic_index')} "
        f"hwnd={report.get('scope_hwnd')} seconds={scope.get('seconds')}"
    )
    print(
        "legacy-write: "
        f"ok={legacy.get('ok')} path={legacy.get('path')} "
        f"seconds={legacy.get('seconds')}"
    )
    print(
        "write-suffix: "
        f"ok={suffix.get('ok')} dynamic_index={suffix.get('dynamic_index')} "
        f"suffix={suffix.get('suffix')}"
    )
    for item in report.get("write_candidates") or []:
        if not item.get("ok"):
            continue
        print(
            f"- index {item.get('dynamic_index')}: ok=True "
            f"name={item.get('name')!r} desc={item.get('description')!r} "
            f"text={item.get('text')!r} role={item.get('role')!r} "
            f"states={item.get('states')!r}"
        )
        print(f"  path={item.get('path')}")


if __name__ == "__main__":
    main()
