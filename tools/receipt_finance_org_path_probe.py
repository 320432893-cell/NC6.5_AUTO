import argparse
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
    FINANCE_ORG_ACCEPTED_TEXT,
    FINANCE_ORG_LABEL_SUFFIX,
    HEADER_DYNAMIC_PREFIX_BASE,
    build_receipt_header_dynamic_label_path,
    build_receipt_header_dynamic_path,
    find_context_by_path_readonly,
    infer_header_text_path_from_label_path,
    locate_receipt_header_scope,
    receipt_header_dynamic_prefix,
)

FINANCE_ORG_LABEL_SUFFIX_VARIANTS = (
    ("configured", FINANCE_ORG_LABEL_SUFFIX),
    ("observed-compact", "0.0.0.1.1.0.0.0.0.1.1.0"),
)


def main():
    parser = argparse.ArgumentParser(
        description="只读探测收款单财务组织 dynamic path 稳定性"
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
        / f"receipt_finance_org_path_probe_{time.strftime('%Y%m%d_%H%M%S')}.json"
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
    try:
        scope = locate_receipt_header_scope(jab, scope_hwnd=args.hwnd)
    except Exception as exc:
        scope = {
            "ok": False,
            "reason": f"{type(exc).__name__}: {exc}",
        }
    scope_hwnd = scope.get("scope_hwnd") or args.hwnd
    candidates = []
    for dynamic_index in range(int(args.min_index), int(args.max_index) + 1):
        for variant, label_path, text_path in build_finance_org_path_variants(
            dynamic_index
        ):
            label = find_context_by_path_readonly(
                jab,
                label_path,
                scope_hwnd=scope_hwnd,
            )
            text = find_context_by_path_readonly(
                jab,
                text_path,
                scope_hwnd=scope_hwnd,
            )
            candidates.append(
                {
                    "dynamic_index": dynamic_index,
                    "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
                    "variant": variant,
                    "label_path": label_path,
                    "text_path": text_path,
                    "label": label,
                    "text": text,
                    "label_matches": is_finance_org_label(label),
                    "text_accepted": is_finance_org_accepted(text),
                }
            )
    matches = [
        item
        for item in candidates
        if item.get("label_matches") or item.get("text", {}).get("ok")
    ]
    best = next(
        (item for item in candidates if item.get("label_matches")),
        matches[0] if matches else None,
    )
    return {
        "scope": redact_scope(scope),
        "scope_hwnd": scope_hwnd,
        "expected_accepted_text": FINANCE_ORG_ACCEPTED_TEXT,
        "candidates": candidates,
        "best_dynamic_index": best.get("dynamic_index") if best else None,
        "best_dynamic_prefix": best.get("dynamic_prefix") if best else None,
    }


def build_finance_org_path_variants(dynamic_index):
    seen = set()
    paths = []
    configured_label_path = build_receipt_header_dynamic_label_path(
        dynamic_index,
        "财务组织",
    )
    configured_text_path = build_receipt_header_dynamic_path(dynamic_index, "财务组织")
    paths.append(("configured-builder", configured_label_path, configured_text_path))
    for name, suffix in FINANCE_ORG_LABEL_SUFFIX_VARIANTS:
        label_path = f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{suffix}"
        text_path = infer_header_text_path_from_label_path("财务组织", label_path)
        paths.append((name, label_path, text_path))
    for variant, label_path, text_path in paths:
        key = (label_path, text_path)
        if not label_path or not text_path or key in seen:
            continue
        seen.add(key)
        yield variant, label_path, text_path


def is_finance_org_label(result):
    if not result.get("ok"):
        return False
    values = {
        str(result.get("name") or "").strip(),
        str(result.get("description") or "").strip(),
        str(result.get("text") or "").strip(),
    }
    return "财务组织(O)" in values


def is_finance_org_accepted(result):
    if not result.get("ok"):
        return False
    values = {
        str(result.get("name") or "").strip(),
        str(result.get("description") or "").strip(),
        str(result.get("text") or "").strip(),
    }
    return FINANCE_ORG_ACCEPTED_TEXT in values


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
    print(
        "scope: "
        f"ok={scope.get('ok')} mode={scope.get('mode')} "
        f"dynamic_index={scope.get('dynamic_index')} hwnd={report.get('scope_hwnd')}"
    )
    print(
        "best: "
        f"dynamic_index={report.get('best_dynamic_index')} "
        f"prefix={report.get('best_dynamic_prefix')}"
    )
    for item in report.get("candidates") or []:
        label = item.get("label") or {}
        text = item.get("text") or {}
        if not (
            label.get("ok")
            or text.get("ok")
            or item.get("label_matches")
            or item.get("text_accepted")
        ):
            continue
        print(
            f"- index {item.get('dynamic_index')}: "
            f"variant={item.get('variant')} "
            f"label_ok={label.get('ok')} label_match={item.get('label_matches')} "
            f"text_ok={text.get('ok')} text_accepted={item.get('text_accepted')}"
        )
        print(f"  label_path={item.get('label_path')}")
        print(f"  text_path={item.get('text_path')}")
        print(
            "  label "
            f"name={label.get('name')!r} desc={label.get('description')!r} "
            f"role={label.get('role')!r} states={label.get('states')!r}"
        )
        print(
            "  text  "
            f"name={text.get('name')!r} desc={text.get('description')!r} "
            f"text={text.get('text')!r} role={text.get('role')!r} "
            f"states={text.get('states')!r}"
        )


if __name__ == "__main__":
    main()
