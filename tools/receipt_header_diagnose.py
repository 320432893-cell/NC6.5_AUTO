# Read-only diagnostic for receipt entry header field resolution.

import argparse
import ctypes
from datetime import datetime
import json
import platform
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FROZEN_BASE = Path(getattr(sys, "_MEIPASS", ROOT))
OUTPUT_BASE = (
    Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else ROOT
)
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from core.jab_environment import prepare_java_access_bridge  # noqa: E402
from core.jab_health_check import check_jab_ready  # noqa: E402
from core.jab_probe import enum_windows  # noqa: E402
from core.receipt_new_probe import foreground_info, root_hwnd  # noqa: E402
from core.receipt_query_guard import guard_receipt_parent_page  # noqa: E402
from core.receipt_self_made_fill_trial import (  # noqa: E402
    HEADER_REQUIRED_LABELS,
    build_receipt_header_dynamic_label_path,
    build_receipt_header_dynamic_path,
    find_receipt_header_field_by_dynamic_path,
    find_receipt_header_field_by_live_semantic,
    find_receipt_header_field_by_semantic_label,
    receipt_header_dynamic_prefix,
    resolve_receipt_header_anchor_in_canvas,
)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="只读诊断当前 NC 收款单自制页表头字段定位，不写字段、不保存。"
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--json", action="store_true")
    return parser.parse_args(argv)


def resolve_config_path(raw):
    path = Path(str(raw or "config.json"))
    if path.is_absolute():
        return path
    for base in (Path.cwd(), OUTPUT_BASE, FROZEN_BASE, ROOT):
        candidate = base / path
        if candidate.exists():
            return candidate
    return path


def context_snapshot(jab, found):
    if not found or not found.get("ok"):
        return None
    context = found.get("context")
    vm_id = found.get("vm_id")
    owned = found.get("owned_contexts") or []
    if not context or vm_id is None:
        return None
    try:
        info = jab.get_context_info(vm_id, context)
        text = jab.get_text_context_value(vm_id, context)
        if not info:
            return {"text": text, "info": None}
        return {
            "text": text,
            "name": info.name.strip(),
            "description": info.description.strip(),
            "role": (info.role_en_US.strip() or info.role.strip()),
            "states": (info.states_en_US.strip() or info.states.strip()),
            "bounds": [info.x, info.y, info.width, info.height],
            "visible": "visible"
            in (info.states_en_US.strip() or info.states.strip()).lower(),
            "showing": "showing"
            in (info.states_en_US.strip() or info.states.strip()).lower(),
        }
    finally:
        jab.release_contexts(vm_id, owned)


def strip_context_handles(found):
    result = {}
    for key, value in (found or {}).items():
        if key in {"context", "vm_id", "owned_contexts"}:
            continue
        result[key] = value
    return result


def diagnose_field(jab, label, dynamic_index, scope_hwnd):
    expected_text_path = build_receipt_header_dynamic_path(dynamic_index, label)
    expected_label_path = build_receipt_header_dynamic_label_path(dynamic_index, label)
    path_found = find_receipt_header_field_by_dynamic_path(
        jab,
        label,
        dynamic_index,
        scope_hwnd=scope_hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    path_snapshot = context_snapshot(jab, path_found)
    live_found = find_receipt_header_field_by_live_semantic(
        jab,
        label,
        scope_hwnd=scope_hwnd,
        timeout=0.8,
        include_scoped=True,
    )
    live_snapshot = context_snapshot(jab, live_found)
    semantic_label = find_receipt_header_field_by_semantic_label(
        jab,
        label,
        scope_hwnd=scope_hwnd,
        timeout=0.8,
    )
    semantic_snapshot = context_snapshot(jab, semantic_label)
    live_path = live_found.get("path") if live_found.get("ok") else None
    path_matches_live = bool(
        path_found.get("ok") and live_path == path_found.get("path")
    )
    return {
        "label": label,
        "expected_text_path": expected_text_path,
        "expected_label_path": expected_label_path,
        "path_lookup": strip_context_handles(path_found),
        "path_snapshot": path_snapshot,
        "live_lookup": strip_context_handles(live_found),
        "live_snapshot": live_snapshot,
        "semantic_label_lookup": strip_context_handles(semantic_label),
        "semantic_label_snapshot": semantic_snapshot,
        "path_matches_live": path_matches_live,
        "risk": field_risk(label, path_found, live_found, path_snapshot, live_snapshot),
    }


def field_risk(label, path_found, live_found, path_snapshot, live_snapshot):
    reasons = []
    if not path_found.get("ok"):
        reasons.append("动态 path 未找到字段")
    if not live_found.get("ok"):
        reasons.append("语义 label 未找到字段")
    if path_found.get("ok") and live_found.get("ok"):
        if path_found.get("path") != live_found.get("path"):
            reasons.append("动态 path 与语义 label 定位到不同字段，疑似错位")
    if label == "客户":
        texts = " ".join(
            str((path_snapshot or {}).get(key) or "")
            for key in ("text", "name", "description")
        )
        if "上海移为" in texts or "移为通信" in texts:
            reasons.append("客户字段 path 当前内容像财务组织主体，疑似读到财务组织")
    return {"ok": not reasons, "reasons": reasons}


def visible_nc_windows(jab):
    rows = []
    fg = foreground_info()
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if class_name not in {
            "YonyouUWnd",
            "SunAwtFrame",
            "SunAwtCanvas",
            "SunAwtWindow",
        }:
            continue
        java = False
        try:
            java = (
                bool(jab.dll.isJavaWindow(hwnd))
                if class_name.startswith("SunAwt")
                else False
            )
        except Exception:
            java = False
        rows.append(
            {
                "hwnd": int(hwnd),
                "root_hwnd": root_hwnd(hwnd),
                "pid": int(pid or 0),
                "class_name": class_name,
                "title": title,
                "visible": bool(visible),
                "is_java": java,
                "is_foreground_root": bool(
                    fg.get("root") and root_hwnd(hwnd) == fg.get("root")
                ),
            }
        )
    return rows


def diagnose(config):
    jab = JABOperator(config)
    report = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "readonly": True,
        "windows": {},
        "environment": {
            "windows": platform.platform(),
            "process_bits": 64 if sys.maxsize > 2**32 else 32,
            "is_admin": is_admin(),
        },
    }
    try:
        report["access_bridge_prepare"] = prepare_java_access_bridge()
        jab.ensure_started()
        report["jab_loaded_path"] = getattr(jab, "loaded_path", None)
        report["jab_health"] = check_jab_ready(jab)
        query_cfg = config.get("receipt_query") or {}
        report["receipt_parent_guard"] = guard_receipt_parent_page(
            jab, config, query_cfg
        )
        report["foreground"] = foreground_info()
        report["windows"]["visible_nc"] = visible_nc_windows(jab)
        scope_reports = []
        for window in report["windows"]["visible_nc"]:
            if window.get("class_name") != "SunAwtCanvas" or not window.get("visible"):
                continue
            if not window.get("is_java"):
                continue
            anchor = resolve_receipt_header_anchor_in_canvas(
                jab,
                window["hwnd"],
                timeout=0.8,
            )
            scope = {
                "hwnd": window["hwnd"],
                "root_hwnd": window["root_hwnd"],
                "is_foreground_root": window["is_foreground_root"],
                "anchor": anchor,
                "fields": [],
            }
            if anchor.get("ok"):
                dynamic_index = anchor.get("dynamic_index")
                scope["dynamic_index"] = dynamic_index
                scope["dynamic_prefix"] = receipt_header_dynamic_prefix(dynamic_index)
                for label in HEADER_REQUIRED_LABELS:
                    scope["fields"].append(
                        diagnose_field(jab, label, dynamic_index, window["hwnd"])
                    )
            scope_reports.append(scope)
        report["scopes"] = scope_reports
        report["conclusion"] = conclude(report)
        return report
    finally:
        jab.close()


def conclude(report):
    issues = []
    health = report.get("jab_health") or {}
    if not health.get("ok"):
        issues.append(f"JAB 链路异常: {health.get('reason')}")
    guard = report.get("receipt_parent_guard") or {}
    if not guard.get("ok"):
        issues.append(f"当前页面不是收款单录入或无法确认: {guard.get('reason')}")
    anchored_scopes = [
        scope
        for scope in report.get("scopes") or []
        if (scope.get("anchor") or {}).get("ok")
    ]
    foreground_scopes = [
        scope
        for scope in report.get("scopes") or []
        if scope.get("is_foreground_root") and (scope.get("anchor") or {}).get("ok")
    ]
    if not anchored_scopes:
        issues.append("未找到收款单自制页表头财务组织锚点")
        diagnostic_scopes = []
    elif len(anchored_scopes) == 1:
        diagnostic_scopes = anchored_scopes
    elif foreground_scopes:
        diagnostic_scopes = foreground_scopes
    else:
        diagnostic_scopes = anchored_scopes
        issues.append(
            f"找到 {len(anchored_scopes)} 个表头 scope，且无法按前台窗口唯一确认；请只保留一个收款单自制页"
        )
    for scope in diagnostic_scopes:
        for field in scope.get("fields") or []:
            risk = field.get("risk") or {}
            if not risk.get("ok"):
                issues.extend(
                    f"{field.get('label')}: {reason}"
                    for reason in risk.get("reasons") or []
                )
    return {
        "ok": not issues,
        "issues": issues,
        "diagnostic_scope_count": len(diagnostic_scopes),
        "anchored_scope_count": len(anchored_scopes),
        "foreground_scope_count": len(foreground_scopes),
        "summary": "未发现表头定位明显异常" if not issues else "发现表头定位/环境异常",
    }


def is_admin():
    if sys.platform != "win32":
        return None
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return None


def write_outputs(report):
    logs = OUTPUT_BASE / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_path = logs / f"receipt_header_diagnose_{stamp}.json"
    txt_path = logs / f"receipt_header_diagnose_{stamp}.txt"
    json_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    lines = [
        "NC 收款单表头诊断报告",
        f"生成时间: {report.get('generated_at')}",
        "",
        "结论",
        f"- {report.get('conclusion', {}).get('summary')}",
    ]
    for issue in (report.get("conclusion") or {}).get("issues") or []:
        lines.append(f"- {issue}")
    lines.extend(
        [
            "",
            "说明",
            "- 本诊断只读，不写入 NC 字段，不保存单据。",
            "- 请在 NC 已打开“新增-自制”后的收款单表头页面运行。",
            "",
            f"JSON 明细: {json_path}",
        ]
    )
    txt_path.write_text("\n".join(lines), encoding="utf-8")
    return json_path, txt_path


def main(argv=None):
    stdout_reconfigure = getattr(sys.stdout, "reconfigure", None)
    if callable(stdout_reconfigure):
        stdout_reconfigure(encoding="utf-8", errors="replace")
    args = parse_args(argv)
    config = load_config(str(resolve_config_path(args.config)))
    report = diagnose(config)
    json_path, txt_path = write_outputs(report)
    if args.json:
        print(
            json.dumps(
                {**report, "json_path": str(json_path), "txt_path": str(txt_path)},
                ensure_ascii=False,
                indent=2,
            )
        )
    else:
        print(f"诊断完成: {txt_path}")
        print(f"JSON 明细: {json_path}")
        print(f"结论: {(report.get('conclusion') or {}).get('summary')}")
    return 0 if (report.get("conclusion") or {}).get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
