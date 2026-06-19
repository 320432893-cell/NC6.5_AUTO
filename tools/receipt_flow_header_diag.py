# 职责：读回/诊断已写表头字段，确认客户名称，并归纳表头写入失败原因
# 不做什么：不写入字段，不触发保存，不做报告落盘，不做行编排
# 允许依赖层：tools.receipt_self_made_flow 的表头字段定位/客户名校验
# 谁不应该 import：core 层模块不应 import；本模块不应反向 import row_runner

import sys

from tools.receipt_self_made_flow import is_valid_customer_name_candidate


class _FlowNamespace:
    # 按调用时从已加载的入口模块取属性：让测试对
    # tools.receipt_full_flow_entry.find_receipt_header_field_by_dynamic_path 的
    # monkeypatch 与拆分前一致地生效，且不在加载期 import 入口模块以避免成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_full_flow_entry"], name)


_flow = _FlowNamespace()


def read_customer_name_after_header(jab, header_steps, dynamic_index, scope_hwnd):
    step = next(
        (item for item in header_steps or [] if item.get("label") == "客户"),
        None,
    )
    attempts = []
    if step and step.get("path"):
        found = _flow.find_receipt_header_field_by_dynamic_path(
            jab,
            "客户",
            step.get("dynamic_index") or dynamic_index,
            scope_hwnd=scope_hwnd,
            require_showing=False,
            require_valid_bounds=False,
            path_template=(step.get("path_attempt") or {}).get("path_template"),
        )
        attempts.append(
            read_customer_name_from_found_field(jab, found, source="path-readback")
        )
    if step:
        attempts.append(
            {
                "ok": True,
                "source": "header-step-snapshot",
                "value": extract_header_accepted_text([step], "客户"),
                "snapshot": step.get("post_write_snapshot")
                or step.get("backend_state")
                or {},
            }
        )
    for attempt in attempts:
        value = str(attempt.get("value") or "").strip()
        if is_valid_customer_name_candidate(value):
            return {
                "ok": True,
                "value": value,
                "source": attempt.get("source"),
                "attempts": attempts,
            }
    return {
        "ok": False,
        "value": "",
        "attempts": attempts,
        "reason": "客户名称未确认：客户字段 description 未读到有效 NC 客户名称",
    }


def read_customer_name_from_found_field(jab, found, source):
    if not found.get("ok"):
        return {
            "ok": False,
            "source": source,
            "reason": found.get("reason"),
            "path": found.get("path"),
        }
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    try:
        info = jab.get_context_info(vm_id, context)
        text = jab.get_text_context_value(vm_id, context)
        description = info.description.strip() if info else ""
        name = info.name.strip() if info else ""
        value = first_valid_text(description, text, name)
        return {
            "ok": bool(value),
            "source": source,
            "value": value,
            "path": found.get("path"),
            "label_path": found.get("label_path"),
            "text": text,
            "name": name,
            "description": description,
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def first_valid_text(*values):
    for value in values:
        text = str(value or "").strip()
        if is_valid_customer_name_candidate(text):
            return text
    return ""


def extract_header_accepted_text(header_steps, label):
    for step in header_steps or []:
        if step.get("label") != label:
            continue
        text = str(step.get("accepted_text") or "").strip()
        if is_valid_customer_name_candidate(text):
            return text
        backend = step.get("post_write_snapshot") or step.get("backend_state") or {}
        for key in ("description", "text", "name"):
            value = str(backend.get(key) or "").strip()
            if (
                value
                and value != str(step.get("value") or "").strip()
                and is_valid_customer_name_candidate(value)
            ):
                return value
    return ""


def diagnose_written_header_fields(
    jab,
    labels,
    dynamic_index,
    dynamic_prefix,
    scope_hwnd,
):
    results = []
    for label in labels or []:
        found = _flow.find_receipt_header_field_by_dynamic_path(
            jab,
            label,
            dynamic_index,
            scope_hwnd=scope_hwnd,
            require_showing=False,
            require_valid_bounds=False,
        )
        if not found.get("ok"):
            results.append(
                {
                    "label": label,
                    "ok": False,
                    "present": False,
                    "reason": found.get("reason") or "field path not found",
                    "dynamic_prefix": dynamic_prefix,
                    "path": found.get("path"),
                }
            )
            continue
        context = found["context"]
        vm_id = found["vm_id"]
        owned_contexts = found["owned_contexts"]
        try:
            info = jab.get_context_info(vm_id, context)
            text = jab.get_text_context_value(vm_id, context)
            description = info.description.strip() if info else ""
            name = info.name.strip() if info else ""
            present = bool(str(text or "").strip() or description or name)
            results.append(
                {
                    "label": label,
                    "ok": True,
                    "present": present,
                    "text": text,
                    "description": description,
                    "name": name,
                    "path": found.get("path"),
                    "dynamic_prefix": found.get("dynamic_prefix") or dynamic_prefix,
                }
            )
        finally:
            jab.release_contexts(vm_id, owned_contexts)
    return results


def summarize_header_failure(header_steps):
    for step in header_steps or []:
        if step.get("ok"):
            continue
        label = step.get("label")
        reason = (
            step.get("reason")
            or step.get("stage")
            or ((step.get("scope") or {}).get("reason"))
            or ((step.get("path_attempt") or {}).get("reason"))
            or "表头字段写入失败"
        )
        return f"表头字段写入失败: {label or step.get('step') or '未知字段'} - {reason}"
    return "表头字段写入失败"
