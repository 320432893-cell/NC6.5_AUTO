# 职责：客户名候选采集与校验——回读 NC 客户字段候选、去重、合法性判断、上下文文本快照
# 不做什么：不写表头字段(委托 writer),不解析 scope/anchor,不算 path 模板
# 允许依赖层：标准库 re;字段定位(writer 的 find_receipt_header_field_by_*,被 monkeypatch)经 _trial 代理
# 谁不应该 import：receipt_header_paths/tree 不应 import 本模块

import re
import sys


class _TrialNamespace:
    # 按调用时从已加载的 tools.receipt_self_made_flow 取属性：
    # 让测试对 trial 上 find_receipt_header_field_by_dynamic_path /
    # find_receipt_header_field_by_semantic_label 的 monkeypatch 与拆分前一致地生效，
    # 且不在加载期 import 入口模块以避免成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_self_made_flow"], name)


_trial = _TrialNamespace()


def is_valid_customer_name_candidate(value):
    text = str(value or "").strip()
    if not text:
        return False
    if re.match(r"^\[L?java(\.|x\.)", text) or re.match(
        r"^\[L[^;]+;@[0-9a-fA-F]+$", text
    ):
        return False
    if re.match(r"^[A-Z]{1,5}\d{3,}$", text):
        return False
    if text in {"客户", "客户编码"}:
        return False
    return bool(re.search(r"[\u4e00-\u9fffA-Za-z]", text))


def first_valid_customer_name(candidates):
    for candidate in candidates:
        values = candidate.get("valid_values") or []
        if values:
            return {
                "value": values[0],
                "source": candidate.get("source"),
                "path": candidate.get("path"),
                "parent_path": candidate.get("parent_path"),
                "field": next(
                    (
                        key
                        for key in ("description", "text", "name")
                        if candidate.get(key) == values[0]
                    ),
                    None,
                ),
            }
    return None


def dedupe_customer_candidates(candidates):
    seen = set()
    unique = []
    for item in candidates:
        key = (
            item.get("source"),
            item.get("path"),
            item.get("parent_path"),
            item.get("text"),
            item.get("name"),
            item.get("description"),
            item.get("reason"),
        )
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
    return unique


def valid_customer_values_from_snapshot(snapshot):
    values = []
    for key in ("description", "text", "name"):
        value = str((snapshot or {}).get(key) or "").strip()
        if is_valid_customer_name_candidate(value) and value not in values:
            values.append(value)
    return values


def context_text_snapshot(info, text):
    return {
        "role": (info.role_en_US.strip() or info.role.strip()) if info else "",
        "states": (info.states_en_US.strip() or info.states.strip()) if info else "",
        "text": str(text or "").strip(),
        "name": info.name.strip() if info else "",
        "description": info.description.strip() if info else "",
    }


def collect_context_text_rows(jab, vm_id, context, path, depth, max_depth, rows):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    text = jab.get_text_context_value(vm_id, context)
    snapshot = context_text_snapshot(info, text)
    if any(snapshot.get(key) for key in ("text", "name", "description")):
        rows.append({"path": path, "depth": depth, **snapshot})
    if depth >= max_depth:
        return
    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if role == "table":
        return
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        try:
            collect_context_text_rows(
                jab,
                vm_id,
                child,
                f"{path}.{index}",
                depth + 1,
                max_depth,
                rows,
            )
        finally:
            jab.release_contexts(vm_id, [child])


def snapshot_header_field_candidate(jab, found, source):
    if not found.get("ok"):
        return [
            {
                "ok": False,
                "source": source,
                "reason": found.get("reason"),
                "path": found.get("path"),
                "label_path": found.get("label_path"),
            }
        ]
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    try:
        info = jab.get_context_info(vm_id, context)
        text = jab.get_text_context_value(vm_id, context)
        snapshot = context_text_snapshot(info, text)
        return [
            {
                "ok": True,
                "source": source,
                "path": found.get("path"),
                "label_path": found.get("label_path"),
                "window": found.get("window"),
                "role": snapshot.get("role"),
                "states": snapshot.get("states"),
                "text": snapshot.get("text"),
                "name": snapshot.get("name"),
                "description": snapshot.get("description"),
                "valid_values": valid_customer_values_from_snapshot(snapshot),
            }
        ]
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def collect_customer_field_candidates_for_scope(
    jab, scope_hwnd, dynamic_index, timeout
):
    candidates = []
    path_found = _trial.find_receipt_header_field_by_dynamic_path(
        jab,
        "客户",
        dynamic_index,
        scope_hwnd=scope_hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    candidates.extend(snapshot_header_field_candidate(jab, path_found, "path"))
    semantic_found = _trial.find_receipt_header_field_by_semantic_label(
        jab,
        "客户",
        scope_hwnd=scope_hwnd,
        timeout=timeout,
    )
    candidates.extend(
        snapshot_header_field_candidate(jab, semantic_found, "semantic-label-path")
    )
    return dedupe_customer_candidates(candidates)
