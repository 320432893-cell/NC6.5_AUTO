# 职责: 解析收款单查询条件窗口的动态前缀，并用稳定后缀定位条件输入框
# 不做什么: 不打开查询窗口，不点击确定，不读取查询结果，不使用旧 near-label 写入兜底
# 允许依赖层: JAB operator-like 对象、tools.receipt_query_pagination_paths 路径工具
# 谁不应该 import: core、Excel/Sheet 写入、收款单录入明细模块不应 import

import time

from tools.receipt_query_pagination_paths import join_context_path, split_context_path


QUERY_CONDITION_PREFIX_BASE = "0.0.1.0.1.0.0.1.0.0.0.0.0.1.0"
QUERY_CONDITION_DYNAMIC_MAX_INDEX = 8
QUERY_FIELD_SUFFIXES = {
    "finance_org": "1.2.0.0.0.0",
    "document_date_from": "3.2.0.0.0.0",
    "document_date_to": "3.2.1.0.0.0",
}
QUERY_FIELD_LABELS = {
    "finance_org": "收款财务组织",
    "document_date_from": "单据日期",
    "document_date_to": "单据日期",
}
QUERY_REQUIRED_FIELDS = ("finance_org", "document_date_from", "document_date_to")


def find_query_condition_scope(jab, jab_cfg):
    cached = getattr(jab, "_receipt_query_condition_scope_cache", None)
    if cached:
        checked = validate_query_condition_scope(
            jab,
            jab_cfg,
            cached.get("prefix"),
            required_fields=("finance_org",),
            source="cached_path",
        )
        if checked.get("ok"):
            checked["cached"] = True
            return checked
    configured = infer_scope_from_configured_paths(jab_cfg)
    if configured:
        checked = validate_query_condition_scope(
            jab,
            jab_cfg,
            configured,
            required_fields=("finance_org",),
            source="configured_path",
        )
        if checked.get("ok"):
            setattr(jab, "_receipt_query_condition_scope_cache", checked)
            return checked
    scanned = scan_query_condition_scope(jab, jab_cfg)
    if scanned.get("ok"):
        setattr(jab, "_receipt_query_condition_scope_cache", scanned)
        return scanned
    semantic = find_query_condition_scope_by_semantic(jab, jab_cfg)
    if semantic.get("ok"):
        semantic["path_attempt"] = scanned
        setattr(jab, "_receipt_query_condition_scope_cache", semantic)
        return semantic
    return {
        "ok": False,
        "reason": "未能定位查询条件动态前缀",
        "configured_attempt": configured,
        "scan_attempt": scanned,
        "semantic_attempt": semantic,
    }


def infer_scope_from_configured_paths(jab_cfg):
    fields = jab_cfg.get("fields") or {}
    path = ((fields.get("finance_org") or {}).get("text_path")) or ""
    suffix = QUERY_FIELD_SUFFIXES["finance_org"]
    path_parts = split_context_path(path)
    suffix_parts = split_context_path(suffix)
    if len(path_parts) <= len(suffix_parts):
        return None
    if path_parts[-len(suffix_parts) :] != suffix_parts:
        return None
    return ".".join(str(part) for part in path_parts[: -len(suffix_parts)])


def scan_query_condition_scope(
    jab,
    jab_cfg,
    max_index=QUERY_CONDITION_DYNAMIC_MAX_INDEX,
    required_fields=QUERY_REQUIRED_FIELDS,
):
    attempts = []
    for dynamic_index in range(max_index + 1):
        prefix = f"{QUERY_CONDITION_PREFIX_BASE}.{dynamic_index}"
        checked = validate_query_condition_scope(
            jab,
            jab_cfg,
            prefix,
            required_fields=required_fields,
        )
        attempts.append(
            {
                "prefix": prefix,
                "ok": checked.get("ok"),
                "ok_fields": checked.get("ok_fields"),
                "fields": checked.get("fields"),
            }
        )
        if checked.get("ok"):
            checked["attempts"] = attempts
            return checked
    return {
        "ok": False,
        "reason": "未能定位查询条件动态前缀",
        "attempts": attempts,
    }


def validate_query_condition_scope(
    jab,
    jab_cfg,
    prefix,
    required_fields=QUERY_REQUIRED_FIELDS,
    source="dynamic_path",
):
    fields = {}
    ok_fields = []
    for field_name in required_fields:
        path = build_query_field_path(prefix, field_name)
        found = jab.wait_context_by_path(
            path,
            title=jab_cfg["dialog_title"],
            class_name=jab_cfg["dialog_class"],
            role="text",
            timeout=float(jab_cfg.get("path_ready_timeout", 0.05)),
            require_showing=True,
            require_valid_bounds=False,
        )
        ok = bool(found)
        fields[field_name] = {"ok": ok, "path": path}
        if ok:
            ok_fields.append(field_name)
    return {
        "ok": len(ok_fields) == len(tuple(required_fields)),
        "prefix": prefix,
        "source": source,
        "ok_fields": ok_fields,
        "fields": fields,
    }


def find_query_condition_scope_by_semantic(jab, jab_cfg):
    fields = {}
    ok_fields = []
    for field_name in QUERY_REQUIRED_FIELDS:
        found = find_query_field_by_semantic_label(jab, jab_cfg, field_name)
        fields[field_name] = {
            "ok": bool(found.get("ok")),
            "path": found.get("path"),
            "label_path": found.get("label_path"),
            "reason": found.get("reason"),
        }
        if found.get("ok"):
            ok_fields.append(field_name)
            jab.release_contexts(found["vm_id"], found["owned_contexts"])
    prefix = infer_common_query_prefix_from_fields(fields)
    return {
        "ok": len(ok_fields) == len(QUERY_REQUIRED_FIELDS) and bool(prefix),
        "source": "semantic_path_inference",
        "prefix": prefix,
        "ok_fields": ok_fields,
        "fields": fields,
        "reason": None
        if len(ok_fields) == len(QUERY_REQUIRED_FIELDS) and prefix
        else "查询条件语义路径推断未命中全部必填字段",
    }


def infer_common_query_prefix_from_fields(fields):
    prefixes = []
    for field_name in QUERY_REQUIRED_FIELDS:
        path = ((fields.get(field_name) or {}).get("path")) or ""
        suffix = QUERY_FIELD_SUFFIXES[field_name]
        path_parts = split_context_path(path)
        suffix_parts = split_context_path(suffix)
        if len(path_parts) <= len(suffix_parts):
            return None
        if path_parts[-len(suffix_parts) :] != suffix_parts:
            return None
        prefixes.append(
            ".".join(str(part) for part in path_parts[: -len(suffix_parts)])
        )
    return prefixes[0] if prefixes and len(set(prefixes)) == 1 else None


def find_query_field_by_semantic_label(jab, jab_cfg, field_name):
    label = QUERY_FIELD_LABELS.get(field_name)
    if not label:
        return {"ok": False, "reason": f"查询条件字段无语义标签: {field_name}"}
    label_found = find_context_with_path_in_query_window(
        jab,
        jab_cfg,
        label,
        roles=("label",),
        timeout=float(jab_cfg.get("semantic_timeout", 1.5)),
        require_showing=True,
    )
    label_context, vm_id, owned_contexts, owned_indexes = label_found
    if not label_context:
        return {"ok": False, "field": field_name, "reason": "semantic label not found"}
    label_path = "0" + "".join(f".{index}" for index in owned_indexes)
    jab.release_contexts(vm_id, owned_contexts)
    text_path = infer_query_text_path_from_label_path(field_name, label_path)
    if not text_path:
        return {
            "ok": False,
            "field": field_name,
            "label_path": label_path,
            "reason": "semantic label path cannot infer text path",
        }
    found = jab.find_context_by_path_once(
        text_path,
        title=jab_cfg["dialog_title"],
        class_name=jab_cfg["dialog_class"],
        role="text",
        require_showing=True,
        require_valid_bounds=False,
    )
    context, vm_id, owned_contexts, window_info = found
    if not context:
        return {
            "ok": False,
            "field": field_name,
            "label_path": label_path,
            "path": text_path,
            "reason": "semantic inferred text path not found",
        }
    return {
        "ok": True,
        "field": field_name,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "path": text_path,
        "label_path": label_path,
        "window": window_info,
    }


def infer_query_text_path_from_label_path(field_name, label_path):
    parts = split_context_path(label_path)
    if not parts:
        return None
    if field_name == "finance_org":
        return join_context_path(
            ".".join(str(part) for part in parts[:-2]), "2.0.0.0.0"
        )
    if field_name == "document_date_from":
        return join_context_path(".".join(str(part) for part in parts[:-2]), "2.0.0.0")
    if field_name == "document_date_to":
        return join_context_path(".".join(str(part) for part in parts[:-2]), "2.1.0.0")
    return None


def find_context_with_path_in_query_window(
    jab,
    jab_cfg,
    name,
    roles=(),
    timeout=None,
    require_showing=False,
):
    if not hasattr(jab, "find_context_once_with_path"):
        return None, None, [], []
    deadline = time.time() + (timeout or 1.5)
    normalized_roles = {role.lower() for role in roles}
    while time.time() < deadline:
        result = jab.find_context_once_with_path(
            name,
            normalized_roles,
            require_showing=require_showing,
            window_title=jab_cfg["dialog_title"],
            window_class=jab_cfg["dialog_class"],
            visible_only=True,
        )
        if result[0]:
            return result
        time.sleep(0.1)
    return None, None, [], []


def build_query_field_path(prefix, field_name):
    suffix = QUERY_FIELD_SUFFIXES.get(field_name)
    if not suffix:
        return None
    return join_context_path(prefix, suffix)


def set_query_dynamic_text(jab, jab_cfg, scope, field_name, value):
    path = build_query_field_path(scope.get("prefix"), field_name)
    if not path:
        return {
            "ok": False,
            "reason": f"查询条件字段未配置动态后缀: {field_name}",
        }
    ok = jab.set_text_by_path(
        path,
        value,
        title=jab_cfg["dialog_title"],
        class_name=jab_cfg["dialog_class"],
        role="text",
        wait=float(jab_cfg.get("text_set_wait", 0.0)),
        timeout=2,
        require_showing=True,
    )
    return {
        "ok": bool(ok),
        "field": field_name,
        "path": path,
        "prefix": scope.get("prefix"),
    }
