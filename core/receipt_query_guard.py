# 职责：校验收款单查询前置页面和结果表是否符合收款单语义。
# 不做什么：不读取分页，不填写查询条件，不生成匹配报告。
# 允许依赖层：JAB 操作对象接口和查询结果表纯数据结构。
# 谁不应该 import：JAB 底层生命周期模块、一次性探针脚本。


class ReceiptPageGuardError(RuntimeError):
    pass


def first_non_empty_cell_at(cells, column):
    if column >= len(cells):
        return ""
    return str(cells[column] or "").strip()


def guard_receipt_parent_page(jab, config, query_cfg):
    guard_cfg = query_cfg.get("page_guard") or {}
    if not bool(guard_cfg.get("enabled", True)):
        return {"enabled": False, "ok": True}

    jab.ensure_started()
    state_label = (config.get("receipt_entry") or {}).get("state_label", "")
    if not state_label:
        raise ReceiptPageGuardError("receipt_entry.state_label is required")

    context, vm_id, owned_contexts = jab.find_context(
        state_label,
        roles=guard_cfg.get("state_label_roles", ()),
        timeout=float(guard_cfg.get("state_label_timeout", 1.5)),
        require_showing=bool(guard_cfg.get("state_label_require_showing", False)),
        window_title=guard_cfg.get("window_title"),
        window_class=guard_cfg.get("window_class"),
        visible_only=bool(guard_cfg.get("visible_only", True)),
    )
    if context:
        jab.release_contexts(vm_id, owned_contexts)
        return {"enabled": True, "ok": True, "state_label": state_label}

    raise ReceiptPageGuardError(
        f"当前 NC 页面未检测到目标页标识: {state_label!r}，拒绝执行收款查询/写回"
    )


def guard_receipt_result_tables(tables, query_cfg):
    guard_cfg = query_cfg.get("result_guard") or {}
    if not bool(guard_cfg.get("enabled", True)):
        return {"enabled": False, "ok": True}

    indexes = query_cfg.get("result_column_indexes") or {}
    document_type_col = int(guard_cfg.get("document_type_column", 2))
    document_type = str(guard_cfg.get("document_type", "收款单"))
    name_column = int(
        (query_cfg.get("result_column_indexes") or {}).get("payer_name", 2)
    )
    guard_name_column = bool(
        guard_cfg.get("name_column_must_not_equal_document_type", True)
    )
    blocked_keywords = tuple(guard_cfg.get("blocked_keywords", ("应收款", "应付款")))
    max_samples = int(guard_cfg.get("max_samples", 20))
    samples = []
    name_samples = []
    blocked = []
    expected = 0

    for table in tables:
        if table.get("col_count", 0) < max(indexes.values(), default=0) + 1:
            continue
        for row in table.get("rows") or []:
            cells = row.get("cells") or []
            value = first_non_empty_cell_at(cells, document_type_col)
            if not value:
                continue
            name_value = first_non_empty_cell_at(cells, name_column)
            if len(samples) < max_samples:
                samples.append(value)
            if len(name_samples) < max_samples and name_value:
                name_samples.append(name_value)
            if document_type in value:
                expected += 1
            row_text = "\t".join(str(cell or "") for cell in cells)
            row_blocked = [
                keyword for keyword in blocked_keywords if keyword in row_text
            ]
            if row_blocked:
                blocked.extend(row_blocked)

    if blocked:
        raise ReceiptPageGuardError(
            "收款查询结果表疑似来自错误页面，检测到禁用单据类型: "
            f"{sorted(set(blocked))[:10]}"
        )
    if samples and expected == 0:
        raise ReceiptPageGuardError(
            "收款查询结果表未检测到目标单据类型: "
            f"expected={document_type!r} samples={samples[:10]}"
        )
    if (
        guard_name_column
        and name_samples
        and all(document_type in value for value in name_samples)
    ):
        raise ReceiptPageGuardError(
            "收款查询结果表匹配名称列疑似配置到单据类型列: "
            f"name_column={name_column} samples={name_samples[:10]}"
        )
    return {
        "enabled": True,
        "ok": True,
        "document_type": document_type,
        "samples": samples,
        "name_samples": name_samples,
    }
