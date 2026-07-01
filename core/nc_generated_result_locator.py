import re
import time

from core.utils import check_abort


MODULE_PREFIX_BASE = "0.0.1.0.0.0.0"
RESULT_TABLE_RELATIVE_PATH = "0.0.0.1.1.0.0.0.1.1.0.2.1.0.0.0.0.0"


def locate_generated_result_table(
    jab,
    batch_cfg,
    voucher_col,
    generated_voucher_max,
):
    jab.ensure_started()
    locator_cfg = batch_cfg.get("generated_result_locator") or {}
    module_prefix_base = str(
        locator_cfg.get("module_prefix_base")
        or batch_cfg.get("generated_result_module_prefix_base")
        or MODULE_PREFIX_BASE
    ).strip(".")
    relative_table_path = str(
        locator_cfg.get("relative_table_path")
        or batch_cfg.get("generated_result_table_relative_path")
        or RESULT_TABLE_RELATIVE_PATH
    ).strip(".")
    max_index = int(locator_cfg.get("max_index", batch_cfg.get("module_dynamic_max_index", 8)))
    window_class = locator_cfg.get("window_class", "SunAwtCanvas")
    sample_rows = int(locator_cfg.get("sample_rows", 5))

    attempts = []
    for dynamic_index in range(max_index + 1):
        check_abort()
        scope_path = f"{module_prefix_base}.{dynamic_index}"
        table_path = f"{scope_path}.{relative_table_path}"
        context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
            table_path,
            class_name=window_class,
            role="table",
            require_showing=False,
            require_valid_bounds=False,
        )
        if not context:
            attempts.append(
                {
                    "index": dynamic_index,
                    "scope_path": scope_path,
                    "table_path": table_path,
                    "ok": False,
                    "reason": "path_not_found",
                }
            )
            continue

        try:
            table_info = jab.get_table_info(vm_id, context)
            voucher_values = read_voucher_values(
                jab,
                vm_id,
                context,
                table_info,
                window_info,
                voucher_col,
                generated_voucher_max,
                sample_rows,
            )
            rows = int(table_info.rowCount) if table_info else None
            cols = int(table_info.columnCount) if table_info else None
            ok = bool(table_info and voucher_values)
            attempt = {
                "index": dynamic_index,
                "scope_path": scope_path,
                "table_path": table_path,
                "ok": ok,
                "rows": rows,
                "cols": cols,
                "voucher_values": voucher_values,
                "reason": "matched" if ok else "voucher_col_not_confirmed",
            }
            attempts.append(attempt)
            if ok:
                return {
                    "ok": True,
                    "reason": "module_index_path_voucher_col",
                    "dynamic_index": dynamic_index,
                    "scope_path": scope_path,
                    "table_path": table_path,
                    "relative_table_path": relative_table_path,
                    "window_class": window_class,
                    "row_count": rows,
                    "col_count": cols,
                    "voucher_values": voucher_values,
                    "attempts": attempts,
                }
        finally:
            jab.release_contexts(vm_id, owned_contexts)

    return {
        "ok": False,
        "reason": "generated_result_table_not_found",
        "attempts": attempts,
    }


def wait_generated_result_table(
    jab,
    batch_cfg,
    voucher_col,
    generated_voucher_max,
    timeout,
    interval=0.2,
):
    deadline = time.time() + float(timeout)
    last = None
    while True:
        last = locate_generated_result_table(
            jab,
            batch_cfg,
            voucher_col,
            generated_voucher_max,
        )
        if last.get("ok"):
            return last
        if time.time() >= deadline:
            return last
        time.sleep(float(interval))


def read_voucher_values(
    jab,
    vm_id,
    context,
    table_info,
    window_info,
    voucher_col,
    generated_voucher_max,
    max_rows,
):
    if not table_info:
        return []
    if voucher_col < 0 or voucher_col >= int(table_info.columnCount):
        return []
    table = jab.read_table_selected_columns_from_context(
        0,
        context,
        vm_id,
        table_info,
        window_info,
        [voucher_col],
        max_rows=max_rows,
    )
    values = []
    for row in table.get("rows") or []:
        cells = row.get("cells") or []
        if voucher_col >= len(cells):
            continue
        text = str(cells[voucher_col] or "").strip()
        if is_strict_voucher_text(text, generated_voucher_max):
            values.append(text)
    return values


def is_strict_voucher_text(text, generated_voucher_max):
    value = str(text or "").strip()
    if not re.fullmatch(r"\d+", value):
        return False
    number = int(value)
    return 1 <= number <= int(generated_voucher_max)
