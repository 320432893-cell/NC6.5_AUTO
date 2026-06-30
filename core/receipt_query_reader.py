# 职责：提供收款单查询读取相关旧导入路径的兼容门面。
# 不做什么：不承载具体读取/分页/匹配实现，不解析 CLI 参数，不写 Excel。
# 允许依赖层：core.receipt_query_page_reader、core.receipt_query_match_reader、core.receipt_query_pagination、core.receipt_query_pagination_paths、core.receipt_query_result_tables。
# 谁不应该 import：新代码不应继续新增对本兼容门面的依赖；应直接 import 具体职责模块。

from core.receipt_query_match_reader import (
    dedupe_page_tables as dedupe_page_tables,
    evaluate_paging_match_stop as evaluate_paging_match_stop,
    read_receipt_result_pages_incremental as read_receipt_result_pages_incremental,
    read_receipt_result_pages_until_match as read_receipt_result_pages_until_match,
    unresolved_excel_rows as unresolved_excel_rows,
)
from core.receipt_query_page_reader import (
    read_receipt_result_pages as read_receipt_result_pages,
)
from core.receipt_query_pagination import (
    click_next_page as click_next_page,
    parse_int_text as parse_int_text,
    parse_page_label as parse_page_label,
    read_page_label as read_page_label,
    read_page_size_text as read_page_size_text,
    set_receipt_page_size as set_receipt_page_size,
    wait_after_query_confirm as wait_after_query_confirm,
    wait_receipt_result_ready as wait_receipt_result_ready,
    wait_receipt_result_stable as wait_receipt_result_stable,
)
from core.receipt_query_pagination_paths import (
    infer_result_area_prefix_from_page_path as infer_result_area_prefix_from_page_path,
    infer_result_area_prefix_from_table_path as infer_result_area_prefix_from_table_path,
    join_context_path as join_context_path,
    resolve_receipt_pagination_paths as resolve_receipt_pagination_paths,
    resolve_receipt_pagination_paths_dynamic as resolve_receipt_pagination_paths_dynamic,
    split_context_path as split_context_path,
    strip_context_path_suffix as strip_context_path_suffix,
    validate_context_path as validate_context_path,
    validate_receipt_pagination_path_report as validate_receipt_pagination_path_report,
    with_runtime_pagination_paths as with_runtime_pagination_paths,
)
from core.receipt_query_result_tables import (
    enumerate_receipt_result_table_paths as enumerate_receipt_result_table_paths,
    find_table_paths_in_context as find_table_paths_in_context,
    first_non_empty_cell as first_non_empty_cell,
    first_non_empty_cell_at as first_non_empty_cell_at,
    is_receipt_result_table_candidate as is_receipt_result_table_candidate,
    read_receipt_result_table_by_path as read_receipt_result_table_by_path,
    read_receipt_result_tables_runtime as read_receipt_result_tables_runtime,
    read_receipt_tables as read_receipt_tables,
    receipt_result_read_columns as receipt_result_read_columns,
    summarize_receipt_tables as summarize_receipt_tables,
)
