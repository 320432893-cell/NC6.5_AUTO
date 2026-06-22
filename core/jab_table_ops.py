# 职责：JAB 表格操作兼容门面——实现已拆到 jab_table_{find,read,cache,select},此处统一 re-export
# 不做什么：不再承载实现;仅保持 jab_table_mixin 等既有 import 入口稳定
# 允许依赖层：core.jab_table_{find,read,cache,select}
# 谁不应该 import：Excel/Sheet 读写、收款匹配、配置解析模块不应 import

from core.jab_table_find import (  # noqa: F401
    find_main_table,
    find_tables_in_tree,
    find_tables_once,
    window_descendant_hwnds,
)
from core.jab_table_read import (  # noqa: F401
    read_all_table_cells,
    read_all_table_selected_columns,
    read_table_snapshot,
    read_table_summaries,
    resolve_amount_col,
    resolve_partner_col,
)
from core.jab_table_cache import (  # noqa: F401
    clear_table_cache,
    get_window_table_cache,
    read_window_table_cells,
    read_window_table_counts,
)
from core.jab_table_select import (  # noqa: F401
    get_selected_child_indexes,
    has_selection_api,
    resolve_selection_col,
    select_table_rows,
    select_table_rows_from_context,
    select_visible_table_rows,
    select_visible_table_rows_once,
)
