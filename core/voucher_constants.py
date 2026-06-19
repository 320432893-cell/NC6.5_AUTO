# 职责：制单/凭证保存相关的业务常量与谓词单一来源(SSOT)——制单表列数判定、保存策略/触发/激活枚举
# 不做什么：不读 JAB、不持有 processor/JAB 状态、不做表格读取或保存编排、不解析 CLI
# 允许依赖层：纯常量模块，无内部依赖；可被 core 与 tools 任意 import
# 谁不应该 import：本模块不 import 任何项目内模块，避免循环依赖

# 制单窗口表格固定为 N 行 x 13 列；行数必须大于 0 才算有效制单表。
# 口径来自 README「制单窗口：表格通常是 N rows x 13 cols」。
VOUCHER_TABLE_COL_COUNT = 13


def is_voucher_table(table) -> bool:
    """判断一张已读取的表格是否为制单表(N 行 x 13 列且有数据行)。

    table 是 read_window_table_cells/read_window_table_counts 返回的表字典，
    可能用普通下标(table["col_count"])或 .get 访问；这里统一用 .get 容忍缺键。
    """
    return (
        table.get("row_count", 0) > 0
        and table.get("col_count") == VOUCHER_TABLE_COL_COUNT
    )


# 制单保存策略/触发方式/Ctrl+S 激活策略的合法枚举值。
# 被 core.nc_voucher_save(运行时分支)、tools.jab_batch(argparse choices)、
# tools.validate_config(配置校验)三处共用，按分层硬约束放 core。
SAVE_STRATEGIES = ("single", "bottom_up", "safe_batch_by_pending_row")
SAVE_TRIGGERS = ("jab_button", "hotkey")
HOTKEY_ACTIVATE_POLICIES = ("always", "first", "foreground_guard")
