# 职责：说明 tools/archive 下历史探测脚本的归档边界
# 不做什么：不承载正式入口、不定义业务规则、不作为测试人员默认菜单
# 允许依赖层：文档说明、历史 tools 探针
# 谁不应该 import：core、正式 tools 入口、tests

# tools/archive

这里保存已退出默认运行路径的现场探测、窗口诊断和历史人工测试脚本。

归档不等于删除：这些文件仍可作为复盘证据或现场排查参考，但不能被正式流程 import，也不能作为交给测试人员的默认入口。若某个归档脚本重新进入正式路径，必须先移回合适目录，补齐职责声明、测试和 `tools/check.py changed` 闭包。

本轮未归档的关键边界：

- `tools/jab_probe.py`：被 core 和多个正式工具依赖。
- `tools/receipt_new_probe.py`：仍承接收款单新增/自制入口能力。
- `tools/archive/receipt_account_reference_try.py`：历史表头账户参照探针，只能用于复盘旧方案；正式流程不得 import。停止热键/剪贴板工具已抽到 `tools/receipt_keyboard_utils.py`。
- `tools/archive/tmp_receipt_account_run.py`：历史表头账户参照手工入口，只能用于复盘旧方案；正式流程不得 import。
- `tools/receipt_table_cell_probe.py`：仍被历史/正式录入流程间接依赖。
- `tools/tmp_receipt_cell_probe_run.py`：仍保留为现场探测脚本；正式明细模块需要的金额比较、剪贴板和受保护按键 helper 已抽到 `tools/receipt_keyboard_utils.py`，正式流程不得再从该 tmp 脚本 import。
- `tools/tmp_receipt_detail_main_line_run.py`：短期兼容壳，现场人员全部改用 `tools/receipt_detail_entry.py` 后再删除。
