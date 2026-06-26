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
- `tools/archive/tmp_receipt_cell_probe_run.py`：历史明细金额/编辑器入口探测脚本；正式明细模块需要的金额比较、剪贴板和受保护按键 helper 已抽到 `tools/receipt_keyboard_utils.py`，正式流程不得再从该 tmp 脚本 import。
- `tools/tmp_receipt_detail_main_line_run.py`：短期兼容壳，现场人员全部改用 `tools/receipt_detail_entry.py` 后再删除。
- `tools/archive/probe_receipt_header_scopes.py`：历史多窗口表头 scope 只读探针；正式流程已在当前入口里做 scope/cache，不从该脚本 import。
- `tools/archive/tmp_jab_recovery_probe.py`：历史 JAB 启动/注册恢复探针，只做现场诊断；正式流程不得自动执行这些窗口唤醒动作。
- `tools/archive/probe_receipt_detail_bounds_stability.py`：明细表 bounds 稳定性只读探针，只用于解释 JAB 返回负坐标、零高度或 `-1,-1,-1,-1` 的现场现象；不得作为正式列定位依据。
- `tools/archive/probe_receipt_detail_layout.py`：明细表布局和附近标题 bounds 诊断，只读；不得作为正式坐标定位依据。
- `tools/archive/probe_receipt_detail_click_x_scan.py`：历史鼠标单击 X 轴扫描探针，会移动鼠标并点击当前 NC 表格，仅限人工现场诊断；正式流程不得 import 或复用。
- `tools/archive/probe_receipt_counterparty_sync.py`：往来对象上下同步/修复探针，用于确认表头 combo 与明细 row0 往来对象状态；正式逻辑只允许复用其已沉淀出的稳定 API 思路，不能复用鼠标/bounds 分支。
