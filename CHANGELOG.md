# 更新日志

只记录影响维护判断的关键节点。具体实验流水账看 git 历史。

## 2026-06-12 - 收款单明细写入正式模块沉淀

- 文档口径修正：明确当前无正式 GUI/前端，`.bat` 只是测试菜单；`tools/nc_auto_test_menu.bat` 中凭证生成项会保存凭证、收款完整流程保存项会保存收款单；`tools/receipt_full_flow_entry.py --query-after-save` 当前只是 deferred 占位；`core/receipt_sheet.py::rewrite_plan_sheet()` 当前会重写 Sheet2 当前计划结果区，不是历史追加表。
- 明细主行/手续费行能力已从 `tools/tmp_receipt_detail_main_line_run.py` 拆到正式 `tools/receipt_detail_*` 模块：字段映射和读回校验、明细表读取、JAB 选中+前台守卫键盘写入、整行重试、手续费流程、清账户和删多余行分别落到独立文件。
- 新增收款单完整流程测试入口 `tools/receipt_full_flow_entry.py`：消费 `ReceiptPlanRow`，默认跑 `新增 -> 自制 -> 表头 -> 明细主行 -> 手续费分支` 后停在保存前；显式 `--save` 才调用 JAB 保存按钮真实保存。入口默认只取 1 行，支持指定 `--excel-row` 和运行前 `--write-plan-sheet`。
- 新增项目级测试菜单 `tools/nc_auto_test_menu.bat`：覆盖工程检查、凭证计划/生成/回填/切已生成、收款本地预检、写 Sheet2、收款完整流程不保存/保存、明细测试、查询读取和历史写回。菜单记录开始时间、结束时间和退出码，并按只读、写 Excel、真实 NC 不保存、真实保存分组提示风险。
- 当前仍未完成的是“保存后按主体统一后验查询 -> 保存/查询结果结构化落 Sheet2”。可测试明细正式模块的主入口是 `tools/receipt_detail_test_menu.bat`，菜单可选写主行、写手续费、清理多余行和查看帮助；脚本化入口是 `tools/receipt_detail_entry.py`。`tools/tmp_receipt_detail_main_line_run.py` 只保留短期兼容并转发到正式入口。
- 手续费覆盖守卫已补强：第 2 行只有为空或已是手续费行时才允许覆盖；删行和字段写入返回结构会标记 `changed`、`partial_success` 和副作用策略，便于失败后人工判断是否已有部分写入。

## 2026-06-11 - 收款单录入/查询 T0 达标并清理临时脚本

- 当前阶段总体速度和稳定性已达标，T0 真实保存脚本开始清理；后续不要再把 `tools/tmp_receipt_two_case_save_run.py`、`tools/tmp_receipt_three_org_real_save_run.py` 当主入口，必要时从 git 历史恢复现场脚本。仓库仍保留若干 `tools/tmp_*` 探测/诊断脚本，只作复盘参考，不是正式入口。
- 后验查询已改为按主体分组查询：同一主体只查一次，查询条件使用本批保存行的日期最小值到最大值；查询结果在内存匹配本批行，最后批量追加 Sheet2。Sheet2 保持业务字段精简，异常原因只写阶段前缀加原因，例如 `保存：...` / `查询：...`。
- 查询结果区已确认可用“动态 path 推导 + 缓存 + 失效回退”：语义先确认当前是收款单录入页，再从结果区推导动态前缀。主结果表后缀为 `0.0.0`，分页 label 后缀 `1.6`，每页条数后缀 `1.7`，下一页后缀 `1.2`。不要只靠 41 列判断主结果表，读取时还要看单据号、日期、金额等主表行特征，避免混入明细表。
- 查询分页策略已收敛：先确保每页条数为 `500`；如果结果总数小于等于 `500`，只读本页；大于 `500` 时先读本页，未命中再翻页。翻页后至少等待约 `0.3s` 并等分页/结果摘要稳定，避免 NC 异步刷新未完成导致闪退或读错页。
- 写入侧已验证“动态 path 推导 + 缓存”方向：`新增 -> 自制` 后首笔定位表头 scope，后续复用同一动态段和字段 path；保存后等待【新增】改为快速轮询，避免固定长等。财务组织仍是阻塞式第一步，写入前后的表头 path 可能不同，必须先填财务组织再填其它字段。
- 最新三笔真实保存验证覆盖 `A001/CNY/无手续费`、`A006/USD/有手续费`、`A001/USD/无手续费`；保存成功 oracle 仍是 `Ctrl+S` 后重新看到【新增】。本轮保存结果已写 Sheet2，测试保存单据和 Sheet2 测试行不自动清理。

## 2026-06-10 - 授权主体真实保存、Sheet2 简化和查询性能结论

- 授权范围按当前账号实际权限收敛为 `A001/移为` 和 `A006/移为香港`；三案例真实保存 T0 已跑通，覆盖 `CNY`、`USD`、有手续费、无手续费。用户已删除测试保存单据，不能把这些单据当后续查询样本。
- 当前银行账号输入统一使用裸账号，不再加 `RMB/USD/CNY` 后缀。只有 `97460154740002297` 是 `CNY`，其余已配置账号按 `USD` 处理。币种只写表头，输入 `USD`/`CNY` 后回车；主行和手续费行不再写币种。
- 表头区 `结算方式` 必须显式写 `网银`。A001 默认可能已有值，但 A006/移为香港默认曾为空，不能依赖 NC 默认值。明细主行和手续费行的 `结算方式` 继续用输入后 Enter 确认，用于提交并关闭联想浮层。
- Sheet2 `收款单自动化结果` 改为人看的业务表：追加写入，不清空旧数据；缺表头时补齐；已有旧噪音列会按表头名删除。保留字段为原 Sheet1 行号、执行主体、银行、日期、客户、币种、银行来款名、实收、手续费、总金额、收款银行账户、本地预检状态、异常原因。录入/保存/查询的调试耗时不写 Sheet2。
- 后验查询必须在全部保存完成后统一进入查询阶段，不能保存一单查一单。当前 T0 脚本已做到保存阶段和查询阶段分离，但查询实现仍是逐案例调用 `receipt_query_fill.py`；上次约 35s/单，主要耗在 `read_receipt_result_pages` 约 27s。下一步性能优化应按 `(主体, 日期)` 分组批量查一次，再在内存匹配本批保存结果。

## 2026-06-09 - 收款单自制录入无保存两案例跑通和提速

- `tools/tmp_receipt_two_case_save_run.py` 当前用于无保存诊断：默认 `SAVE_ENABLED = False`，只填写到保存前停止，不发送 `Ctrl+S`。本轮用户指定测试银行账号固定为 `FTE1219165931831RMB`，选择最近两条 A001 人民币有效行，手续费测试值覆盖为 `20.00`、`33.00`。
- 最新现场无保存测试已跑通一条有手续费案例：Excel 行 `1791`，日期 `2026-05-22`，客户 `YW03574`，金额 `391000.00`，手续费 `20.00`。表头、主行、手续费行、手续费账户清空、多余第 3 行删除、保存前表头回看均通过；未保存、未暂存。
- 当前最新耗时基线：单案例总耗时 `17.765s`；`新增->自制 1.843s`，表头 `2.964s`，财务组织写完后回看 `0.301s`，主行 `4.512s`，手续费 `5.579s`，保存前表头回看 `0.914s`。
- `receipt_new_probe.run()` 支持复用外部已启动的 JAB、开单前窗口快照和已定位的【新增】按钮。两案例脚本现在同一案例内复用 JAB，不再在起始检测、新增探测、填写阶段反复初始化。
- `新增->自制` 入口态检测改成快速确认：点中【自制】后快速看关键按钮，识别不到时记录 `partial_ok` 并交给后续表头定位兜底，不再把入口态识别失败直接判死。实际 `wait-for-entry-state` 从约 `1.465s` 降到约 `0.336s`。
- 表头写入成功后缓存字段 path，财务组织、单据日期、客户回看优先按缓存 path 读，失败再回退语义扫描。保存前表头回看从约 `1.700s` 降到约 `0.914s`。
- 明细写入仍保持逐字段 JAB selection + 受保护键盘写入 + 整行最终读回校验。不要恢复跨字段方向键批量写；该方案现场曾打开额外查询/参照窗口，风险高于收益。
- 结算方式/网银字段使用输入后 Enter 确认；现场观察确认推荐浮层会立即消失。Enter 可能带来空白第 3 行，但当前多余行删除已稳定，手续费分支会在写完后用 `Ctrl+D` 删除第 3 行。
- 当时仍不能对外宣称真实保存闭环完成：本轮用户明确不要保存，只做填写诊断和速度调优。正式保存前仍需用户明确授权，并以 `Ctrl+S` 后重新看到【新增】作为保存成功 oracle。

## 2026-06-08 - 收款单本地预检和 Sheet2 机器结果表

- 收款单新主线改为：从 `receipt_entry.excel.start_row` 开始读 Sheet1，先做本地配置识别和异常识别；录入前不查 NC，不再用“最近 2 个月”或 `是否NC已做过` 状态列决定本批候选。业务前提是交给机器的行本来就是未做过的；全部录入后再按主体各查一次 NC 做后验验证。
- `config.json receipt_entry.excel` 新增 `start_row`、`result_sheet_name`、`currency_column`、`customer_code_column`、`fee_column`；`receipt_entry.validation_policy` 新增 `mode` 和 `skip_invalid_rows`。默认 `strict`，任意异常停止整批；`skip_invalid_rows` 只跳过异常行/重复组。
- `core.receipt_entry` 新增 `ReceiptPlanRow` 和 `ReceiptPlanIssue`。`ReceiptEntryWorkbook.build_local_plan()` 负责本地预检、重复识别和 Sheet2 输出；Sheet2 默认名为 `收款单自动化结果`。本日早期口径曾写“生成/覆盖”，后续已收敛为追加写入、缺表头补齐、旧噪音列按表头删除。
- 本地预检异常必须精确到 `原行号/阶段/异常类型/字段/原值/配置节点/说明/处理动作`，不能只写“配置错误”或“数据异常”。当前异常覆盖缺必需列、起始行无效、银行未配置、账户禁用、主体缺失、日期/金额/币种/手续费错误、客户编码为空和本批重复。
- 本地重复键为 `主体 + 到款日期 + 银行 + 币种 + 客户编码 + 银行来款名 + 金额`。重复组标 `DUPLICATE_EXCEL_ROWS`，整组不录入，避免在不查 NC 的前提下重复制单。
- `tools/receipt_entry_check.py` 默认改为新本地预检入口；`--write` 写 Sheet2，`--validation-mode skip_invalid_rows` 可临时跳过异常行/重复组。旧的“最近 N 个月 + 空状态列”候选逻辑保留为 `--legacy-candidates`，只作兼容诊断。
- 配置校验已纳入新字段和策略枚举；新增/修改收款配置后仍必须跑 `tools/validate_config.py`。

## 2026-06-05 - 当时接手重点

- 当前不要宣称收款单真实保存已完成。明细主行和手续费行 T0 已验证，真实 `Ctrl+S` 两案例保存循环还没跑通。
- 最新 blocker 是表头【客户】字段：`YW03200` 后台写入和一次屏幕兜底都出现读回为空。客户已设为硬门槛，表头写完后和保存前都必须非空，否则不写明细、不保存。
- 最新代码已把客户兜底改为：后台写入失败后，点击客户字段输入，再点击【单据日期】失焦提交，并轮询客户非空。这个版本尚未现场重跑，下一步应先验证它。
- `tools/tmp_receipt_two_case_save_run.py` 会真实保存两条测试单；GUI 动作前必须声明并等用户回复“开始/可以”。紧急停止键是空格。
- JAB 若出现 SunAwt 窗口存在但 `isJavaWindow=False/getContext=False`，优先判断为 NC Java 窗口未注册到 JAB；低风险恢复无效时重启 UClient/NC，不要继续调探索深度。

## 2026-06-04 - 收款单回写口径、主体实测和自制录入探测

- 严重踩坑：无标题小 `SunAwtWindow` 不一定是残留，`新增 -> 自制/应收单` 菜单和参照/下拉也可能是这种窗口。旧清理逻辑对可见小窗执行禁用/隐藏/移走后，会导致菜单像截图一样粘在窗口上、按钮点击无响应，属于阻塞级故障。
- `core/jab_operator.py` 修正 AWT 清理边界：JAB 启动/关闭时只做窄范围残留处理；打开菜单、参照、下拉的 JAB action 必须传 `cleanup_blank_awt=False`，动作后先找可见 enabled popup，再点其 JAB 控件。业务 popup 打开期间不能泛清。
- JAB action 默认值已收紧：`do_action()`、`do_action_by_path()`、`trigger_action_by_path_async()` 默认不再自动跑 AWT cleanup；`tools/jab_action_once.py` 也改为只有显式 `--cleanup-blank-awt` 才清理。
- 坐标/bounds 回退继续收紧：`click_context_center()` 现在直接拒绝真实点击，收款分页不再 action 失败后回退 bounds，配置校验禁止 `next_bounds_timeout` 和 `open_query.click_mode=bounds`。
- 收款单录入 `新增(Ctrl+N)` 按钮自身已确认可用；问题发生在按钮弹出的 `自制` / `应收单` AWT 菜单被旧清理逻辑误伤。健康状态下可见 popup 为 `SunAwtWindow`，`自制` 菜单项 path `0.0.1.0.0.0`，`应收单` path `0.0.1.0.0.2`，但 hwnd/path 只能作为现场复验依据，不能长期硬编码。
- `receipt_new_probe.py` 加入本次 popup 定向清理：点击 `新增` 前后做窗口差分，记录新出现且包含 `自制/应收单` 的可见 `SunAwtWindow`，点完 `自制` 后 hide/close 本次 popup，并显式清理所有无标题小型 `SunAwtWindow` 残留。该显式清理不再按 visible 区分；不可见残留也必须处理，否则后续 JAB/窗口状态容易出错。
- 收款单 `自制录入态` 已确认：进入自制后上方有 `保存(Ctrl+S)`、`暂存`、`取消(Ctrl+Q)` 三个按钮，可作为状态判据；在用户明确允许前不保存、不暂存。
- 新增收款明细表定位探测：`tools/receipt_body_table_locator.py` 按 25 列 body table 特征找明细表，空白录入态也能定位到 1 行 25 列表；JAB selection child index 规则为 `row * 25 + col`。
- 明细表当前关键列位按 0 基记录：`col=1` 收款业务类型，`col=3` 币种，`col=4` 收款银行账户，`col=5` 科目，`col=7` 金额/贷方原币金额，`col=11` 结算方式。已验证选中单元格不等于进入编辑器，`收款银行账户`、`结算方式` 等参照列不能靠剪贴板/键盘盲填，必须继续递归 popup/参照控件树。
- 业务口径补充：`结算方式` 上方字段会随下方明细表录入自动同步；自动化主路径应写明细表 `col=11`，上方字段仅用于同步校验。
- row 1424 自制录入进度：已从主窗口 `新增 -> 自制` 跑通到表头 `收款银行账户` 参照打开前；表头顺序修正为 `财务组织(O)` 写 `A001` 后先写 `客户` `YW03200`，触发财务组织名称和后续表头控件树加载，再写 `单据日期` `2026-04-02`、`币种` `美元`。表头文本框默认不强求立刻变中文或解锁下一字段，JAB 写入成功且控件 description/text 仍能读到写入值即可继续；NC 是否自行纠正只记录为状态。未保存、未暂存、未进入明细表写入。
- 失败记录：不要把 `财务组织` 后立即找不到 `单据日期` 误判为路径漂移；应先写 `客户` 让 NC 内部加载完成，再继续日期和币种。
- 失败记录：JAB 搜索深度 `25` 不够进入收款单表头深层控件；已把默认和配置改为 `max_depth=50`。深度 50 已确认可搜到 `客户`、`单据日期`、`币种`、`收款银行账户` label。
- 失败记录：收款单页签路径里的 `.2/.3` 段位是当前打开界面数量导致的动态索引，不是业务状态。不能因为用户打开另一个界面导致 `.2 -> .3` 就改路径口径；应先按 `收款单录入 / 客户收款单` 语义定位，再在深层子树找字段。
- 新严重踩坑：账户参照窗口虽然可打开，用户确认人工 `Alt+F` 可聚焦搜索，但自动化若用全局键盘发送 `Alt+F`/账号/Enter，且 `使用权参照` 没有成为前台，输入会落到用户聊天框或 VS Code。后续所有全局键盘输入必须先单独确认 `GetForegroundWindow()` 等于目标 hwnd；确认失败必须停，不能输入，也不能进入长等待。
- `receipt_self_made_fill_trial.py` 已加起点保护：`新增` 只允许匹配主窗口 `SunAwtFrame`，避免遗留 `使用权参照` 里的 `新增(N)` 被误当主按钮。若 `新增 -> 自制` 没确认成功，必须停在开单阶段。
- 自制录入脚本安全口径收紧：`新增 -> 自制` 只有在 `保存(Ctrl+S)`、`暂存`、`取消(Ctrl+Q)` 三按钮同时出现时才算开单成功；25 列明细表不再作为开单成功条件，只作为填明细前校验。
- `receipt_self_made_fill_trial.py` 撤销 `setTextContents` 失败后的全局键盘兜底；写字段失败必须停，不自动改用 `ctrl+a`/typing。
- `receipt_self_made_fill_trial.py` 表头验收改为后台优先：JAB 写入后先以控件 description/text 仍含写入值作为可继续证据；业务中文回显和后续控件可操作只作为更强状态记录，不作为财务组织等字段的硬门槛。Windows 前台不是 NC 只能阻止全局键盘输入，不能否定后台 JAB 已写入成功的字段。表头文本框默认不发送 Enter，不抢前台、不回退全局 `press_key`。
- `receipt_self_made_fill_trial.py --fill-detail` 的明细默认输入也收紧为后台 JAB 写表格；不再自动尝试剪贴板粘贴或 typing 兜底。
- `receipt_self_made_fill_trial.py` 默认不再把账户参照做完整闭环：表头账号阶段只打开 `使用权参照` 并返回 blocked，后续前台检查、`Alt+F` 搜索、等待结果、选择确定必须拆成独立阶段。明细填入也改为必须显式传 `--fill-detail`。
- `tools/close_awt_popup_residue.py --all-small` 用于显式残留清理：匹配无标题、小尺寸 `SunAwtWindow`，不区分 visible true/false，统一隐藏、移到 `-32000,-32000` 并发送 `WM_CLOSE`。旧 `--all-disabled-small` 仅保留为兼容别名，行为同 `--all-small`。
- `receipt_account_reference_try.py` 已加前台校验雏形，但不能再作为黑盒长命令直接跑。下一步要拆成“1 秒以内前台检查”和“用户确认后的 Alt+F 搜索”两个动作。
- 收款单查询条件的 `收款财务组织` 改为按 label 定位输入框写入，不再依赖旧固定 path、坐标范围或右侧未知按钮；默认不按 Enter，后续以查询结果是否切到目标主体作为验证。
- `tools/receipt_query_fill.py` 新增 `--include-filled-status`，用于重跑时覆盖已经写过的 `是否NC已做过` 状态；常规候选仍默认跳过已有状态。
- A001/移为已按 `2026-03-31` 至 `2026-05-31` 完成覆盖写回：NC 读取 `437` 行，Excel A001 候选 `407` 行，写回 `407` 行；其中唯一匹配 `315` 行、未找到 `4` 行、重复 `24` 行、异常/人工确认 `88` 行。
- A006 已在真实查询窗口验证成功：查询结果 `24` 行，Excel A006 候选 `24` 行，金额和对手方 `24/24` 唯一匹配，无重复、无未找到、无异常。
- A003 已聚焦测试：输入框可以写入 `A003`，但确认后结果仍保持 A001 口径的 `437` 行，说明 NC 没有采用 A003 查询条件；当前按权限或 NC 条件未生效处理，暂不允许据此写回。
- 收款单未唯一匹配的 Excel 回写说明已细化：金额匹配名称不一致、名称匹配金额不一致都会写明 `Excel` 值和 `NC` 候选值；完全无命中当时诊断 reason 为 `金额和对手方均未匹配`，当前实际写回状态已改为 `未做过`；重复写实际数量 `重复N条：名称和金额相同，需人工确认`。

## 2026-06-05 - 收款单明细主行屏幕写入验证

- 收款单明细主行 T0 验证通过：在用户确认的测试单据上，受保护屏幕输入可写第 1 行 `收款业务类型=货款`、`币种=美元`、`收款银行账户=FTE1219165931831`、`科目=1002`、`贷方原币金额=1090`、`结算方式=网银`；未保存、未暂存。
- 修正明细表点击坐标算法：旧逻辑把表格总高度当单行高，3 行表格时会点到第 2 行。新逻辑按 JAB 当前 `bounds / row_count / col_count` 动态计算单元格中心，并在测试输出里打印前 3 行读回。
- 主行字段顺序调整为 `收款业务类型 -> 币种 -> 收款银行账户 -> 科目 -> 贷方原币金额 -> 结算方式`。`结算方式` 放最后，最后点回第一个字段提交；默认不按 Enter/Tab，避免增行或移动到未知位置。
- 校验口径补充：`科目` 输入 `1002` 后 NC 回显 `1002\银行存款`，按编码前缀成功；金额 `1090` 回显 `1,090.00`，按金额归一化成功。
- 兼容边界：当前 T0 坐标按表格 bounds 动态计算，可适配窗口位置和整体尺寸变化，但仍假设 25 列可见且列宽未被用户拖动。正式化前应升级为按 JAB 单元格/列标题定位列，避免均分列宽假设。
- 手续费规则明确：只有 Excel 手续费非零时才允许进入手续费分支并增行；手续费行写 `收款业务类型=手续费`、`科目=660305`、`贷方原币金额=手续费金额`、`结算方式=网银`，收款银行账户列不填手续费。手续费为 0/空时禁止增行。增行入口改为稳定业务快捷键 `Ctrl+I`，不要用 Enter；必须先单独 T0 验证行数 +1、第 2 行写入和读回。
- `receipt_entry` 配置升级为面向后续配置维护/前端扩展的 `schema_version=2`：新增 `banks` 银行字典、账户稳定 `id/enabled/bank_id/display_name/excel_bank_aliases/nc_candidates_by_currency/entry_policy`、以及 `detail_entry_policy`。账户匹配只看账户自己的 label/aliases/excel_bank_aliases，银行字典只做分类，避免同一银行多主体账户时别名误匹配。
- 手续费分支 T0 已验证通过：主行完成后用 `Ctrl+I` 增行，新增手续费行写 `手续费 / 660305 / 金额 / 网银`。若手续费行自动带出收款银行账户，必须选中手续费行账户格后 `Delete` 清空；若写结算方式后 NC 自动多出空白第 3 行，选中第 3 行后用 `Ctrl+D` 删除。该流程已在测试单据上验证成功，未保存、未暂存。
- 收款银行账户明细主行改为优先直接填下方表格，不再走上方 `Alt+F` 参照作为主路径。原因：`Alt+F` 太慢且依赖前台；下方表格直接输入候选账号更快。账号候选从 `config.json receipt_entry.accounts[].nc_candidates_by_currency` 读取，例如招行人民币优先 `FTE1219165931831RMB`，不要把裸账号写死在脚本里。
- 真实保存循环 T0 新增 `tools/tmp_receipt_two_case_save_run.py`，计划连续保存两条测试单：无手续费和有手续费。保存成功 oracle 只看 `Ctrl+S` 后是否重新看到【新增】；不看提示框，不写 Excel，不关闭窗口。当前没有跑通保存闭环。
- 严重踩坑：表头客户字段不能沿用“JAB 写入返回成功但读回为空也软通过”的口径。真实测试中客户 `YW03200` 的 `setTextContents` 返回成功，但页面客户实际为空；脚本若软通过会继续写明细，存在误保存风险。已把客户改为硬门槛：表头写完后和保存前都必须检测【客户】非空，否则停止且不发送 `Ctrl+S`。
- 当前客户 blocker：客户字段能按 label/path 定位到，JAB 后台写入后 `text='' description=''`；随后屏幕兜底点击客户字段中心点并输入也成功发送过一次，现场输出 `target=[183,293] bounds=[97,283,172,20]`，但兜底后仍读回空。最新代码已改为客户屏幕输入后点击【单据日期】字段做失焦提交，并轮询客户非空；这个新版尚未现场重跑验证。
- 表头定位策略修正：收款单表头字段优先按 label 在当前可见 `SunAwtCanvas` 语义定位，固定深层 path 只作为兜底，并在输出中标记 `fallback_path_used`。不要再把路径里的 `.2/.3` 或 `.17.0` 当长期真相源。
- 真实保存 T0 的失败收敛已补强：底层 `SendInput failed` 会带 Windows error code；明细点击、`Ctrl+I`、`Delete`、`Ctrl+D`、`Ctrl+S` 都应返回结构化失败对象，避免只看到裸 `RuntimeError` 不知道失败在业务哪一步。
- JAB 注册失效排查结论：如果 SunAwt 窗口存在、Java 进程已加载 `JavaAccessBridge-64.dll`/`JAWTAccessBridge-64.dll`，但所有 `isJavaWindow=False` 且 `getAccessibleContextFromHWND=False`，说明当前 NC Java 窗口没有注册到 JAB。低风险 `Call info` 等动作未恢复；退出并重新启动 UClient/NC 后健康检查可恢复 `ok=True`。不要把这种情况误判为表格探测深度问题。
- 临时脚本交互口径修正：不要再用 `.bat`/额外 txt 作为主要反馈，T0 输出直接 `print` 中文说明；启动后统一 `sleep 2` 让用户切 NC；紧急停止键是空格。Windows 侧依赖缺失会直接导致 `ModuleNotFoundError`，已发现过 `keyboard`、`pyautogui`，运行前要确认 Windows Python 环境依赖齐。

## 2026-06-03 - 收款单查询结果抽取骨架

- 新增 `ReceiptNCResultExtractor`，按 NC 结果表表头语义抽取 `单据日期`、`原币金额`、`客户`，不依赖固定视觉列位。
- 收款单结果抽取会跳过空行；坏日期、坏金额、空客户会返回结构化抽取问题。
- `tools/receipt_query_fill.py` 默认会先用 F3 打开收款单查询条件窗口；`--no-open-query` 仅用于窗口已打开时的诊断。
- `tools/receipt_query_fill.py` 新增 `--read-results`，可在填条件或 `--confirm` 查询后读取可见 NC 表格并打印前 20 条抽取结果/问题。
- 收款单 dry-run 查询后会把结果页每页条数改为 500，并按分页逐页读取；总数超过 500 时通过“下页”继续采集。
- “下页”优先走 JAB 动作接口，失败时按同一路径做 bounds 点击回退；dry-run 输出会携带分页报告。
- 分页 path 查找新增 `scope_hwnd` 保护：收款查询结果区的 `SunAwtCanvas` 标题为空，必须先锁定分页 label 所在窗口，再在同一 hwnd 内改每页条数/点下页，避免串到其他 Java Canvas。
- 收款 dry-run 默认改为只读取匹配所需的 NC 结果列，降低大结果表读取耗时；这只是降耗优化，不再作为 NC 重启根因记录。
- 收款分页加入稳定等待：查询后、改每页条数后、读表前、下页后都等待结果表稳定，避免 NC 结果表刷新未完成时继续读表/翻页/改分页导致状态混乱。
- 收款分页等待从纯 sleep 升级为稳定判定：分页 label 和结果表摘要连续一致后才改每页条数、读表或翻页。
- 收款分页固定等待收紧为 0.5s 级短等待，主要依赖稳定判定放行。
- 收款分页去掉重复稳定等待：默认跳过改每页条数前的完整稳定轮询，并复用“改页后/下页后”的稳定结果作为后续读表依据。
- 收款 dry-run 修复分页聚合：多页采集到的同结构结果表会全部进入索引抽取，并按单据号去重，不再只取最大页表。
- 收款 dry-run 输出新增 NC 摘要：金额范围、名称样本和名称金额命中诊断；日期只用于查询范围，不参与匹配诊断。
- 收款单金额解析支持 NC 显示的 `- 1,368.10` 和 `(1,368.10)` 负数格式。
- 收款单结果列位按 NC/JAB 0 基索引配置；当前重新确认单据类型列 `2`、名称/客户列 `4`、原币金额列 `6`、本币金额列 `7`；`8` 是 Excel/人工 1 基口径误填，已加配置校验拦截。
- 已在真实 NC 跑通稳定判定分页：`659` 条记录按 `500 + 159` 两页采集完成，并全部进入 dry-run 索引抽取。
- 误判恢复记录：A001 写回验证曾按错误列位 `name=2`、`amount=7` 执行；后续确认当前页面就是目标页，但 `2` 是单据类型 `收款单`，`7` 是本币金额。已把 Excel A001 `2026-03-31` 之后 404 个自动写入状态清回空白，A001 候选恢复为 407 行；配置改为 `name=4`、`amount=6`。

## 2026-06-02 - 模型、契约和收款单查询准备

- `ExcelVoucherItem`、`PendingMatch`、`GeneratedVoucherMatch`、`VoucherPendingMatch`、`VoucherSaveMatch`、`MatchIssue` 收口为 dataclass。
- 删除模型的字典访问兼容层，workflow 改为属性访问。
- `ExcelVoucherItem` 增加处理前契约检查：无解析错误时必须有正 Excel 行号、金额和对手方。
- `VoucherSaveMatch` 增加保存前契约检查：制单表索引、表行数、制单行号和单元格内容必须有效。
- `match_generated_voucher_table` 显式返回 `GeneratedVoucherMatch`，不再复用待生成匹配类型。
- 待生成表重复匹配定义为异常；`generate` 默认暂停在点击 NC 前，显式传 `--on-duplicate skip` 时写入异常行并跳过继续。
- `config.json` 新增 `receipt_entry`，记录收款单录入状态标签、财务组织清单和组织-账户映射，并纳入配置校验。
- 收款单 Excel 预处理已支持按银行映射主体，候选行默认限定最近 2 个月且跳过已有状态。
- 收款单匹配规则已落地：Excel `原始金额 + 银行来款名` 对齐 NC `原币金额 + 名称列`；客户名做归一化和包含匹配；日期只用于查询/候选范围，不参与匹配；未唯一命中时区分“金额匹配但名称不一致”“名称匹配但金额不一致”“重复N条”“金额和对手方均未匹配”，不一致原因会带 Excel 值和 NC 候选值。
- 收款单查询窗口已枚举到右侧条件行，并将 `收款财务组织`、`单据日期`、`原币金额`、`客户` 的 JAB path 写入配置校验。
- 新增 `tools/receipt_query_fill.py`，已验证能在 `查询条件` 窗口写入 `A001` 和日期区间；默认不点确定，`--confirm` 才触发查询。
- 修复 `tools/jab_probe.py --inspect-path` 忽略 `--depth` 的问题；该问题曾导致右侧查询条件区被误判为空。

保留坑点：

- `--inspect-path` 的输出路径是相对被 inspect 的目标节点，不是完整窗口路径；配置 JAB path 前必须用完整 path 复验。
- `setTextContents` 对 NC 查询条件文本框可写入，但 JAB 文本读回可能为空，不能仅凭读回为空判失败。

## 2026-06-01 - workflow 拆分和状态守卫

- 新增页面状态识别和状态守卫：`pending`、`generated`、`voucher_open`、`query_open`、`loading`、`error`。
- 页面状态识别同时检查父页面标签、按钮布局、表格数据特征；制单、查询按阻塞式子窗口处理。
- 增加状态事件和状态跳转记录。
- 新增 workflow 领域异常：`WorkflowStateError`、`TableMatchError`、`ContractViolation`、`JABControlNotFound`、`JABActionError`。
- 架构检查阻止 workflow 模块新增裸 `raise RuntimeError(...)`，避免业务失败原因继续混成一种异常。
- 新增 `ExcelLockedError`，Excel 拆分列、生成状态写入、凭证号回填写入遇到文件占用时统一报清楚。
- `backfill` 默认会从 `pending` 自动切到 `generated`；遇到制单、查询、加载、异常状态会停止，不再直接读表。
- `backfill` 增加结构化审计事件 `backfill_audit`，记录 Excel 行、金额、对手方、NC 行、凭证号和失败状态。
- 将原 `JABBatchProcessor` 拆成：
  - `core/nc_state.py`
  - `core/nc_page_probe.py`
  - `core/nc_pending_workflow.py`
  - `core/nc_voucher_workflow.py`
  - `core/nc_switch_generated_workflow.py`
  - `core/nc_backfill_workflow.py`
  - `core/nc_table_matcher.py`
- `JABBatchProcessor` 收敛为装配入口，只保留 CLI 任务级方法和共享运行状态。
- 清理重构过渡期的纯转发 wrapper。

当前结论：没有继续低风险拆分项；后续结构演进优先做 dataclass 模型和契约检查。

## 2026-05-28 - 查询切换和保存策略收敛

- `tools/jab_batch.py` 增加 `--generated-date YYYY-MM-DD`。
- `generated_date_value` 优先级：命令行参数、`config.json`、当天日期。
- `switch-generated --perf` 细分查询窗口内步骤耗时。
- `switch-generated` 主查询入口保持 F3；查询窗口内部 `正式单据`、`确定` 走 JAB AccessibleAction，日期框走 `setTextContents`。
- 增加 path guard，F3 后等待查询窗口内目标控件出现，不再固定 sleep。
- `目的业务日期` 是 `介于` 条件，限定当天必须同时填写起始和结束两个日期框。
- 补充验证：`目的业务日期` 依赖先选择 `正式单据` 才出现；选择前的 `.2` 行是 `生效日期`，正式查询日期仍使用 `.11` 行。
- 删除无继续维护价值的 `batch_reverse_select` 和 `batch` 保存策略。
- 正式保存主线收敛为 `single + jab_button + use_voucher_queue_cache`。
- 保留 `safe_batch_by_pending_row` 作为快速备选，但不承诺凭证号严格按 Excel 递增。

保留坑点：

- JAB `bounds` 不是底层动作，不能作为主路径恢复坐标点击。
- `ok=True` 只能代表 JAB 动作返回成功，业务上仍需要后置状态验证。
- `setTextContents` 对日期框写入有效，但 JAB 文本读取可能返回空。
- 隐藏或 `visible=False` 的查询窗口可能仍能被 JAB 枚举到，不应作为可操作窗口依据。
- NC 查询条件区视觉换行不一定对应 JAB 结构换行。

## 2026-05-27 - JAB 批量凭证生成主线成型

- Java Access Bridge 成为当前主线方案。
- 旧 `pyautogui` 坐标点击、截图识别、固定坐标方案停止维护，旧入口和旧 GUI 模块删除。
- 开发迁移到 WSL 仓库 `/home/queclink/project/nc_auto_v2`，H 盘目录作为 Windows/JAB 运行镜像。
- 明确 JAB 实际操作 NC 时必须调用 Windows Python。
- 新增/确认命令：`plan`、`generate`、`resume-voucher`、`switch-generated`、`backfill`、`split-keys`。
- 生成阶段采用全量匹配、一次性选中、只点一次 `生成 -> 前台生成`、在同一个 `制单` 窗口按 Excel 顺序保存。
- 保存后不能只信提示，必须验证制单表行消失/行数减少/空表，并最终回待生成表 F5 复核。
- Excel 索引始终是 `金额 + 对手方`，不是单据号，也不是单金额。

保留坑点：

- JAB 不能由 WSL/Linux Python 直接运行，必须调用 Windows Python。
- JAB path 和 hwnd 不稳定，不要长期硬编码。
- 开启 JAB 后可能出现左上角空白蓝框/截图样遮挡窗口，通常是 `SunAwtWindow` 无标题小窗口；后续已确认清理必须显式触发：业务 popup 打开期间不能泛清，完成选择后对无标题小型残留统一清理，不按 visible 区分。
- Excel/WPS 打开文件时，写入 C 列或自动拆分 A/B 时可能报 `PermissionError`。
- 多行批量保存时，凭证号顺序可能不等于 Excel 顺序。严格顺序场景默认使用 `single`。

真实反序案例：

```text
excel_rows=[25, 26]
voucher_rows=[1, 9]
Excel行25 -> 370
Excel行26 -> 369
```
