# NC JAB 凭证自动化

这个项目用于在 NC6.5 中按 Excel 顺序批量生成凭证，并把已生成凭证号回填到 Excel。

当前主线是 Java Access Bridge（JAB）。旧的 `pyautogui` 坐标点击、截图识别、固定坐标方案已经删除，不再作为新功能方向。

## 开发和运行边界

开发已经可以完全转到 WSL：

- WSL 源码仓库：`/home/queclink/project/nc_auto_v2`
- H 盘 Windows/JAB 运行镜像：`/mnt/h/python脚本/.venv/nc_auto_v2`
- JAB 运行 Python：`/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe`

注意：JAB 不能由 WSL/Linux Python 直接控制 NC。JAB 要连接 Windows 桌面上的 Java/NC UI，依赖 Windows Access Bridge DLL、窗口句柄和 Windows 消息泵。所以：

- 源码、git、文档、重构在 WSL。
- 实际读 NC 表格、选行、点击按钮时，必须调用 Windows Python。
- 可以从 WSL shell 调用 Windows Python，但解释器必须是 `.venv-local/Scripts/python.exe`。

## 关键文件

- `tools/jab_batch.py`：批量命令入口。
- `core/jab_batch_processor.py`：批量流程装配入口，只保留 CLI 任务级入口和共享运行状态。
- `core/nc_state.py` / `core/nc_page_probe.py`：NC 页面状态识别和 JAB 页面特征探测。
- `core/nc_pending_workflow.py`：待生成列表匹配、生成入口、恢复当前制单窗口。
- `core/nc_voucher_workflow.py`：制单窗口匹配、保存、关闭后验证。
- `core/nc_switch_generated_workflow.py`：切换到已生成/正式单据列表。
- `core/nc_backfill_workflow.py`：已生成列表凭证号回填。
- `core/nc_table_matcher.py`：NC 表格按金额、对手方、日期的匹配逻辑。
- `core/receipt_entry.py`：收款单 Excel 本地预检、主体映射、Sheet2 机器结果表和 NC 后验匹配基础模型。
- `core/errors.py`：NC workflow 领域异常，区分页面状态、表格匹配、JAB 控件、JAB 动作、Excel 写入锁和流程契约失败。
- `core/jab_operator.py`：JAB 底层封装，负责读表、选行、按钮动作、F3/F5、关闭窗口、受保护的 AWT 残留清理。
- `core/data_handler.py`：Excel 读取、拼接 key 解析、结果写回、拆分 key。
- `config.json`：Excel 路径、JAB DLL、列位和查询切换配置。
- `TODO_JAB_HANDOFF.md`：后续开发 TODO。
- `CHANGELOG.md`：已实现功能、验证记录和历史变更。

辅助探测工具：

- `tools/jab_probe.py`
- `tools/receipt_body_table_locator.py`
- `tools/receipt_reference_cell_probe.py`
- `tools/query_jab.bat`
- `tools/run_jab_probe.bat`
- `tools/test_voucher_selection.py`

## WSL 开发流程

在 WSL 仓库开发：

```bash
cd /home/queclink/project/nc_auto_v2
git status --short
```

实际操作 NC 前，同步到 H 盘运行镜像：

```bash
tools/sync_to_windows.sh
```

需要持续自动同步时：

```bash
tools/sync_to_windows.sh --watch
```

先看同步会改什么，不实际写入：

```bash
tools/sync_to_windows.sh --dry-run
```

进入 H 盘镜像目录，用 Windows Python 运行：

```bash
cd /mnt/h/python脚本/.venv/nc_auto_v2
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py plan
```

## 常用命令

只读规划，不点击 NC：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py plan
```

真实生成并保存：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py generate --yes
```

切到已生成/正式单据列表：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py switch-generated
```

按指定目的业务日期切到已生成/正式单据列表，并记录性能：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py switch-generated --generated-date 2026-05-27 --perf --perf-label switch-20260527
```

从已生成列表回填凭证号：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py backfill
```

只校验当前页面，不从待生成页自动切换：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py backfill --no-backfill-auto-switch
```

拆分 Excel A 列拼接 key：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py split-keys
```

收款单本地预检主入口。该命令从 `receipt_entry.excel.start_row` 开始读 Sheet1，只做 Excel/配置侧识别和异常拦截，不查 NC；异常会精确输出到行号、字段、原值、配置节点和处理动作：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_entry_check.py
```

写入机器生成的 Sheet2 结果表，默认表名为 `收款单自动化结果`：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_entry_check.py --write
```

正式口径默认 `strict`：发现任意异常整批停止。临时允许跳过异常行/重复组继续生成可运行计划时：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_entry_check.py --validation-mode skip_invalid_rows --write
```

旧的“最近 N 个月 + 是否NC已做过为空”候选预览只保留为兼容诊断入口：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_entry_check.py --legacy-candidates
```

收款单查询窗口只填条件、不点确定；默认会先在收款单录入页按 F3 打开查询条件窗口：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-05-01 --date-to 2026-06-02
```

收款单查询后读取可见结果表：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-05-01 --date-to 2026-06-02 --confirm --read-results
```

收款单查询后只读匹配预演；查询后会把每页条数改为 500，并按分页读取。输出 JSON 包含 `page_report`、金额范围、名称样本和重复原因；日期只用于查询范围，不参与匹配诊断。注意：这是历史/诊断工具，不是当前新主线的录入前筛选步骤：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-03-31 --date-to 2026-05-31 --confirm --dry-run-match --max-rows 600 --max-cols 140
```

收款单匹配结果写回 Excel；唯一匹配写 `已做过`。金额和对手方都没有命中时写 `金额和对手方均未匹配`；金额命中但名称不符、名称命中但金额不符会写明 Excel 值和 NC 候选值；重复命中按实际条数写 `重复N条：名称和金额相同，需人工确认`，重复行也会在 JSON 的 `duplicate_rows` 中报告。注意：该写回旧入口不再作为新批量录入主线的前置判断，当前用户口径是假定交给机器的行均未做过，录入完成后再按主体查询 NC 做后验验证：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-03-31 --date-to 2026-05-31 --confirm --dry-run-match --write-back --max-rows 600 --max-cols 140
```

重跑并覆盖 Excel 已有 `是否NC已做过` 状态时，加 `--include-filled-status`：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-03-31 --date-to 2026-05-31 --confirm --dry-run-match --write-back --include-filled-status --max-rows 600 --max-cols 140
```

写回前必须先确认当前 NC 页面是目标 `收款单录入`。工具会做页面和结果表守卫；如果结果表的匹配名称列疑似读成 `收款单` 单据类型列，会拒绝继续。

`收款财务组织` 当前按字段 label 聚焦右侧输入框后直接写组织编码，例如 `A001`、`A006`。不要点输入框右侧未知按钮；不要为了提交该输入框额外按 Enter，除非现场已确认该主体必须这样做。写入是否有效以后置查询结果和 Excel 主体匹配数为准，不能只看 JAB 文本读回。

## NC 现场操作纪律

自动化和用户共用同一个桌面输入焦点。任何 `pyautogui.write`、`pyautogui.hotkey`、剪贴板粘贴、普通按键都可能打进用户聊天框或 VS Code。执行这类全局键盘输入前，必须先单独确认目标窗口是前台窗口；如果 `GetForegroundWindow()` 不是目标 hwnd，脚本必须停，不允许发送快捷键、账号或 Enter。

默认录入路径必须优先走后台 JAB：`setTextContents`、JAB action、selection/table API 和 JAB 控件树读回。前台窗口只用于判断能不能发送全局键盘输入，不能用来否定已经由后台 JAB 写入并回显成功的业务状态。

表头文本框默认只做后台写入并继续，不强求立刻变中文或解锁下一字段。`财务组织(O)` 写 `A001` 后，如果后台 description/text 仍能读到 `A001`，就先写 `客户`，用客户输入触发财务组织名称和后续表头控件树加载，再写 `单据日期`、`币种`；如果 NC 自行把描述纠正成 `上海移为通信技术股份有限公司`，只作为更强状态记录。不得因为前台仍是 VS Code 就否定后台写入成功。

例外：`客户` 是保存前硬门槛，不能按普通表头字段软通过。2026-06-05 真实测试确认，客户 `YW03200` 曾出现 `setTextContents` 返回成功但页面仍为空；因此收款单自制保存流程必须在表头写完后和保存前分别确认【客户】非空。客户为空时必须停，不写明细、不发送 `Ctrl+S`。当前 T0 已加入后台写入失败后的受保护屏幕输入兜底：点击客户字段、全选、输入客户编码，再点击【单据日期】字段失焦提交并轮询非空；该最新版尚未完成现场跑通。

NC 表头层级很深，JAB 默认搜索深度必须保持 `max_depth=50`。`25` 层只能看到 `客户收款单` 页签和浅层财务组织，搜不到 `客户`、`单据日期`、`币种`、`收款银行账户` 等真实表头 label，会造成误判。

收款单页签路径里的 `.2/.3` 这类段位会随当前打开的 NC 界面数量变化；用户打开另一个界面后当前收款页可能从 `.2` 变成 `.3`。这不是业务状态，也不是失败原因，不能把固定页签索引当长期依据。路径只能作为现场短期定位，主逻辑必须先确认当前 `收款单录入 / 客户收款单` 语义，再在足够深度的子树里找字段。

表头字段同样不能固定深层 path 作为主路径。当前实现应先按 label 在当前可见 `SunAwtCanvas` 里语义定位，固定 path 只作为兜底，并在输出里标明。看到 `fallback_path_used` 时只能说明这次现场兜底命中，不能把该 path 固化进正式流程。

不要把“前台化、输入、等待结果”合成一个长命令。先做短前台检查，再执行用户已确认的动作。遇到遗留 `使用权参照`、NC 主窗口在屏幕外、`新增` 匹配到非主窗口、或任何状态污染时，只报告状态，不继续补救性点击或写字段。

收款单自制录入当前已从 row 1424 阶段推进到明细/手续费 T0：`新增 -> 自制`、明细主行、手续费行、手续费账户清空、多余空行删除都已验证；真实 `Ctrl+S` 保存闭环尚未跑通。最新 blocker 是表头【客户】非空：客户后台写入和一次屏幕兜底均出现读回为空，代码已改为客户输入后点击【单据日期】失焦提交并轮询，但该版本还未现场重跑。没有明确允许时禁止保存、暂存。

`新增 -> 自制` 的成功标准是上方 `保存(Ctrl+S)`、`暂存`、`取消(Ctrl+Q)` 三个按钮同时出现；只看到 JAB action 返回成功或菜单项被点击，不算进入自制录入态。25 列明细表只作为填明细前的单独校验，不作为开单成功条件。

`tools/receipt_self_made_fill_trial.py` 默认只负责开单和表头阶段；账户参照打开后会停止，后续前台检查、`Alt+F` 搜索、等待结果、选择确定必须拆成独立动作。明细填入默认禁用，必须显式传 `--fill-detail` 才会尝试进入明细；即使显式填明细，默认也只允许后台 JAB 写表格，不自动退到剪贴板粘贴或 typing。

`tools/tmp_receipt_two_case_save_run.py` 是当前真实保存循环 T0。它会跑两条测试单：无手续费和有手续费；保存前会双检客户非空；成功 oracle 是保存后【新增】重新出现。该脚本会真实 `Ctrl+S` 保存，运行前必须向用户声明并等用户确认。若当前页面是上一轮失败留下的自制页，第一条允许复用当前页覆盖测试值；后续案例必须从保存后【新增】重新开始。

明细表当前已验证的填写顺序：主行写 `收款业务类型=货款`、`币种`、`收款银行账户`、`科目=1002`、`贷方原币金额`、`结算方式=网银`。结算方式必须放最后，最后点回第一个字段提交，不按 Enter/Tab。手续费非零时才 `Ctrl+I` 增行，手续费行写 `手续费`、`660305`、手续费金额、`网银`；手续费行账户必须为空，自动带出时用 `Delete` 清空；若多出空白第 3 行，用 `Ctrl+D` 删除。

AWT 小窗清理分两类处理：业务 popup 打开期间禁止泛清，必须先在当前 popup 控件树里完成选择；但 `新增 -> 自制` 选择完成后要显式清理本次菜单和所有无标题小型 `SunAwtWindow` 残留。残留清理不按 visible 区分，不可见小窗也要清，否则后续 JAB/窗口状态可能出错。手工清理用 `tools/close_awt_popup_residue.py --all-small`；旧 `--all-disabled-small` 只是兼容别名。

## Excel 规则

当前 Excel：

```text
C:\Users\Queclink\Desktop\6.1凭证.xlsx
```

`Sheet1` 约定：

- A 列：金额，或临时追加的 `金额+对手方` 拼接索引。
- B 列：对手方。
- C 列：状态/凭证号。
- 同一张表里允许部分行是 A/B 拆分，部分行是 A 列拼接；程序逐行判断。

示例：

```text
141688.50深圳市科贸电子科技有限公司
```

配置项：

- `sheet_my = Sheet1`
- `has_header = true`
- `jab_batch.key_col = 1`
- `jab_batch.amount_out_col = 1`
- `jab_batch.partner_out_col = 2`
- `jab_batch.result_col = 3`

解析规则：

- 每行优先判断 A 列是不是 `金额+对手方` 拼接索引；是则按拼接解析。
- A 列不是拼接时，按 A 列金额、B 列对手方解析。
- A 列拼接索引必须以金额开头。
- 金额允许逗号和小数，统一到 `Decimal("0.01")`。
- 对手方取金额后的全部文本，并去掉空白。
- 当前索引始终是 `金额 + 对手方`，不是单据号，也不是单金额。
- `generate` / `backfill` 读取到 A 列拼接 key 后，会自动拆回 A/B 列。

状态/凭证号列规则：

- C 列统一承担状态/凭证号写入。
- `plan` / `generate` 跳过 C 列已有值的行。
- 生成成功但未回填凭证号时，C 列写 `已生成待回填`。
- `backfill` 只读取 C 列等于 `已生成待回填` 的 key 行，并替换为凭证号。

## NC 表格列位

除 Excel/openpyxl 读写使用 1 基列号外，NC/JAB 表格列位一律按 0 基索引记录。

待生成主表：

- 金额列：`col=4`
- 对手方列：`col=3`
- 选择列：`col=0`

制单窗口：

- 标题：`制单`
- class：`SunAwtDialog`
- 表格通常是 `N rows x 13 cols`
- 匹配方式：金额匹配，并在整行文本里查找对手方。

已生成/正式单据表：

- 金额列：`col=4`
- 对手方列：`col=3`
- 凭证日期列：`col=18`
- 凭证号列：`col=22`
- 凭证号回填时去前导 0。
- 凭证号有效范围：`1 <= 凭证号 <= 9999`。

收款单查询结果表：

- 单据类型列：`col=2`，常见值是 `收款单`，不能作为匹配名称列。
- 名称/客户列：`col=4`
- 原币金额列：`col=6`
- 本币金额列：`col=7`，不能用于 Excel 原始金额匹配。
- `col=8` 是把 1 基列号误填到 NC/JAB 配置里的旧错误值，不要使用。

收款单配置 schema：

- `receipt_entry.schema_version=2` 是面向后续 GUI 的配置模型；GUI 应编辑业务对象，不要把银行、账号、快捷键继续写死到脚本里。
- `receipt_entry.excel.start_row` 是 Sheet1 业务起始行；新主线从该行开始读取，不再用“最近 2 个月”或 `是否NC已做过` 状态列决定本批候选。
- `receipt_entry.excel.result_sheet_name` 是机器生成 Sheet2 名称，默认 `收款单自动化结果`。Sheet2 只由机器生成/覆盖，人工不维护。
- `receipt_entry.excel.currency_column`、`customer_code_column`、`fee_column` 分别指定币种、客户编码、手续费列，用于本地预检和后续录入。
- `receipt_entry.validation_policy.mode` 支持 `strict` 和 `skip_invalid_rows`。`strict` 有任意异常就停止整批；`skip_invalid_rows` 会把异常行/重复组写 Sheet2 并跳过。
- `receipt_entry.banks` 是银行字典，只做展示、分类和可选别名维护；同一家银行可能有多个主体账户，银行别名不能自动当账户匹配规则。
- `receipt_entry.accounts` 是账户字典，每个账户必须有稳定 `id`、`enabled`、`organization_code`、`bank_id`、`account_label`、`account_no`。Excel 银行列匹配只看账户自己的 `account_label`、`aliases`、`excel_bank_aliases`。
- `nc_candidates_by_currency` 用于配置 NC 可输入候选，例如人民币账户优先 `...RMB`，避免代码继续猜 `RMB/USD/CNY` 后缀。
- `entry_policy` 记录账户录入策略：当前主线是 `account_input=detail_first`、`success_rule=non_empty`、必要时 `fallback_reference=true`。
- `detail_entry_policy` 记录明细主行/手续费行顺序和快捷键：手续费增行 `ctrl+i`，手续费账户清空，额外空行删除 `ctrl+d`。
- `tools/validate_config.py` 会校验组织、银行、账户引用、账户别名冲突、候选值和策略枚举；新增银行或账号后必须先跑配置校验。

收款单本地预检异常口径：

- 本地预检由 `ReceiptEntryWorkbook.build_local_plan()` 执行。输出模型是 `ReceiptPlanRow` 和 `ReceiptPlanIssue`。
- `ReceiptPlanIssue` 必须带 `excel_row`、`stage`、`issue_type`、`field`、`raw_value`、`config_node`、`message`、`action`；不要新增笼统的“配置错误/数据错误”。
- 当前会识别缺必需列、起始行无效、银行为空/未配置/账户禁用、账户主体不存在、日期错误、银行来款名为空、金额错误或非正、币种为空或不支持、客户编码为空、手续费错误或负数、以及本批 Sheet1 内重复。
- 重复键为 `主体 + 到款日期 + 银行 + 币种 + 客户编码 + 银行来款名 + 金额`。同一 key 出现多行时标 `DUPLICATE_EXCEL_ROWS`，整组不录入。
- 通过预检的行按主体分组供后续录入使用；录入前不查 NC。全部录入后，每个主体只查询一次 NC 做后验验证，并把结果写回 Sheet2 的后验列。

收款单自制录入明细表：

- 当前用 `tools/receipt_body_table_locator.py` 按 25 列 body table 特征定位，空白录入态通常是 1 行 25 列，旁边另有独立的合计表。
- JAB selection API 的 child index 规则是 `row * 25 + col`。
- 已探测关键列位按 0 基记录：`col=1` 收款业务类型，`col=3` 币种，`col=4` 收款银行账户，`col=5` 科目，`col=7` 金额/贷方原币金额，`col=11` 结算方式。
- 选中表格单元格不等于键盘焦点进入编辑器。`col=11` 结算方式曾在 Enter/F2 后把焦点送到右上角全局搜索框，说明不能用剪贴板或全局键盘盲填参照列。
- `结算方式` 上方也有输入框，但下方明细表录入后会自动同步上方字段；主路径应写明细表 `col=11`，上方字段只作为同步验证，不作为优先录入口。
- `收款银行账户`、`结算方式` 这类列要继续递归参照/下拉 popup 控件树，找到真实可操作控件后再写入；`setTextContents` 成功或读回为空都不能单独作为业务成功依据。

## 业务流程

### 规划

`plan` 只读 Excel 和 NC 待生成表，不点击 NC：

1. 读取 Excel A 列 key。
2. 跳过 C 列已有状态/凭证号的行。
3. 解析金额和对手方。
4. 读取当前 NC 待生成主表。
5. 按 `金额 + 对手方` 精确匹配。
6. 输出可匹配、格式错误、未找到、重复项。

### 生成

`generate --yes` 的正确流程：

1. 从 Excel 读取全部待处理行。
2. 在 NC 待生成主表一次性匹配所有 Excel 行。
3. 一次性选中所有匹配到的 NC 待生成行。
4. 只点击一次 `生成 -> 前台生成`。
5. 进入同一个 `制单` 弹窗。
6. 在制单表中按 Excel 顺序查找记录。
7. 选择当前可保存批次并点击保存。
8. 验证制单表行数减少、目标行消失，或制单表为空。
9. 对已保存 Excel 行写 `已生成待回填`。
10. 全部保存后关闭制单窗口。
11. 回待生成表按 F5 刷新。
12. 验证本轮所有记录已从待生成表消失。

关键结论：

- 第一阶段必须先全量选中 Excel 匹配到的待生成行。
- 只点一次前台生成。
- 不要每条 Excel 单独生成。
- 不要拆成多轮待生成表 `选几行 -> 生成 -> 前台生成`。

原因是只选 1 行时，NC 制单界面可能和多行场景不同；反复前台生成会让界面状态、顺序和验证逻辑变复杂。

### 切换已生成

`switch-generated` 从待生成界面切到已生成/正式单据列表：

1. 如有 `制单` 窗口，先尝试关闭。
2. 激活 NC 主窗口。
3. 用 F3 打开查询窗口。
4. 等查询窗口内 `正式单据` 控件出现。
5. 用 JAB AccessibleAction 选择 `正式单据`。
6. 等 `目的业务日期` 条件出现。
7. 用 JAB `setTextContents` 写入两个日期框。
8. 用 JAB AccessibleAction 点击 `确定`。
9. 读取表格确认已进入已生成列表。

日期默认优先级：

1. 命令行 `--generated-date YYYY-MM-DD`。
2. `config.json` 中 `jab_batch.generated_date_value`。
3. 当天日期。

已生成和待生成占同一个业务窗口。不要每生成一张就来回切；全部生成后统一切到已生成表回填。

JAB 查询入口现状：

- 已验证单据生成页存在 JAB 查询按钮：
  `0.0.1.0.0.0.0.2.0.0.0.0.0.0.0.2`
- 控件信息：`查询` / `push button` / `单击`。
- 该 path 可以打开查询窗口。
- 但主查询入口用 JAB action 会触发 Access Bridge 不稳定，因此默认仍用 F3 打开查询窗口。
- 查询窗口内部已验证：`正式单据` / `确定` 走 JAB AccessibleAction，日期框走 JAB `setTextContents`。
- `目的业务日期` 条件操作符是 `介于`，限定当天时两个日期框都要填同一天。
- F3 后不要固定 sleep 盲点，必须等查询窗口内目标 path 出现后再执行下一步。
- 开启 `--perf` 后，查询窗口步骤会拆分记录 `switch_step_formal_action`、`switch_step_date_from`、`switch_step_date_to`、`switch_step_confirm_action`。
- NC 查询条件区的视觉布局不等于 JAB 结构：有些输入框看起来像上下两行，实际可能像同一个 div/容器里换行显示，JAB 仍会归到同一个容器或同一行组。定位时必须按 label、role、row 容器、bounds 和后置状态验证综合判断。

### 回填

`backfill` 从已生成表回填凭证号到 Excel C 列。默认会先识别页面状态：

- 已在 `generated`：直接回填。
- 当前是 `pending`：先自动执行 `switch-generated`，再回填。
- 当前是 `voucher_open`、`query_open`、`loading`、`error`：停止并报告状态，避免按错误表格列位读取。
- 如需旧的只校验行为，使用 `--no-backfill-auto-switch`。

逻辑：

1. 读取 Excel 中 C 列等于 `已生成待回填` 的 key 行。
2. 读取已生成表。
3. 按 `金额 + 对手方` 匹配。
4. 历史重复时，优先取凭证日期等于当天的记录。
5. 读取凭证号列。
6. 去前导 0。
7. 校验 `1-9999`。
8. 写回 Excel C 列。

回填会在 `logs/run_state.json` 记录 `backfill_audit` 事件；开启 `--perf` 时，也会写入性能 JSONL。审计记录包含 Excel 行、金额、对手方、NC 已生成表行、原始凭证号、写回值和失败状态。

## 制单保存策略

当前正式主线是 `single + jab_button + use_voucher_queue_cache`：

- 一张一保存，保证凭证号按 Excel 顺序递增。
- 保存触发走 JAB 保存按钮，稳定性高于快捷键。
- 制单表只读取一次初始队列；每保存一张后，缓存中位于被删行下方的行号统一减 1。
- Excel 状态在保存结束后批量写入，减少 I/O。

代码只保留稳定主线、一个快速备选策略和旧兼容策略。

可配置策略：

- `single`：当前正式策略。一行一保存，凭证号顺序最稳，速度较慢。
- `bottom_up`：旧策略，只合并制单行号严格递减的连续 Excel 行。已保留但不再默认。
- `safe_batch_by_pending_row`：快速备选策略。只合并 Excel 顺序中待生成 NC 行号递增、且制单窗口行号递增的连续段；其它自动拆成单张。它整体线性，但不承诺凭证号严格按 Excel 递增。

已发生真实案例：Excel 行 25/26 的制单行分别是 1/9，但回填凭证号为 370/369。后续多轮实验也证明行号递增、递减、选择顺序都不能稳定保证凭证号按 Excel 顺序递增。因此严格顺序主线固定使用 `single`。

`Ctrl+S` 保存触发也做过对照：可保存成功，但端到端没有优于 JAB 按钮。因此默认仍使用 JAB 按钮。

## 验证规则

不要只信“保存成功”提示。

保存后当前验证：

1. 看制单窗口是否还存在。
2. 读取制单表。
3. 目标行消失且行数减少，判定当前批保存成功。
4. 制单窗口还在但表为空，视为本轮可能完成。
5. 制单窗口关闭，也转待生成表复核。
6. 最终关闭制单窗口。
7. 回待生成表按 F5 刷新。
8. 按 `金额 + 对手方` 验证本轮记录是否仍存在。

`制单` 窗口还在但表为空是正常状态，通常表示本轮选中的制单数据已保存完。正确处理是关闭窗口、刷新待生成表、再复核。

## JAB 注意事项

- JAB 只能在 Windows Python 下连接 NC UI。
- 老版 JAB 使用 `Windows_run()`；不要随手改 `tools/jab_probe.py` 和 `core/jab_operator.py` 的消息泵/初始化方式。
- 不要硬编码 JAB path 或 hwnd，窗口重开后会变。
- 不要信 `bounds` 做坐标点击，很多控件会返回负坐标或 `-1,-1,-1,-1`。
- NC 自动化彻底禁用坐标点击、bounds 中心点点击和截图找坐标；只能用 JAB action、语义快捷键、JAB 文本/表格 API。
- 优先使用 JAB action、selection、table API。
- 隐藏或非 visible 的 `SunAwtDialog` 查询窗口可能是残留，不应作为可操作窗口依据。

左上角奇怪截图/蓝框遮挡：

- 开启 JAB 后曾多次出现左上角小窗口，表现为“NC 的截图还去不掉”“蓝框里边空白”“没有标题”。
- 严重注意：无标题小 `SunAwtWindow` 不一定是残留，`新增 -> 自制/应收单` 菜单、参照框、下拉框也可能是同类窗口。
- 旧清理逻辑曾把可见菜单禁用/隐藏/移走，造成菜单像截图一样粘住、按钮点击无反应。这是阻塞级故障，不是普通视觉瑕疵。
- 当前规则：业务 popup 打开期间不泛清；完成选择后才显式清理。`新增` 前快照窗口，`新增` 后记录本次出现且包含 `自制/应收单` 的可见 `SunAwtWindow` hwnd，点完 `自制` 后 hide/close 这个 hwnd，并清理所有无标题小型 `SunAwtWindow` 残留。残留不按 visible 区分，不可见小窗也清。
- `core/jab_operator.py` 的 `hide_blank_awt_windows()` 只作为残留清理入口，不能在业务 popup 刚打开时调用。会打开 popup 的动作应显式传 `cleanup_blank_awt=False`。
- JAB action 默认不自动清 AWT 残留；`tools/jab_action_once.py` 只有显式 `--cleanup-blank-awt` 才会触发清理。
- bounds 点击已经在底层拒绝，配置校验也禁止分页 `next_bounds_timeout` 和查询入口 `click_mode=bounds`。如果 JAB action 不可用，应继续探测控件树，而不是回退坐标。
- 已验证 `新增(Ctrl+N)` 按钮本身可用，随后会弹出 `自制` / `应收单` 菜单；健康 popup 中 `自制` path 曾为 `0.0.1.0.0.0`，`应收单` path 曾为 `0.0.1.0.0.2`。path/hwnd 不稳定，只能现场复验后使用。
- 遇到截图样残留时不要点它、不要按坐标关它；先确认它不是当前业务菜单/参照。如果已经完成当前 popup 选择，就用 `tools/close_awt_popup_residue.py --all-small` 显式清理无标题小型 `SunAwtWindow`，包括不可见残留，然后复扫 AWT 窗口。

## 常见问题

Q: 现在开发是否可以完全转到 WSL？

A: 可以。源码、git、文档、重构都放 WSL。只有实际控制 NC/JAB 的 Python 进程必须是 Windows Python。

Q: 为什么不能用 WSL Python 直接跑 JAB？

A: WSL Python 看不到 Windows 桌面上的 Java 窗口、窗口句柄和 Windows Access Bridge 上下文。

Q: 金额重复怎么办？

A: 不按金额单索引。当前索引是 `金额 + 对手方`。复杂重复数据可以先由用户在 Excel 里预处理。

Q: Excel 打开时能运行吗？

A: 不建议。WPS/Excel 打开文件时写 C 列或自动拆分 A/B 时可能失败。程序会把底层 `PermissionError` 包装成 `ExcelLockedError`，明确提示当前写入操作和 Excel 路径。

Q: 旧坐标流程还能用吗？

A: 不能。旧坐标入口和旧 GUI 模块已经删除，后续新功能都应走 JAB。

## 质量检查

提交前统一入口：

```bash
.venv/bin/python tools/check.py
```

该入口包含 JSON 格式、配置语义、ruff、format、compileall、basedpyright、架构边界和 pytest 纯逻辑测试。

当前架构边界：

- `JABBatchProcessor` 只做装配入口和 CLI 任务级方法，不再堆业务细节。
- `core/nc_*_workflow.py` 之间不能直接互相 import，跨流程协作通过 processor 装配对象。
- `pyautogui` 只能留在 `core/jab_operator.py` 边界内。
- workflow 模块不允许新增裸 `raise RuntimeError(...)`，要使用 `core.errors` 中的领域异常。
