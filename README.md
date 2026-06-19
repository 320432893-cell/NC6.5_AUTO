# NC JAB 凭证自动化

这个项目用于在 NC6.5 中按 Excel 顺序批量生成凭证，并把已生成凭证号回填到 Excel。

当前主线是 Java Access Bridge（JAB）。旧的 `pyautogui` 坐标点击、截图识别、固定坐标方案已经删除，不再作为新功能方向。

本仓库当前不是桌面 GUI 应用，也没有正式图形界面入口。对 NC 的真实操作只通过 CLI 工具驱动 Windows Python + JAB；若后续要做可视化配置或前端，也必须复用现有配置对象、计划模型和 JAB 工作流，不能另起坐标/截图方案。

## GUI / 前端接手口径

当前没有正式 GUI，也不要把 `tools/tmp_*`、`.bat` 菜单或现场探测脚本理解成 GUI。它们只是 CLI/测试入口。未来如果要做桌面 GUI、Web 前端或 TypeScript 配置台，必须接现有业务对象，而不是重写自动化路径。

可视化入口应该围绕这些对象设计：

- 配置编辑：`config.json` 的 `receipt_entry.schema_version=2`，尤其是 `finance_organizations`、`banks`、`accounts`、账户候选值、录入策略和明细/手续费策略。
- 本地预检：`core/receipt_entry.py`、`core/receipt_plan.py`、`core/receipt_plan_issue.py`、`core/receipt_sheet.py` 和 `tools/receipt_entry_check.py`。
- 完整录入测试：`tools/receipt_full_flow_entry.py` 是正式业务入口；现场测试只直接用 `tools/receipt_full_flow_save_query_write_test.py`，在同一个文件内选择保存、不保存、故障恢复或 verify 审查。
- 查询与后验：`tools/receipt_query_fill.py` 及拆分出的 `tools/receipt_query_*` 模块。
- JAB 底层边界：`core/jab_operator.py`、`core/jab_window.py`、`core/jab_*` mixin/helper；全局键盘、前台窗口和 pyautogui 只能留在这里。
- 工程检查：`tools/check.py`、`tools/validate_config.py`、`tools/check_architecture.py` 和对应测试。

前端页面优先级：

1. 配置对象编辑器：银行、账号、主体、别名、币种、录入策略。保存前必须跑 `tools/validate_config.py`。
2. Excel 本地预检结果页：展示 `ReceiptPlanRow` 和 `ReceiptPlanIssue`，能按主体、银行、异常类型过滤。
3. 执行计划确认页：只展示将要处理的行、主体、金额、手续费和风险提示；真实保存前仍要用户明确授权。
4. 运行监控页：只展示 CLI/JAB 工作流事件、当前阶段、耗时和可恢复状态，不直接绕过后置校验。
5. 后验查询和 Sheet2 结果页：按主体汇总查询结果、匹配状态和写回结果。

禁止方向：

- 不做坐标配置器、截图识别配置器或固定 path 编辑器。
- 不把 `pyautogui.write`、剪贴板、全局快捷键暴露成普通按钮。
- 不让前端直接拼临时脚本执行真实保存；必须调用正式入口并保留确认、日志和检查闭包。
- 不把旧 `是否NC已做过` 写回工具重新包装成新批量录入的前置筛选主线。

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
- `core/receipt_sheet.py`：Sheet2 表头维护、旧噪音列删除和当前计划结果区重写。
- `core/errors.py`：NC workflow 领域异常，区分页面状态、表格匹配、JAB 控件、JAB 动作、Excel 写入锁和流程契约失败。
- `core/jab_operator.py`：JAB 底层封装，负责读表、选行、按钮动作、F3/F5、关闭窗口；AWT 残留清理只保留显式入口，不随 JAB 启停自动执行。
- `core/data_handler.py`：Excel 读取、拼接 key 解析、结果写回、拆分 key。
- `config.json`：Excel 路径、JAB DLL、列位和查询切换配置。
- `tools/check.py`：统一检查入口，默认包含基础检查和 audit 门禁。
- `.semgrep.yml` / `.importlinter` / `tools/check_detect_secrets.py`：`tools/check.py` audit profile 使用的安全、架构和密钥检查配置。
- `TODO_JAB_HANDOFF.md`：后续开发 TODO。
- `CHANGELOG.md`：已实现功能、验证记录和历史变更。

辅助探测工具：

- `tools/jab_probe.py`
- `tools/receipt_body_table_locator.py`
- `tools/receipt_full_flow_entry.py`：收款单完整流程正式业务入口，消费 `ReceiptPlanRow`，默认开单/表头/明细/手续费后停在保存前，显式 `--save` 才真实保存。
- `tools/receipt_full_flow_save_query_write_test.py`：现场测试入口，一个文件内选择完整流程保存、不保存、故障恢复诊断或 verify 审查；默认保存、后验查询并写 Sheet2 本批结果。
- `tools/receipt_detail_entry.py`：收款单明细主行/手续费行正式 Python 入口；供脚本化测试明细写入能力，不保存、不暂存。
- `tools/receipt_entry_check.py`
- `tools/receipt_query_fill.py`
- `tools/receipt_self_made_flow.py`
- `tools/close_awt_popup_residue.py`
- `tools/query_jab.bat`
- `tools/run_jab_probe.bat`

现场探测、复盘或窄场景诊断用的临时脚本已清理，不再保留为入口；明细主行/手续费行能力已沉淀到正式入口 `tools/receipt_detail_entry.py`。旧表头账户参照探针已移入 `tools/archive/`，已删除的真实保存 T0 脚本只能从 git 历史恢复。

历史探针和人工现场测试入口已收口到 `tools/archive/`。这些文件只作复盘证据，不是测试人员默认入口，也不能被正式流程 import；需要重新启用时先移回合适目录并补检查闭包。

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

收款单完整流程正式业务入口是 `tools/receipt_full_flow_entry.py`。它从 `ReceiptEntryWorkbook.build_local_plan()` 取通过本地预检的 `ReceiptPlanRow`，默认只跑一行，执行 `新增 -> 自制 -> 表头 -> 明细主行 -> 手续费分支` 后停在保存前，不发送保存：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_full_flow_entry.py --excel-row 1791 --limit 1
```

运行前先把本地预检结果写入 Sheet2：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_full_flow_entry.py --excel-row 1791 --limit 1 --write-plan-sheet
```

现场测试只保留入口 `tools/receipt_full_flow_save_query_write_test.py`。直接运行 `.py` 后先选择测试功能，再按提示输入行号/条数/等待秒数；直接回车默认功能是“保存 + 后验查询 + 写 Sheet2”，默认用当前三笔 `811,839,828`、条数 `3`、启动前等待 `2` 秒。功能 1 会自动追加 `--save --query-after-save --write-selected-plan-sheet`，仍需要输入 `SAVE` 确认；功能 2 不保存，只跑到保存前并执行 verifier；功能 3 在客户写完后暂停，人工打开干扰窗口后继续，后续动作失败时才触发现有 `Alt+C` 故障恢复并重试当前动作一次；功能 4 不保存并输出 JSON，重点看后台 verifier 和最终报告。保存动作是确认前台属于收款单录入页后，用键盘热键触发 `Ctrl+S`；保存 oracle 是保存后回到收款单录入父页并检测到可用【新增】。后验查询按本批保存成功行的主体分组，最多四个主体；每个主体用本批日期区间查一次 NC。查询入口用 `pyautogui.press("f3")` 每 `0.2s` 重试直到看到可见 `查询条件/SunAwtDialog`，不是裸 SendInput F3。结果页按动态模块前缀 + 固定后缀定位结果表和分页控件，缓存命中后以 `cached_trusted` 复用；先确保每页 500，再按页只读匹配需要的配置列并只匹配本批目标。匹配口径是原始金额精确相等、日期优先、NC 客户显示名归一化相似度不低于 90；结果只写 Sheet2，不写 Sheet1 的 `是否NC已做过`：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_full_flow_save_query_write_test.py
```

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

收款单本地预检主入口。该命令从 `receipt_entry.excel.start_row` 开始读 Sheet1，只做 Excel/配置侧识别和异常拦截，不查 NC；CLI 异常会保留行号、字段、原值、配置节点和处理动作，Sheet2 只写人看的单列 `异常原因`：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_entry_check.py
```

写入 Sheet2 结果表，默认表名为 `收款单自动化结果`。当前正式写入规则是：有表头就复用，缺表头就补齐，旧噪音列会按表头名删除，然后清空表头下方旧行并重写本次计划结果；不再写录入/保存/查询调试明细：

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

收款单查询窗口只填条件、不点确定；默认会先在收款单录入页按 F3 打开查询条件窗口。正式入口用 `pyautogui.press("f3")` 重试，直到 Win32 检测到可见 `title=查询条件`、`class=SunAwtDialog`；现场不要改回只发裸 SendInput F3：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-05-01 --date-to 2026-06-02
```

收款单查询后读取可见结果表：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-05-01 --date-to 2026-06-02 --confirm --read-results
```

收款单查询后只读匹配预演；查询后会把每页条数改为 500，并按分页读取。结果页 ready 优先看结果表 path，路径策略是模块动态前缀 `0.0.1.0.0.0.0.{index}` + 结果区固定后缀，再拼结果表后缀 `0.0.0`、页码 label `1.6`、每页行数 `1.7`、下一页 `1.2`。同一轮查询会缓存结果表 path、分页 hwnd 和分页控件 path；如果每页已经是 500，不再写入和 Enter。输出 JSON 包含 `page_report`、金额范围、名称样本和重复原因；日期只用于查询范围，不参与匹配诊断。注意：这是历史/诊断工具，不是当前新主线的录入前筛选步骤：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-03-31 --date-to 2026-05-31 --confirm --dry-run-match --max-rows 600 --max-cols 140
```

收款单匹配结果写回 Excel；唯一匹配写 `已做过`，完全未命中写 `未做过`。金额命中但名称不符、名称命中但金额不符会写明 Excel 值和 NC 候选值；重复命中按实际条数写 `重复N条：名称和金额相同，需人工确认`，重复行也会在 JSON 的 `duplicate_rows` 中报告。`金额和对手方均未匹配` 是诊断原因口径，不是当前最终写回状态。注意：该写回旧入口不再作为新批量录入主线的前置判断，当前用户口径是假定交给机器的行均未做过，录入完成后再按主体查询 NC 做后验验证：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-03-31 --date-to 2026-05-31 --confirm --dry-run-match --write-back --max-rows 600 --max-cols 140
```

重跑并覆盖 Excel 已有 `是否NC已做过` 状态时，加 `--include-filled-status`：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_query_fill.py --org-code A001 --date-from 2026-03-31 --date-to 2026-05-31 --confirm --dry-run-match --write-back --include-filled-status --max-rows 600 --max-cols 140
```

写回前必须先确认当前 NC 页面是目标 `收款单录入`。工具会做页面和结果表守卫；如果结果表的匹配名称列疑似读成 `收款单` 单据类型列，会拒绝继续。

`收款财务组织` 当前按字段 label 聚焦右侧输入框后直接写组织编码，例如 `A001`、`A006`。不要点输入框右侧未知按钮；不要为了提交该输入框额外按 Enter，除非现场已确认该主体必须这样做。写入是否有效以后置查询结果和 Excel 主体匹配数为准，不能只看 JAB 文本读回。

收款单明细测试直接调用 `tools/receipt_detail_entry.py`。它不保存、不暂存，只定位当前自制录入页的 25 列明细表，并调用正式 `tools/receipt_detail_*` 模块。运行前必须把 NC 停在【收款单自制录入】界面，且没有参照窗口或提示框遮挡：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_detail_entry.py
```

只测试手续费分支，要求主行已写好：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_detail_entry.py --fee-only --fee-amount 10
```

只清理第 1 行以外的多余明细行：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_detail_entry.py --cleanup-extra-rows-only
```

## 入口安全矩阵

| 入口 | 读 Excel | 写 Excel | 点击/读取 NC | 保存业务单据 | 当前定位 |
| --- | --- | --- | --- | --- | --- |
| `tools/jab_batch.py plan` | 是 | 否 | 只读待生成表 | 否 | 凭证批量只读规划 |
| `tools/jab_batch.py generate --yes` | 是 | 是 | 是 | 是，保存凭证 | 凭证正式生成入口 |
| `tools/jab_batch.py switch-generated` | 否 | 否 | 是 | 否 | 凭证页签切换 |
| `tools/jab_batch.py backfill` | 是 | 是 | 读取已生成表 | 否 | 凭证号回填 |
| `tools/receipt_entry_check.py` | 是 | 否 | 否 | 否 | 收款单本地预检 |
| `tools/receipt_entry_check.py --write` | 是 | 是，重写 Sheet2 当前计划结果区 | 否 | 否 | 收款单计划结果写入 |
| `tools/receipt_full_flow_entry.py` | 是 | 可选写 Sheet2 预检；保存后可写本批 Sheet2 结果 | 是 | 默认否，`--save` 会保存 | 收款单完整流程测试入口 |
| `tools/receipt_full_flow_save_query_write_test.py` | 是 | 默认保存后写 Sheet2 本批结果 | 是 | 可选保存 | 现场测试：保存/不保存/故障恢复/verify 审查 |
| `tools/receipt_query_fill.py --confirm --read-results` | 是 | 否 | 是 | 否 | 收款单查询/抽取组件 |
| `tools/receipt_query_fill.py --dry-run-match --write-back` | 是 | 是，写 Sheet1 状态列 | 是 | 否 | 历史查重/诊断入口 |
| `tools/receipt_self_made_flow.py` | 是 | 否 | 是 | 默认否 | 单行/分阶段现场试填 |
| `tools/receipt_detail_entry.py` | 否 | 否 | 是，写当前明细表 | 否 | 明细正式测试入口 |
| `tools/tmp_*` | 视脚本而定 | 视脚本而定 | 视脚本而定 | 禁止当正式入口 | 探测/复盘参考 |

收款单完整流程正式入口已经可以消费 `ReceiptPlanRow` 跑到保存前，也可以显式真实保存。现场测试只用固定 wrapper：`tools/receipt_full_flow_save_query_write_test.py`，在同一个文件内选择保存、不保存、故障恢复或 verify 审查；保存功能会按主体统一查询 NC 并写 Sheet2。本主线不写 Sheet1 状态列；`tools/receipt_query_fill.py --dry-run-match --write-back` 只保留为历史查重/诊断入口。

## NC 现场操作纪律

自动化和用户共用同一个桌面输入焦点。任何 `pyautogui.write`、`pyautogui.hotkey`、剪贴板粘贴、普通按键都可能打进用户聊天框或 VS Code。执行这类全局键盘输入前，必须先单独确认目标窗口是前台窗口；如果 `GetForegroundWindow()` 不是目标 hwnd，脚本必须停，不允许发送快捷键、账号或 Enter。

默认录入路径必须优先走后台 JAB：`setTextContents`、JAB action、selection/table API 和 JAB 控件树读回。前台窗口只用于判断能不能发送全局键盘输入，不能用来否定已经由后台 JAB 写入并回显成功的业务状态。明细表是当前例外：正式代码已改为先用 JAB 选中单元格并校验前台窗口，再用方向键短距移动到目标列、剪贴板整段粘贴、`Enter` 确认、后台表格读回校验。

表头文本框默认只做后台写入并继续，不强求立刻变中文或解锁下一字段。`财务组织(O)` 写 `A001` 后，不等待后台 description/text 变成主体中文名；先写 `客户`，再写 `单据日期`、`币种`、`结算方式`。写完后的 JAB text/description 只作为快照记录，不能因为读回为空就否定已经返回成功的写入。

客户不再在表头刚写完后用 JAB 读回空值阻塞流程。2026-06-16 起，客户是否真正落入 NC 由后续明细 verifier、保存结果和后验查询共同闭包；肉眼可见已成功但 JAB 同步读回为空时，不能判定表头失败。

NC 表头层级很深，JAB 默认搜索深度必须保持 `max_depth=50`。`25` 层只能看到 `客户收款单` 页签和浅层财务组织，搜不到 `客户`、`单据日期`、`币种`、`收款银行账户` 等真实表头 label，会造成误判。

收款单页签路径里的 `.2/.3` 这类段位会随当前打开的 NC 界面数量变化；用户打开另一个界面后当前收款页可能从 `.2` 变成 `.3`。这不是业务状态，也不是失败原因，不能把固定页签索引当长期依据。正式逻辑用稳定后缀扫描当前表头 scope，必要时用语义定位当前页并推导本次动态前缀，再拼接字段 path。

表头字段正式定位口径是“当前 canvas + `财务组织(O)` 锚点确认动态前缀 + 稳定后缀 path”。主 path 失败时只允许先判断是否被可取消 Java 弹窗打断，命中则聚焦弹窗并 `Alt+C` 后重试当前动作一次；仍失败就停止。不得回退旧 near-label、表头账户参照、逐字段语义接管或坐标方案。

不要把“前台化、输入、等待结果”合成一个长命令。先做短前台检查，再执行用户已确认的动作。遇到遗留 `使用权参照`、NC 主窗口在屏幕外、`新增` 匹配到非主窗口、或任何状态污染时，只报告状态，不继续补救性点击或写字段。

收款单自制录入当前已完成授权主体真实保存 T0：`新增 -> 自制`、表头、明细主行、手续费行、手续费账户清空、多余空行删除、保存前表头回看和 `Ctrl+S` 后回到【新增】均已验证。授权范围以当前账号权限为准，已确认可用主体是 `A001/移为` 和 `A006/移为香港`。没有明确允许时仍禁止保存、暂存。

> 2026-06-16 晚当前代码阻塞：不要把上面的历史 T0 当作当前表头代码已跑通。最新完整流程停在 Sheet 行 811 的 `客户 - resolve`。
> `财务组织` 步骤只证明代码定位到 `财务组织(O)` scoped text 并发送 guarded paste + Enter，未证明 NC 已落入 `A001`；用户肉眼确认财务组织未写入。
> `setTextContents(A001)` 已确认返回失败。客户字段随后 scoped label-following-text 和固定 path 均定位失败。接手前必须先做单字段探针确认财务组织真实输入/提交方式和客户字段真实路径。

`新增 -> 自制` 的成功标准是上方 `保存(Ctrl+S)`、`暂存`、`取消(Ctrl+Q)` 三个按钮同时出现；只看到 JAB action 返回成功或菜单项被点击，不算进入自制录入态。25 列明细表只作为填明细前的单独校验，不作为开单成功条件。完整流程复用同一个主 JAB 完成开单、表头、明细和保存；`自制` 后不得再起子进程或重新启动主 JAB 造成空等。

`tools/receipt_self_made_flow.py` 默认只负责开单和表头阶段；表头字段主路是“当前 canvas + 动态前缀 + 固定后缀 path”写入，`财务组织(O)` 是前缀硬锚点。完整流程开单成功后立即写表头首字段；下方表格 path 预热只在财务组织写入成功后后台启动，不得阻塞财务组织写入。明细填入默认禁用，必须显式传 `--fill-detail` 才会尝试进入明细；明细正式入口统一调用 `tools/receipt_detail_*`，使用受保护前台键盘和剪贴板粘贴，不能回退到坐标或无守卫 typing。

可取消 Java 弹窗恢复不是正常路径前置扫描。完整流程只在某个动作失败、异常、前台窗口不匹配、JAB 写入失败、剪贴板失败或键盘写入失败后，才检查是否出现 `SunAwtDialog` 且带 `取消/Alt+C` 控件的可恢复弹窗；命中后先聚焦弹窗，再发送 `Alt+C`，随后只重试刚才失败的当前动作一次。

收款单自制录入和授权主体真实保存 T0 已达标；当前代码已沉淀出 `tools/receipt_full_flow_entry.py`，用于消费 `ReceiptPlanRow` 跑完整流程测试。默认模式跑到保存前停止；真实保存必须显式 `--save` 并确认。保存后按主体统一后验查询由 `--query-after-save` 触发，Sheet2 本批结果由 `--write-selected-plan-sheet` 写入。真实保存 T0 脚本已清理，仓库仍保留若干 tmp 探测脚本；后续正式入口不要继续引用临时脚本。真实保存运行前仍必须由用户明确授权，成功 oracle 是保存后【新增】重新出现。测试保存单据可能被用户手工删除，后续查询不能依赖历史 T0 单据仍存在。

明细表当前已验证的填写顺序：主行写 `收款业务类型=货款`、`收款银行账户`、`科目=1002`、`贷方原币金额`、`结算方式=网银`。币种只写表头，输入 `USD`/`CNY` 后回车；主行和手续费行不写币种。正式写入方式是：首字段用 JAB selection 选格，后续字段沿当前选中格用方向键移动；字段值整段剪贴板粘贴后 `Enter` 确认。账号列必须从相邻列方向键进入，整段粘贴裸账号后等待 `0.1s` 再 `Enter`，已用 row `1424` 账号 `FTE310066674143603000377` 验证落格。表头 `结算方式` 必须显式写 `网银`，不能依赖主体默认值；明细主行和手续费行的 `结算方式` 输入后用 Enter 确认，避免联想浮层残留。手续费非零时才 `Ctrl+I` 增行，手续费行写 `手续费`、`660305`、手续费金额、`网银`；手续费行账户必须为空，自动带出时用 `Delete` 清空；若多出空白第 3 行，用 `Ctrl+D` 删除。

表头字段提交和校验口径：财务组织编码、客户编码、单据日期、币种、表头结算方式按稳定 path 写入后提交，现场验证口径是输入后 Enter 再进入下一字段。即时 JAB `setTextContents` 返回成功或同控件读回为空都不是业务 oracle；当前保存前闭包依赖表头写入返回、明细后台 verifier、保存结果和后验查询，不做逐字段同步总扫描。收款银行账户由明细账号带出并由后台 verifier 校验；表头账户 JAB 读回为空只记 warning，不作为保存前硬失败。

状态污染恢复：只有当前动作失败或异常后，才检查是否为可取消 Java 弹窗；确认后聚焦弹窗并 `Alt+C`，然后重试当前失败动作一次。未知状态下禁止继续写字段或保存。

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

- `receipt_entry.schema_version=2` 是面向后续配置维护/前端扩展的业务对象模型；任何 UI 都应编辑这些对象，不要把银行、账号、快捷键继续写死到脚本里。
- `receipt_entry.excel.start_row` 是 Sheet1 业务起始行；新主线从该行开始读取，不再用“最近 2 个月”或 `是否NC已做过` 状态列决定本批候选。
- `receipt_entry.excel.result_sheet_name` 是 Sheet2 名称，默认 `收款单自动化结果`。当前正式函数 `rewrite_plan_sheet()` / `rewrite_batch_result_sheet()` 会复用/补齐表头、删除旧噪音列，然后清空表头下方旧行并重写本次计划或本批结果；不要把它理解成历史追加表。
- `receipt_entry.excel.currency_column`、`customer_code_column`、`fee_column` 分别指定币种、客户编码、手续费列，用于本地预检和后续录入。
- `receipt_entry.validation_policy.mode` 支持 `strict` 和 `skip_invalid_rows`。`strict` 有任意异常就停止整批；`skip_invalid_rows` 会把异常行/重复组写 Sheet2 并跳过。
- `receipt_entry.banks` 是银行字典，只做展示、分类和可选别名维护；同一家银行可能有多个主体账户，银行别名不能自动当账户匹配规则。
- `receipt_entry.accounts` 是账户字典，每个账户必须有稳定 `id`、`enabled`、`organization_code`、`bank_id`、`account_label`、`account_no`。Excel 银行列匹配只看账户自己的 `account_label`、`aliases`、`excel_bank_aliases`。
- `nc_candidates_by_currency` 用于配置 NC 可输入候选。当前收款银行账号统一输入裸账号，不再加 `RMB/USD/CNY` 后缀；只有账号 `97460154740002297` 是 `CNY`，其余已配置账号按 `USD`。
- `entry_policy` 记录账户录入策略：当前主线是 `account_input=detail_first`、`success_rule=non_empty`；`fallback_reference` 已弃用，配置校验会拒绝。
- `detail_entry_policy` 记录明细主行/手续费行顺序和快捷键：主行不包含币种，手续费增行 `ctrl+i`，手续费账户清空，额外空行删除 `ctrl+d`。
- 当前业务口径是币种只写表头，输入 `USD`/`CNY` 后回车；主行和手续费行不写明细币种。
- `tools/validate_config.py` 会校验组织、银行、账户引用、账户别名冲突、候选值和策略枚举；新增银行或账号后必须先跑配置校验。

收款单本地预检异常口径：

- 本地预检由 `ReceiptEntryWorkbook.build_local_plan()` 执行。输出模型是 `ReceiptPlanRow` 和 `ReceiptPlanIssue`。
- `ReceiptPlanIssue` 必须带 `excel_row`、`stage`、`issue_type`、`field`、`raw_value`、`config_node`、`message`、`action`；不要新增笼统的“配置错误/数据错误”。Sheet2 展示时会压缩为单列 `异常原因`，详细结构只留给 CLI/日志和代码判断。
- 当前会识别缺必需列、起始行无效、银行为空/未配置/账户禁用、账户主体不存在、日期错误、银行来款名为空、金额错误或非正、币种为空或不支持、客户编码为空、手续费错误或负数、以及本批 Sheet1 内重复。
- 重复键为 `主体 + 到款日期 + 银行 + 币种 + 客户编码 + 银行来款名 + 金额`。同一 key 出现多行时标 `DUPLICATE_EXCEL_ROWS`，整组不录入。
- 通过预检的行按主体分组供后续录入使用；录入前不查 NC。全部录入后再进入后验查询阶段，查询结果只影响 Sheet2 的 `NC单据号` 和 `异常原因`。`--query-after-save` 已接入正式后验查询：按本批保存成功行的主体分组，最多四个主体；每个主体用本批日期区间查一次 NC，先改每页 500，再按页读取并只匹配本批目标。匹配使用录入客户编码后 NC 主表显示出的客户名，不使用 Sheet1 银行来款名；名称归一化相似度阈值不低于 90，金额必须精确相等。
- 收款单录入 path 策略的稳定事实：当前页路径前缀会随打开的 NC 页面/页签动态变化，字段后缀在同类收款单页面内稳定。正式入口已按“当前 canvas + `财务组织(O)` 锚点确认动态前缀 + 稳定后缀”为主接入；表头主 path 失败不再语义接管，只在动作失败后尝试 `Alt+C` 阻塞弹窗恢复并重试当前动作一次。明细后台 verifier 已接入完整流程。
- 收款单查询条件也按同一口径处理：F3 打开查询条件后先解析查询窗口自己的动态前缀，拼接 `收款财务组织`、`单据日期起止` 的稳定后缀写入；动态 path 失败后走当前查询窗口内的语义路径推断，不再回退旧 near-label 写入。F3 触发用 `pyautogui.press("f3")` 按 `0.2s` 重试直到 `查询条件/SunAwtDialog` 可见，裸 SendInput F3 现场验证不稳定。
- 收款单查询结果页正式策略：优先用模块动态前缀 `0.0.1.0.0.0.0.{index}` 拼结果区固定后缀 `0.0.0.1.1.0.0.0.1.1.1.0.0.0`，再拼结果表 `0.0.0`、页码 label `1.6`、每页行数 `1.7`、下一页 `1.2`；失败才枚举 table path 反推。缓存命中后按 `cached_trusted` 复用，不重复校验所有分页控件。读取结果只取匹配必要列，不全表扫列。

收款单自制录入明细表：

- 当前用 `tools/receipt_body_table_locator.py` 按 25 列 body table 特征定位，空白录入态通常是 1 行 25 列，旁边另有独立的合计表。
- JAB selection API 的 child index 规则是 `row * 25 + col`。
- 已探测关键列位按 0 基记录：`col=1` 收款业务类型，`col=3` 币种，`col=4` 收款银行账户，`col=5` 科目，`col=7` 金额/贷方原币金额，`col=11` 结算方式。当前业务口径不写明细币种列，币种只在表头写 `USD`/`CNY` 后回车。
- 选中表格单元格不等于键盘焦点进入编辑器。正式明细写入必须先确认当前前台窗口属于本次定位到的 NC 表格，再发送方向键、剪贴板粘贴和 Enter；禁止无前台守卫盲填。
- 当前正式明细策略：首字段 JAB 选格，后续字段按当前列用 Left/Right 方向键移动；所有非清空字段使用整段剪贴板粘贴，普通字段粘贴后短等待再 Enter，账号列粘贴后等待 `0.1s` 再 Enter 让 NC 完成候选匹配。`setTextContents`、逐字符 typing、直接双击/聚焦粘贴账号列都不是当前主路径。
- `结算方式` 上方也有输入框，且 A006/移为香港默认可能为空；表头区必须先显式写 `网银`。明细表 `col=11` 仍要写 `网银` 并用 Enter 确认，下方写入后的同步只作为验证，不能替代表头显式录入。
- `收款银行账户` 当前主路径是明细表 `col=4` 方向键进入后粘贴裸账号并 Enter，不再走表头账户参照；表头账户只做一次快照记录。明细账号和最终行数由后台 verifier 校验，表头账户 JAB 读回为空只记 warning，不作为保存前硬失败。

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
- `目的业务日期` 依赖先选择 `正式单据` 才出现；选择前可见的 `.2` 行是 `生效日期`，不要用于 `generated_date_value`。当前目的业务日期起止 path 是 `.11.1.0.0` / `.11.1.2.0`。
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

默认 profile 是 `all`，`changed` 是同一组本地门禁别名。二者包含：

- 基础检查：JSON 格式、配置语义、ruff、format、compileall、basedpyright、架构边界和 pytest 纯逻辑测试。
- audit 检查：semgrep、import-linter、detect-secrets、pip-audit。

可用 profile：

```bash
.venv/bin/python tools/check.py all
.venv/bin/python tools/check.py changed
.venv/bin/python tools/check.py audit
.venv/bin/python tools/check.py deep
.venv/bin/python tools/check.py rule-tool-contracts
.venv/bin/python tools/check.py --list
```

`audit` 依赖 `.semgrep.yml`、`.importlinter`、`tools/check_detect_secrets.py` 和开发依赖。仅改 Markdown 时，通常先跑 `git diff --check -- README.md TODO_JAB_HANDOFF.md CHANGELOG.md` 即可；涉及 Python、配置或依赖时再跑对应 profile。

当前架构边界：

- `JABBatchProcessor` 只做装配入口和 CLI 任务级方法，不再堆业务细节。
- `core/nc_*_workflow.py` 之间不能直接互相 import，跨流程协作通过 processor 装配对象。
- `pyautogui` 只能留在 `core/jab_operator.py` 边界内。
- workflow 模块不允许新增裸 `raise RuntimeError(...)`，要使用 `core.errors` 中的领域异常。
