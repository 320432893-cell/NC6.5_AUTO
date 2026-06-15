# 收款单正式流程口径

更新时间：2026-06-15

本文是当前收款单自动化的正式口径快照。老 `README.md` 和 `TODO_JAB_HANDOFF.md` 可能混有历史探测、T0 脚本和已淘汰方案；判断正式流程时以本文和当前正式代码调用链为准。

## 1. 正式入口

主入口：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_full_flow_entry.py --excel-rows 811,839,828 --limit 3 --save --query-after-save --write-selected-plan-sheet
```

Windows 镜像路径：

```text
H:\python脚本\.venv\nc_auto_v2
```

默认不保存：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_full_flow_entry.py --excel-row 1791 --limit 1
```

真实保存必须显式传 `--save`，并且需要交互确认 `SAVE`，或由 `tools/nc_auto_test_menu.bat` 菜单二次确认后加 `--yes-i-understand`。

禁止把 `tools/tmp_*`、`tools/archive/*` 或历史真实保存 T0 脚本当正式入口。明细单独测试入口是 `tools/receipt_detail_entry.py` 或 `tools/receipt_detail_test_menu.bat`，但它们只测当前自制录入页明细，不负责开单、表头、保存和后验查询。

## 2. 整体流程

正式完整流程按以下顺序执行：

1. 读取 `config.json` 和 Excel Sheet1。
2. 用 `ReceiptEntryWorkbook.build_local_plan()` 做本地预检，生成 `ReceiptPlanRow`。
3. 对选中行逐笔执行 NC 录入：`新增 -> 自制 -> 表头 -> 明细主行 -> 手续费分支 -> 保存前守卫 -> 可选保存`。
4. 如果启用 `--query-after-save`，保存成功后按主体分组做统一后验查询。
5. 如果启用 `--write-selected-plan-sheet`，把本批结果写入 Sheet2 `收款单自动化结果`。

录入前不查 NC，不再用 Sheet1 的 `是否NC已做过` 决定候选。当前前提是交给机器的行就是本批待录入行；录入完成后再用 NC 后验查询验证保存结果。

## 3. 本地预检

本地预检从 `receipt_entry.excel.start_row` 开始读取 Sheet1。

必需业务列包括：

- `到款日期`
- `🟪银行来款名`
- `🟪原始金额`
- `银行`
- `币种`
- `客户编码`
- `手续费`

银行主体识别顺序固定为：

```text
Sheet1 银行 -> receipt_entry.accounts -> organization_code -> finance_organizations
```

本地预检异常必须保留结构化信息：原 Sheet1 行号、阶段、异常类型、字段、原值、配置节点、说明和处理动作。Sheet2 只展示压缩后的 `异常原因`。

重复键：

```text
主体 + 到款日期 + 银行 + 币种 + 客户编码 + 银行来款名 + 金额
```

同一重复键出现多行时，整组标记为 `DUPLICATE_EXCEL_ROWS`，不进入录入。

`validation_policy.mode`：

- `strict`：任意异常停止整批。
- `skip_invalid_rows`：异常行写 Sheet2，跳过异常行继续处理可运行行。

## 4. 表头录入策略

表头正式策略是：

```text
语义定位当前收款单页面 -> 推导动态前缀 -> 拼接稳定后缀 path 写字段 -> 失败时语义接管
```

稳定事实：

- 当前页 path 前缀会随打开的 NC 窗口/页签数量变化。
- 同类收款单页面内，各字段深层后缀稳定。
- 不能把 `.2/.3` 这类前缀段当长期真相。

当前正式代码：

- `locate_receipt_header_scope()` 先用稳定后缀扫描当前前台收款单表头 scope。
- 快速 path 定位失败时，用语义定位 `财务组织` 推断表头动态前缀。
- `fill_header()` 启动 `HeaderSemanticPreload` 并发预热客户、日期、币种、结算方式等语义路径。
- 每个字段主路仍走动态前缀 + 固定后缀 path。
- 主路 path 失败时，等待语义预热结果接管。

表头字段顺序：

1. 财务组织：写主体编码，例如 `A001`。
2. 客户：写客户编码，例如 `YW04278`。
3. 单据日期：写 Sheet1 到款日期。
4. 币种：写表头币种代码或账户配置币种，例如 `USD` / `CNY`。
5. 结算方式：写 `网银`。

表头收款银行账户不直接写。它由下方明细收款银行账户同步带出，所以保存前只读回校验，不作为表头输入步骤。

客户是保存前硬门槛。客户为空时必须停止，不能软通过、不能写明细、不能保存。

## 5. 明细表录入策略

明细表正式模块：

- `tools/receipt_body_table_locator.py` 定位 25 列收款单明细表。
- `tools/receipt_detail_fields.py` 定义字段和读回校验。
- `tools/receipt_detail_screen_writer.py` 执行受保护前台键盘写入。
- `tools/receipt_detail_writer.py` 编排主行字段写入。
- `tools/receipt_detail_rows.py` 编排手续费行。
- `tools/receipt_detail_row_cleanup.py` 清账户和删除多余行。
- `tools/receipt_detail_async_verifier.py` 后台并发验证。

明细输入原则：

- 用 JAB selection 选中单元格。
- 发送全局键盘前必须确认当前前台窗口属于目标 NC 窗口。
- 近距离列移动用方向键，不用反复双击或直接聚焦目标单元格。
- 字段值用剪贴板整段粘贴，然后 `Enter` 确认。
- 收款银行账户列必须从相邻单元格方向键进入，粘贴裸账号后等待约 `0.1s` 再 `Enter`，让 NC 匹配候选并确认落格。

主行字段顺序：

1. 收款业务类型：`货款`
2. 收款银行账户：账户裸账号
3. 科目：`1002`
4. 贷方原币金额：Sheet1 原始金额
5. 结算方式：`网银`

币种只写表头。主行和手续费行不写明细币种。

手续费为 0：

- 不新增手续费行。
- 如果主行后自动带出多余空行，只允许在确认不会删除已写主行的条件下删除多余行。
- 最终行数期望为 1。

手续费大于 0：

- 主行写完后保留或新增第 2 行给手续费。
- 第 2 行无论原内容是什么，都按手续费行覆盖。
- 手续费行字段：业务类型 `手续费`、科目 `660305`、手续费金额、结算方式 `网银`。
- 手续费行收款银行账户必须为空；如果自动带出账号，用 `Delete` 清空。
- 如果冒出第 3 行空行，只删除第 3 行。
- 最终行数期望为 2。

## 6. 并发验证策略

正式完整流程已接入 `DetailPipelineVerifier`。

写明细时，每个字段提交后会把验证任务交给后台 verifier；主线程继续写后续字段，不逐字段同步等待。

当前同步等待点只保留在必要闭包：

- 等最后一个字段验证结果。
- 等最终行数验证结果。

如果后台 verifier 失败，才执行整表 `read-after-fallback` 做同步读表兜底。

表头语义预热也是并发的。正常路径不等待全扫描；只有动态 path 定位失败时才等待语义预热接管。

## 7. 保存前守卫

保存前必须确认：

- 财务组织正确
- 客户非空
- 单据日期正确
- 币种正确
- 表头结算方式为 `网银`
- 表头收款银行账户已由明细同步带出
- 明细主行正确
- 手续费行正确或已跳过
- 多余空行已处理

如果保存前守卫识别不到收款单自制录入态，且当前是误打开的阻塞式查询/参照/异常窗口，正式逻辑允许用 `Alt+C` 关闭可取消弹窗，然后重新检查页面状态。

`Alt+C` 只用于保存前恢复。未知状态下禁止继续写字段或保存。

## 8. 保存

默认不保存，停在保存前。

启用 `--save` 后，正式流程发送保存动作。保存成功 oracle 是 NC 回到可新增状态，而不是单纯快捷键发送成功。

真实保存前必须确保：

- 用户明确授权。
- 当前 Excel 行号和测试范围正确。
- 当前 NC 账号权限覆盖本批主体。
- 测试单据是否保留或人工清理已明确。

## 9. 保存后查询

启用 `--query-after-save` 后，正式流程只查询保存成功的本批行。

查询分组：

- 按主体分组。
- 同一主体内按到款日期升序、原 Sheet1 行号升序。
- 最多四个主体。

每个主体查询流程：

1. 回到收款单查询父页。
2. 用 `pyautogui.press("f3")` 打开 `查询条件`；未检测到弹窗时每 `0.2s` 重按一次，直到 Win32 看到可见 `title=查询条件`、`class=SunAwtDialog`。
3. 定位查询条件窗口自己的动态前缀。
4. 写入收款财务组织、单据日期起、单据日期止。
5. 确认查询。
6. 查询结果 ready 信号优先看结果表 path 是否出现；出来后先把每页条数改为 500。
7. 如果无需翻页，只读当前页匹配。
8. 如果需要翻页，先读第一页；若本批目标已全部匹配，不翻页；未匹配完再翻页继续。

查询条件 path 策略和表头一致：

```text
查询窗口动态前缀 + 稳定后缀 path 为主；失败后当前查询窗口语义路径推断接管
```

查询条件不再使用旧 near-label 写入兜底。

查询结果页分页和结果表定位也走 path 优先：

- 优先尝试模块动态前缀 `0.0.1.0.0.0.0.{index}`，拼接结果区固定后缀 `0.0.0.1.1.0.0.0.1.1.1.0.0.0`。
- 结果表稳定后缀是 `0.0.0`；若模块动态前缀失败，再枚举当前查询结果页可见 table path 反推结果区域动态前缀。
- 拼接分页控件稳定后缀：页码标签 `1.6`、每页行数输入框 `1.7`、下一页按钮 `1.2`。
- 同一前缀下三个分页控件都能验证通过时，确认该结果区域 path。
- “每页 500”、读取结果表、下一页都使用这组动态 path/scope。
- 同一轮查询内缓存 `result_table_path`、结果区前缀、分页 hwnd 和三个分页控件 path；缓存命中且 `trust_cached_paths=true` 时直接按 `cached_trusted` 使用，不重复验证 label/table/page size/next。
- 如果每页行数已经是 `500`，不再写入和 Enter；只按配置读取必要的分页 label/每页行数。
- 结果表读取只读匹配需要的配置列，例如单据号、单据日期、客户/来款名、原币金额，不全表扫列。
- 结果表列数不是契约，用户勾选列会导致列数变化；列数只用于抽取/匹配阶段判断配置列是否存在，不能作为分页或表格定位门槛。

如果配置列索引超过当前结果表实际列数，正式逻辑应报“结果表列不足/抽取失败”，而不是退回旧的固定 41 列定位逻辑。

2026-06-15 现场复核：

- 收款单自制录入态表头区域动态前缀示例：`0.0.1.0.0.0.0.5`。
- F3 查询条件窗口使用自己的前缀体系，示例：`0.0.1.0.1.0.0.1.0.0.0.0.0.1.0.1`。
- 查询结果页仍回到同一收款单模块下的结果区域，结果区前缀示例：`0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0.0.0`。
- 当次表头动态索引和结果页动态索引同为 `.5`，但查询条件窗口是弹窗内独立前缀；正式逻辑不能拿查询条件前缀推结果页，只能在各自区域用 path 锚点推当前区域前缀。
- 正式查询结果页优先用模块动态索引 `0.0.1.0.0.0.0.{index}` 拼结果区固定后缀；失败再枚举 table path 反推结果区。
- 分页控件在 JAB 中可能是 `visible/editable` 但没有 `showing` 和有效 bounds；分页 path 校验、读页码、读/写每页行数不以 `showing` 作为硬门槛。
- 每页行数写入后用 Enter 触发刷新；如果当前 Windows Python 没有 `pyautogui`，正式逻辑用 SendInput 发送 Enter 兜底。
- F3 打开查询条件不能只用裸 SendInput F3；现场验证裸 SendInput 只让窗口变化但没有稳定弹出查询条件。正式入口以 `pyautogui.press("f3")` 重试为主。
- 最新样本读取耗时基线：`ensure_query_window=0.527s`、`query.dynamic-scope=0.275s`、`result_wait_before_read=3.327s`、`read_receipt_result_pages=1.134s`；其中 `setup_seconds=0.546s`、`read_tables_seconds=0.295s`、`pager_resolution=cached_trusted`、`dynamic_resolution=dynamic_module_index_table_ready`。`result_wait_before_read` 仍是后续可优化项。

## 10. 后验匹配口径

后验匹配只匹配本批保存成功的目标行。

匹配字段：

- 原始金额必须精确相等。
- 日期优先匹配本批到款日期；如果同金额候选没有同日，再保留金额候选做异常诊断。
- 名称使用录入客户编码后 NC 主表显示出的客户名称，不用 Sheet1 银行来款名。
- 名称归一化相似度必须不低于 90。

结果分类：

- 唯一命中：写 NC 单据号。
- 多条命中：异常原因写重复匹配。
- 金额相同但名称不达标：异常原因写名称相似度不足，并带期望名和候选名。
- 名称匹配但金额不一致：异常原因写 NC 候选金额。
- 金额无对应：异常原因写金额无对应。
- 查询过程异常：异常原因只写查询阶段的问题。

## 11. Sheet2 写入

正式结果页是 Sheet2，默认表名：

```text
收款单自动化结果
```

正式主线不写 Sheet1 的 `是否NC已做过`。

Sheet2 当前字段：

- 原Sheet1行号
- 执行主体名称
- 到款日期
- 🟪银行来款名
- 客户编码
- NC客户名称
- 🟪到账金额
- 🟪原始金额
- 手续费
- 币种
- 银行
- 收款银行账户
- 本地预检状态
- NC单据号
- 异常原因

写入规则：

- 复用已有 Sheet2。
- 缺少表头则补齐。
- 删除旧噪音列。
- 清空表头下方旧行。
- 重写本次计划或本批结果，不做历史追加。

排序规则：

```text
主体 -> 到款日期 -> 原 Sheet1 行号
```

同一主体的数据应集中在一起。

## 12. 已淘汰逻辑

以下逻辑不得回到正式主线：

- 表头收款银行账户参照搜索。
- `fallback_reference` 配置。
- 查询条件旧 near-label 写入兜底。
- 录入前按 NC 查询结果跳过 Sheet1 候选。
- Sheet1 `是否NC已做过` 作为正式批量结果写回。
- 用 `tools/tmp_*` 作为正式入口。
- 无前台窗口守卫的全局键盘输入。
- 用坐标点击、截图找坐标或 bounds 中心点点击替代 JAB/表格语义。

## 13. 当前验证命令

Linux 侧检查：

```bash
.venv/bin/python -m pytest -q tests/test_receipt_full_flow_entry.py tests/test_receipt_query_fill.py tests/test_receipt_post_save_query.py tests/test_receipt_self_made_fill_trial.py tests/test_receipt_entry.py tests/test_validate_config.py
.venv/bin/python tools/check.py changed
python3 tools/validate_config.py config.json
git diff --check
```

Windows 镜像同步：

```bash
tools/sync_to_windows.sh
```

Windows 侧轻检查：

```bat
set PYTHONIOENCODING=utf-8&& cd /d H:\python脚本\.venv\nc_auto_v2&& py -3.11 -m compileall core tools tests\test_receipt_full_flow_entry.py tests\test_receipt_query_fill.py tests\test_receipt_entry.py
set PYTHONIOENCODING=utf-8&& cd /d H:\python脚本\.venv\nc_auto_v2&& py -3.11 tools\validate_config.py config.json
```

## 14. 下一步现场验证

下一步应使用正式入口做三笔真实保存验证：

- 一笔有手续费。
- 一笔无手续费且人民币。
- 一笔香港移为。

验证目标：

- 连续保存成功。
- 保存后按主体统一查询。
- Sheet2 写入业务列、NC客户名称、NC单据号和异常原因。
- 记录总耗时，以及每笔录入、保存、查询耗时。

真实 NC 操作必须等用户明确授权后执行。
