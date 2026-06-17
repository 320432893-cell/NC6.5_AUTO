# 收款单正式流程口径

更新时间：2026-06-15

本文是当前收款单自动化的正式口径快照。老 `README.md` 和 `TODO_JAB_HANDOFF.md` 可能混有历史探测、T0 脚本和已淘汰方案；判断正式流程时以本文和当前正式代码调用链为准。

## 1. 正式入口

底层正式业务入口：

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

现场测试只使用这个入口，一个文件内选择保存、不保存、故障恢复或 verify 审查：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/receipt_full_flow_save_query_write_test.py
```

该脚本直接回车默认选择“保存 + 后验查询 + 写 Sheet2”，默认使用当前三笔 `811,839,828`，条数默认 `3`，启动前等待默认 `2` 秒。可选功能：

1. 保存 + 后验查询 + 写 Sheet2。
2. 不保存，只跑到保存前并执行 verifier。
3. 故障恢复诊断：客户写完后暂停，人工打开干扰窗口后继续；后续动作失败时才触发 `Alt+C` 恢复并重试当前动作一次。
4. verify 审查：不保存，输出 JSON，重点看后台 verifier 和最终报告。

真实保存必须显式传 `--save` 或使用上述现场测试 wrapper；两种方式都需要交互确认 `SAVE`，除非自动化调用方明确传 `--yes-i-understand`。

禁止把 `tools/tmp_*`、`tools/archive/*` 或历史真实保存 T0 脚本当正式入口。明细单独测试入口是 `tools/receipt_detail_entry.py`，但它只测当前自制录入页明细，不负责开单、表头、保存和后验查询。

## 2. 整体流程

正式完整流程按以下顺序执行：

1. 读取 `config.json` 和 Excel Sheet1。
2. 用 `ReceiptEntryWorkbook.build_local_plan()` 做本地预检，生成 `ReceiptPlanRow`。
3. 对选中行逐笔执行 NC 录入：启动主 JAB -> `新增 -> 自制` -> 表头 -> 明细主行 -> 手续费分支 -> 可选保存。
4. 如果启用 `--query-after-save`，保存成功后按主体分组做统一后验查询。
5. 如果启用 `--write-selected-plan-sheet`，把本批结果写入 Sheet2 `收款单自动化结果`。

录入前不查 NC，不再用 Sheet1 的 `是否NC已做过` 决定候选。当前前提是交给机器的行就是本批待录入行；录入完成后再用 NC 后验查询验证保存结果。

完整流程默认从收款单录入父页开始执行 `新增 -> 自制`，不先探测“是否已经在自制录入态”。若现场已经停在自制录入页，应先人工回到父页再启动完整流程；当前自制页明细单独诊断才使用 `tools/receipt_detail_entry.py`。开单、表头、明细和保存复用同一个主 JAB，不在 `自制` 后重新起子进程或重新启动主 JAB。

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
开单点击自制 -> 从当前 canvas/缓存取得动态前缀 -> 用 `财务组织(O)` 锚点确认前缀 -> 拼接稳定后缀 path 直接写字段 -> 失败时只做弹窗故障判断并重试当前动作一次，仍失败即停止
```

> 2026-06-16 晚现场阻塞更新：上面这条“稳定后缀 path 直接写字段”的口径当前不能继续当作已验证事实。
> 最新测试停在 Sheet 行 811 的表头 `客户` resolve 阶段：`财务组织` 步骤日志显示已在当前 canvas 内用
> `财务组织(O)` 找到 `0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.2.1.0`，并执行了
> guarded clipboard paste + Enter，但 JAB snapshot 仍为空，肉眼也未确认财务组织落值；随后 `客户`
> scoped label-following-text 未找到，固定 path
> `0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.17.0`
> 也未找到。`setTextContents(A001)` 已由日志确认返回失败，不得再作为财务组织写入主路径。
> 后续接手必须先只读/小探针确认：财务组织真实进入编辑器的动作、A001 落值证据、客户字段在财务组织成功提交后的真实路径。
> 未确认前不要继续改保存、明细、查询，也不要再靠猜测后缀推进。

稳定事实：

- 当前页 path 前缀会随打开的 NC 窗口/页签数量变化。
- 同类收款单页面内，各字段深层后缀稳定。
- 不能把 `.2/.3` 这类前缀段当长期真相。

当前正式代码：

- `新增 -> 自制` 后复用开单结果里的当前页窗口和动态索引；已有缓存命中时直接复用。
- `fill_header()` 正常路径不做 `0..8` 动态 index 扫描，不用所谓 fast path 在主线程试探财务组织。
- 当前代码曾改为“当前 canvas scoped label-following-text 优先，固定后缀 path 仅作诊断 fallback”，但该方案还没有现场跑通到客户字段；不要把它写成已验证正式口径。
- 主路 path 失败时，先判断是否被可取消 Java 弹窗打断；如是，聚焦弹窗并 `Alt+C` 后重试当前失败动作。
- 不是弹窗导致的 path 失败时直接停止，不再等待或使用语义预热 path 接管。
- 字段写完后的 JAB text/description 只记录快照，不作为同步阻塞 oracle；保存前和后验查询才是业务闭包。
- 开单成功后立即进入表头写入；下方表格 path 预热只在财务组织写入成功后后台启动，不得阻塞财务组织首字段写入。

表头字段顺序：

1. 财务组织：写主体编码，例如 `A001`。
2. 客户：写客户编码，例如 `YW04278`。
3. 单据日期：写 Sheet1 到款日期。
4. 币种：写表头币种代码或账户配置币种，例如 `USD` / `CNY`。
5. 结算方式：写 `网银`。

表头收款银行账户不直接写。它由下方明细收款银行账户同步带出，所以保存前只读回校验，不作为表头输入步骤。

客户不再在表头刚写完后用 JAB 同步读回空值阻塞流程。客户是否真正落入 NC 由后续明细 verifier、保存结果和后验查询共同闭包。

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

明细表 path 定位后立即启动后台 verifier，由后台线程做 path 预热和阶段快照；完整流程不在写明细前同步读一次整表。

当前同步等待点只保留在必要闭包：

- 等最后一个字段验证结果。
- 等最终行数验证结果。

如果后台 verifier 失败，才执行整表 `read-after-fallback` 做同步读表兜底。

表头不再做逐字段语义预热或语义接管。正常路径只走当前 canvas 下的动态 path；动态 path 定位失败时只允许弹窗恢复后重试一次，仍失败就停止。表头字段写完后不再同步等待 JAB text/description 读回再判成功。

## 7. 故障恢复与保存前闭包

可取消 Java 弹窗恢复不是正常路径前置扫描。正式流程只在某个动作失败、异常、前台窗口不匹配、JAB 写入失败、剪贴板失败或键盘写入失败后，才检查是否出现 `SunAwtDialog` 且带 `取消/Alt+C` 控件的可恢复弹窗。若确认是该弹窗打断，则先聚焦弹窗，再发送 `Alt+C`，随后只重试刚才失败的当前动作一次。

当前保存前闭包由已完成的写入步骤和 verifier 结果组成：表头字段按 path 写入返回成功；明细主行、手续费行和最终行数由后台 verifier 校验；表头收款银行账户由明细收款银行账户同步带出，JAB 读回为空只记录 warning，不作为保存前硬失败 oracle。没有单独的“逐字段保存前同步总扫描”，避免把 JAB 同步读回空值误判为失败。

未知状态下禁止继续写字段或保存。`Alt+C` 只用于已识别的可取消 Java 弹窗恢复，不作为通用关闭窗口手段。

## 8. 保存

默认不保存，停在保存前。

启用 `--save` 后，正式流程在确认当前前台窗口属于收款单录入页后，用键盘热键触发 `Ctrl+S`。保存成功 oracle 是 NC 回到可新增状态，而不是单纯快捷键触发成功，也不是查找/点击凭证制单保存按钮；不要走 `SendInput(Ctrl+S)`。

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

Sheet2 当前字段（顺序以代码 `core/receipt_sheet.py` 的 `RESULT_SHEET_HEADERS` 为准）：

- 原Sheet1行号
- 执行主体名称
- 到款日期
- 🟪银行来款名
- 客户编码
- NC客户名称
- 🟪原始金额
- 手续费
- 🟪到账金额
- 币种
- 银行
- 收款银行账户
- 本地预检状态
- 后验核对状态
- 异常原因

金额列口径（与代码 `core/receipt_amounts.py` 一致）：

- `🟪原始金额` 列写 NC 原币金额合计 = Sheet1 原始金额 + 手续费（`receipt_nc_amount`，即 NC 主行+手续费行的原币合计）。
- `🟪到账金额` 列写 Sheet1 原始金额（到账净额，`receipt_net_amount`）。
- `后验核对状态` 由保存后查询回填；旧的 `NC单据号` 列已淘汰（见 `DEPRECATED_RESULT_SHEET_HEADERS`）。

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

- 表头主 path 成功时同步等待语义扫描或读回确认。
- 旧 near-label 或窗口邻近搜索兜底。
- 表头字段写完后等待 JAB text/description 同步读回再判成功。
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
.venv/bin/python -m pytest -q tests/test_receipt_full_flow_entry.py tests/test_receipt_full_flow_test_wrappers.py tests/test_receipt_query_fill.py tests/test_receipt_post_save_query.py tests/test_receipt_self_made_fill_trial.py tests/test_receipt_entry.py tests/test_validate_config.py
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
