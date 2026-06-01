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
- `core/jab_operator.py`：JAB 底层封装，负责读表、选行、按钮动作、F3/F5、关闭窗口、隐藏空白 AWT 小窗。
- `core/data_handler.py`：Excel 读取、拼接 key 解析、结果写回、拆分 key。
- `config.json`：Excel 路径、JAB DLL、列位和查询切换配置。
- `TODO_JAB_HANDOFF.md`：后续开发 TODO。
- `CHANGELOG.md`：已实现功能、验证记录和历史变更。

辅助探测工具：

- `tools/jab_probe.py`
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

拆分 Excel A 列拼接 key：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py split-keys
```

## Excel 规则

当前 Excel：

```text
C:\Users\Queclink\Desktop\5.27凭证.xlsx
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

`backfill` 从已生成表回填凭证号到 Excel C 列。当前假设界面已经在已生成/正式单据列表；不确定时先跑 `switch-generated`。

逻辑：

1. 读取 Excel 中 C 列等于 `已生成待回填` 的 key 行。
2. 读取已生成表。
3. 按 `金额 + 对手方` 匹配。
4. 历史重复时，优先取凭证日期等于当天的记录。
5. 读取凭证号列。
6. 去前导 0。
7. 校验 `1-9999`。
8. 写回 Excel C 列。

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

已发生真实案例：

```text
excel_rows=[25, 26]
voucher_rows=[1, 9]
Excel行25 -> 370
Excel行26 -> 369
```

45-54 后续实测又出现递减制单行批次反号，说明行号递增/递减不是唯一规律。后续反序选择也未解决顺序问题，因此批量保存不再作为顺序固定场景的正式策略。

2026-05-28 用 5.28 凭证做过历史实验：

- `batch_reverse_select`：Excel 行 2/3/4 回填为 561/560/559。
- `batch`：Excel 行 5/6/7 回填为 565/564/563。

两组都显示选择顺序不能保证 Excel 顺序对应凭证号递增，且没有继续维护价值，相关策略已删除。顺序固定优先时，正式策略使用 `single`。

`Ctrl+S` 保存触发也做过对照：每张激活和首张激活一次都能保存成功，但只是把“触发保存”的耗时转移到“等待 NC 删行验证”，端到端没有优于 JAB 按钮。因此默认仍使用 JAB 按钮。

批量保存保留为快速备选：`safe_batch_by_pending_row` 基于待生成行号和制单窗口行序做保守合并，不满足递增条件时自动拆成单张。2026-05-28 小样本显示它整体线性，但局部仍可能出现反序，因此不作为严格顺序主线。

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
- 优先使用 JAB action、selection、table API。
- 隐藏或非 visible 的 `SunAwtDialog` 查询窗口可能是残留，不应作为可操作窗口依据。

左上角奇怪截图/蓝框遮挡：

- 开启 JAB 后曾多次出现左上角小窗口，表现为“NC 的截图还去不掉”“蓝框里边空白”“没有标题”。
- 这不是 NC 业务界面，也不是需要点击的弹窗。
- 根因基本确定是 Java/AWT/JAB 辅助窗口残留。
- 常见特征：`SunAwtWindow`、title 为空、小尺寸、空白蓝框或像截图一样遮挡视线。
- `core/jab_operator.py` 已有 `hide_blank_awt_windows()`，启动和关闭 JAB 时都会尝试隐藏。

## 常见问题

Q: 现在开发是否可以完全转到 WSL？

A: 可以。源码、git、文档、重构都放 WSL。只有实际控制 NC/JAB 的 Python 进程必须是 Windows Python。

Q: 为什么不能用 WSL Python 直接跑 JAB？

A: WSL Python 看不到 Windows 桌面上的 Java 窗口、窗口句柄和 Windows Access Bridge 上下文。

Q: 金额重复怎么办？

A: 不按金额单索引。当前索引是 `金额 + 对手方`。复杂重复数据可以先由用户在 Excel 里预处理。

Q: Excel 打开时能运行吗？

A: 不建议。WPS/Excel 打开文件时写 C 列或自动拆分 A/B 时可能报 `PermissionError`，影响状态写入、凭证号写入和 `split-keys`。

Q: 旧坐标流程还能用吗？

A: 不能。旧坐标入口和旧 GUI 模块已经删除，后续新功能都应走 JAB。

## 静态检查

提交前目标检查：

```bash
.venv/bin/python -m json.tool config.json
.venv/bin/ruff check .
.venv/bin/ruff format --check .
.venv/bin/python -m compileall -q core tools
.venv/bin/basedpyright .
```
