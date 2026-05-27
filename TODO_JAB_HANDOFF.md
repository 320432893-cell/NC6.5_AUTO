# NC JAB 自动化交接

日期：2026-05-27

## 0. 结论

当前主线只维护 Java Access Bridge（JAB）方案。旧的 `pyautogui` 坐标点击、截图识别、固定坐标方案已经废弃，不要再沿着坐标方案扩功能。

开发和运行分工：

- WSL 源码仓库：`/home/queclink/project/nc_auto_v2`
- H 盘 Windows/JAB 运行镜像：`/mnt/h/python脚本/.venv/nc_auto_v2`
- JAB 运行 Python：`/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe`

JAB 必须用 Windows Python 连接 Windows 上的 NC Java UI；WSL/Linux Python 只能做开发、git、重构和普通测试。

不要再新建 `/home/queclink/projects/nc_auto_v2` 这类平行目录，当前仓库固定在 `/home/queclink/project/nc_auto_v2`。

## 1. 关键文件

- `tools/jab_batch.py`：命令入口，支持 `plan`、`generate`、`switch-generated`、`backfill`、`split-keys`。
- `core/jab_batch_processor.py`：凭证批量生成/回填业务流程，下一个 AI 优先读这里。
- `core/jab_operator.py`：JAB 底层封装，负责读表、选行、点按钮、关闭窗口、F3/F5、隐藏空白 AWT 小窗。
- `core/data_handler.py`：Excel 读取、拼接 key 解析、结果写回、拆分 key。
- `config.json`：Excel 路径、JAB DLL、列位、查询切换配置。

辅助探测工具：

- `tools/jab_probe.py`
- `tools/query_jab.bat`
- `tools/run_jab_probe.bat`
- `tools/test_voucher_selection.py`

## 2. WSL 开发和运行

开发、提交、重构在 WSL：

```bash
cd /home/queclink/project/nc_auto_v2
git status --short
```

实际跑 NC/JAB 前同步到 H 盘运行镜像：

```bash
rsync -a --delete \
  --exclude .git \
  --exclude .venv-local \
  /home/queclink/project/nc_auto_v2/ \
  /mnt/h/python脚本/.venv/nc_auto_v2/
```

在 H 盘镜像目录里调用 Windows Python：

```bash
cd /mnt/h/python脚本/.venv/nc_auto_v2
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py plan
```

常用命令：

```bash
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py plan
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py generate --yes
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py switch-generated
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py backfill
/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe tools/jab_batch.py split-keys
```

## 3. Excel 业务规则

当前 Excel：`C:\Users\Queclink\Desktop\5.27凭证.xlsx`

`Sheet1` 约定：

- A 列：`金额+对手方` 拼接索引。
- B 列：状态/凭证号。
- C 列：可由 `split-keys` 拆出的金额。
- D 列：可由 `split-keys` 拆出的对手方。

示例：

```text
141688.50深圳市科贸电子科技有限公司
```

配置：

- `sheet_my = Sheet1`
- `has_header = true`
- `jab_batch.key_col = 1`
- `jab_batch.result_col = 2`
- `jab_batch.amount_out_col = 3`
- `jab_batch.partner_out_col = 4`

解析规则：

- A 列必须以金额开头。
- 金额允许逗号和小数，统一转为 `Decimal("0.01")`。
- 对手方取金额后的文本，并去掉全部空白。
- 当前索引始终是 `金额 + 对手方`，不是单据号，也不是单金额。

B 列规则：

- `plan` / `generate` 跳过 B 列已有值的行。
- 生成成功但未回填凭证号时，B 列写 `已生成待回填`。
- `backfill` 会读取所有 key 行，把 `已生成待回填` 替换为凭证号。

用户准备数据的方式：

- Excel 行顺序来自真实附件顺序，是业务顺序的唯一权威。
- 用户会先把 NC 数据放到 Sheet2，再用辅助列拼出金额+对手方 key。
- Sheet1 按真实附件顺序整理 A 列 key。
- 重复金额、重复对手方的复杂情况可由用户预处理，少量异常允许人工处理。

## 4. NC 表格列位

待生成主表：

- 金额列：`col=4`
- 对手方列：`col=3`
- 选择列：`col=0`
- 常见规模：约 `246-252 rows x 25 cols`

制单窗口：

- 窗口标题：`制单`
- 窗口类：`SunAwtDialog`
- 制单表通常是 `N rows x 13 cols`
- 匹配方式：金额匹配，并在整行文本里查找对手方。

已生成/正式单据表：

- 金额列：`col=4`
- 对手方列：`col=3`
- 凭证日期列：`col=18`
- 凭证号列：`col=22`
- 凭证号回填时去前导 0。
- 凭证号有效范围：`1 <= 凭证号 <= 9999`。

## 5. 已实现命令

### `plan`

只读规划，不点击 NC：

1. 读取 Excel A 列 key。
2. 跳过 B 列已有状态/凭证号的行。
3. 解析金额和对手方。
4. 读取当前 NC 待生成主表。
5. 按 `金额 + 对手方` 精确匹配。
6. 输出可匹配、格式错误、未找到、重复项。

### `generate --yes`

真实点击 NC，完成生成和保存。

当前正确流程：

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

关键业务结论：

- 第一阶段必须先全量选中 Excel 匹配到的待生成行。
- 只点一次前台生成。
- 不要每条 Excel 单独生成。
- 不要拆成多轮待生成表 `选几行 -> 生成 -> 前台生成`。

原因：

- 只选 1 行时，NC 制单界面可能和多行场景不同。
- 反复前台生成会让界面状态、顺序、验证逻辑变复杂。
- 用户确认过：正确场景是先把 Excel 本轮数据抓住，再在制单窗口内按 Excel 顺序保存。

### `switch-generated`

从待生成界面切到已生成/正式单据列表：

1. 如有 `制单` 窗口，先尝试关闭。
2. 激活 NC 主窗口。
3. 用 F3 打开查询窗口。
4. 选择 `正式单据`。
5. 点击 `确定`。
6. 读取表格确认已进入已生成列表。

已生成和待生成占同一个业务窗口。不要每生成一张就来回切；正确方式是全部生成后统一切到已生成表回填。

### `backfill`

从已生成表回填凭证号到 Excel B 列。

当前假设界面已经在已生成/正式单据列表；如果不确定，先跑 `switch-generated`。

逻辑：

1. 读取 Excel 所有 key 行。
2. 读取已生成表。
3. 按 `金额 + 对手方` 匹配。
4. 历史重复时，优先取凭证日期等于当天的记录。
5. 读取凭证号列。
6. 去前导 0。
7. 校验 `1-9999`。
8. 写回 Excel B 列。

### `split-keys`

把 Excel A 列拼接 key 拆成独立金额列和对手方列。

- 默认 C 列写金额。
- 默认 D 列写对手方。
- A/B 不动。

用途是方便人工复核，后续也可以直接读取 C/D，减少反复解析 A 列字符串。

## 6. 制单保存策略

当前代码在制单窗口里：

1. 读取制单表。
2. 对每个 Excel 待保存项匹配制单行。
3. 按制单行号递增拆成可保存批次。
4. 选中当前批次。
5. 点击保存。
6. 验证当前批次消失。

两种策略：

- `single`：一行一保存，凭证号顺序最稳。
- `batch`：递增批量保存，速度更快，但凭证号可能不递增。

已发生真实案例：

```text
excel_rows=[25, 26]
voucher_rows=[1, 9]
Excel行25 -> 370
Excel行26 -> 369
```

结论：

- 如果凭证号必须严格按 Excel 顺序递增，制单窗口内应一行一保存。
- 如果优先速度，可以批量保存，但必须以已生成表真实匹配结果回填，不能用保存顺序推断凭证号。

当前用户倾向：先保证正确性，按 Excel 顺序保存；后续优化速度不能牺牲验证。

## 7. 验证逻辑

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

重要经验：

- `制单` 窗口还在但表为空是正常状态。
- 这通常表示本轮选中的制单数据已保存完。
- 正确处理是关闭窗口，刷新待生成表，再复核。
- 如果制单窗口还在但完全找不到表，要区分“空表正常”和“JAB 读表异常”。

## 8. 已验证结果

2026-05-27 已完成 33 行端到端测试。

关键结果：

- 前 3 行成功从已生成表回填历史凭证号。
- 第 5-8 行小批量测试成功。
- 第 9-10 行验证“制单窗口仍在但制单表为空”的正常完成状态。
- 第 11-34 行按正确全量逻辑测试：
  - 待生成主表一次性选中 24 条。
  - 只点一次 `生成 -> 前台生成`。
  - 制单窗口出现 24 条。
  - 制单表从 24 行递减到 0。
  - 待生成表从 246 行变成 222 行。
  - 已生成表凭证号回填成功。

已验证能力：

- JAB 可读取 NC 待生成主表。
- JAB 可多选待生成表行。
- JAB 可点击 `生成` 和 `前台生成`。
- JAB 可读取 `制单` 弹窗表格。
- JAB 可在制单表中跨行选择。
- JAB 可切到已生成/正式单据列表。
- JAB 可读取已生成表凭证号。
- JAB 可隐藏空白 AWT 小窗。

## 9. JAB 底层坑

JAB 只能在 Windows Python 下连接 NC UI。

底层配置：

- DLL：`C:\Users\Queclink\AppData\Local\UClient\share\java1.7.0_51-x64\bin\WindowsAccessBridge-64.dll`
- NC Java：`C:\Users\Queclink\AppData\Local\UClient\share\java1.7.0_51-x64`

不要乱改启动方式：

- 老版 JAB 使用 `Windows_run()`。
- `tools/jab_probe.py` 和 `core/jab_operator.py` 已处理消息泵。
- 如果随手拆线程/改初始化，可能导致 `isJavaWindow()` 全部返回 false。

不要硬编码 JAB path 或 hwnd：

- 窗口重开后 hwnd 会变。
- 控件 path 也可能变化。
- 正式逻辑应按窗口标题、role、name、table 结构动态搜索。

不要信 `bounds`：

- 很多控件返回负坐标或 `-1,-1,-1,-1`。
- 当前主线应优先使用 JAB action、selection、table API。

左上角奇怪截图/蓝框遮挡：

- 开启 JAB 后曾多次出现左上角小窗口，用户描述为“NC 的截图还去不掉”“蓝框里边空白”“没有标题”。
- 这不是 NC 业务界面，也不是需要点击的弹窗。
- 根因基本确定是 Java/AWT/JAB 辅助窗口残留。
- 常见特征：`SunAwtWindow`、title 为空、小尺寸、空白蓝框或像截图一样遮挡视线。
- `core/jab_operator.py` 已有 `hide_blank_awt_windows()`，启动和关闭 JAB 时都会尝试隐藏。
- 如果后续又出现，先检查该函数是否被调用、`config.json` 里 `jab.hide_blank_awt_windows` 是否为 `true`。

## 10. 业务坑

- 不要回到坐标方案。
- 不要按金额单索引，金额可能重复；当前索引是 `金额 + 对手方`。
- 不要依赖 NC 当前表顺序；Excel 顺序来自真实附件顺序，是业务顺序。
- 不要每保存一张就去已生成表查；全部生成后统一切已生成表回填。
- Excel/WPS 打开文件时写 B 列会报 `PermissionError`，影响状态写入、凭证号写入、`split-keys`。
- `max_batches` 不适合全量生成中途停止；当前全量生成模式不能中途留下制单窗口。
- 调试时更适合先少量 Excel 数据，而不是截断制单窗口。

## 11. 当前 TODO

### 11.1 开发/运行流程

1. 明确 WSL 仓库为源，H 盘只作为 Windows/JAB 运行镜像。
2. 增加 `tools/sync_to_windows.sh`：同步到 H 盘，排除 `.git`、`.venv-local`、缓存、日志大文件，同步后可选跑 `plan` 冒烟。
3. 更新 README：改成 JAB 主线和 WSL 开发方式；旧 `pyautogui` 只放历史兼容/已废弃段落。

### 11.2 当前功能补强

4. `backfill` 自动切换：先 `switch-generated`，再回填；切换失败要明确当前界面状态。
5. Excel 文件锁检测：写入前检测是否被 WPS/Excel 占用，避免 NC 已生成但 Excel 没写上。
6. 保存策略显式配置：新增 `save_strategy=single|batch`，默认建议 `single`。
7. 性能优化：记录每张耗时，拆成读表、选行、前台生成、保存、验证、写 Excel；减少固定 `sleep`，改等具体 UI 状态。
8. 异常恢复：中断时明确当前处于待生成表、制单窗口、已生成表、查询窗口哪个阶段。

### 11.3 数据模型和契约

9. 新增 `core/models.py`，把 dict 改成 dataclass：
   - `ExcelVoucherItem`
   - `PendingTableRow`
   - `PendingMatch`
   - `VoucherTableRow`
   - `VoucherSaveMatch`
   - `GeneratedVoucherRow`
   - `BackfillUpdate`

10. 新增 `core/contracts.py`，统一前置/后置契约：
   - 进入生成前：Excel 待处理行必须全部可解析。
   - 选行前：每个 Excel 行必须唯一匹配待生成表。
   - 前台生成后：必须出现 `制单` 窗口。
   - 保存后：目标制单行必须消失、行数减少或制单表为空。
   - 关闭制单后：本轮记录必须从待生成表消失。
   - 回填前：当前表必须是已生成/正式单据表。

11. 统一错误类型，替代散落的裸 `RuntimeError`：
   - `WorkflowStateError`
   - `TableMatchError`
   - `ContractViolation`
   - `ExcelLockedError`
   - `JABControlNotFound`

### 11.4 状态机

12. 加状态机，但不要过度设计。建议状态：
   `excel_loaded -> pending_snapshot_read -> pending_rows_selected -> front_generate_clicked -> voucher_window_opened -> voucher_saving -> voucher_empty -> pending_refreshed -> generated_opened -> backfilled`

13. 状态机用途：
   - 限制非法跳转。
   - 失败时告诉用户卡在哪一步。
   - 支持人工接管后继续。
   - 不要把 GUI 操作改成全异步事件驱动；NC UI 更适合同步状态机。

14. 状态持久化：写 `logs/run_state.json`，记录本轮 Excel 行、NC 行、已保存行、已回填行，辅助异常恢复。

### 11.5 管道过滤

15. 数据侧用管道过滤，GUI 点击不要塞进管道。
16. 生成 pipeline：
   `LoadExcelRows -> ParseConcatKey -> DropFilledRows -> ReadPendingSnapshot -> MatchAmountPartner -> ValidateUniqueMatches -> BuildSavePlan`
17. 回填 pipeline：
   `LoadExcelRows -> ReadGeneratedSnapshot -> MatchGeneratedRows -> PreferToday -> NormalizeVoucherNo -> BuildBackfillUpdates`

### 11.6 命令模式

18. 抽 NC 操作为命令对象：
   - `ReadPendingTableCommand`
   - `SelectPendingRowsCommand`
   - `OpenFrontGenerateCommand`
   - `ReadVoucherTableCommand`
   - `SelectVoucherRowsCommand`
   - `SaveVoucherRowsCommand`
   - `CloseVoucherWindowCommand`
   - `RefreshPendingTableCommand`
   - `SwitchGeneratedCommand`
   - `BackfillExcelCommand`

19. 命令模式目标：统一日志、重试、dry-run、审计每一步输入输出。

### 11.7 总线、事件驱动、事件总线

20. 新增轻量同步事件总线 `core/events.py`，不要引入消息队列。
21. 事件示例：
   - `ExcelRowsLoaded`
   - `PendingRowsMatched`
   - `PendingRowsSelected`
   - `FrontGenerateOpened`
   - `VoucherRowsSaved`
   - `VoucherWindowEmpty`
   - `PendingVerificationPassed`
   - `GeneratedRowsMatched`
   - `ExcelBackfilled`

22. 事件总线用途：日志、审计、进度显示、后续截图/报警/人工确认扩展点。
23. 不要做全事件驱动重写；主流程保持同步 workflow，事件总线只做观测和扩展。

### 11.8 审计和复核

24. 增加 run 日志/审计表：每次运行生成 run id，记录 Excel 行、金额、对手方、待生成 NC 行、制单行、凭证号、状态。
25. 增加回填后复核报表：输出成功、未找到、重复、凭证号异常；凭证号倒序只提示，不直接判失败。
26. 最终凭证号以已生成表真实结果为准。

## 12. 最近提交

```text
c3a5624 Document JAB handoff pitfalls
2f72704 Condense JAB handoff TODO
55b2354 Document handoff notes and pitfalls
b74f18d Add Excel key splitting command
a9f72df Add NC JAB voucher automation workflow
```
