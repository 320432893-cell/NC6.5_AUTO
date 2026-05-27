# Java Access Bridge 交接 TODO

日期：2026-05-27

## 2026-05-27 最新进展

当前 JAB 方案已经从“能探测”推进到“可跑批量生成闭环的大部分步骤”。核心原则已经确定：

- Excel `Sheet1` 行顺序是生成顺序的唯一权威。
- `Sheet1` 只需要两列：
  - A 列：`金额+对手方` 拼接 key，例如 `141688.50深圳市科贸电子科技有限公司`
  - B 列：回填凭证号/异常状态
- NC 匹配索引始终是：`金额 + 业务关联方信息1/对手方名称`。
- 先在 NC 待生成主表一次性选中 Excel 中存在的所有行，再点 `生成 -> 前台生成`。
- 进入 `制单` 窗口后，在制单表中按 Excel 行顺序查找对应记录并保存。
- 保存验证不能只依赖“保存成功”提示，优先看制单表行数减少、目标行消失，制单窗口关闭时回到待生成表刷新复核。
- 最后自动切到已生成/正式单据列表，根据已生成表回填凭证号。

### 已验证通过

- 待生成主表 JAB 可读：
  - 常见规模约 `249-252 rows x 25 cols`
  - 金额列 `col=4`
  - 对手方列 `col=3`
  - 选择列 `col=0`
- 待生成主表可通过 JAB selection API 多选。
- `生成 -> 前台生成` 可通过 JAB action 执行。
- `制单` 窗口可读：
  - `SunAwtDialog title='制单'`
  - 制单表通常是 `N rows x 13 cols`
  - 可根据金额和对手方在整行文本中匹配目标行。
- 制单表可选中指定行，也验证过跨行多选：
  - 例如选择 `[0, 2]` 时 selected indexes 为 `0-12` 和 `26-38`。
- 已验证自动切换：
  - 从待生成界面按 F3 打开查询
  - 点击 `正式单据`
  - 点击 `确定`
  - 成功进入已生成列表
  - 已生成列表读取到 `12272 rows x 23 cols`
- 已生成列表列位已验证：
  - 金额列 `col=4`
  - 对手方列 `col=3`
  - 凭证号列 `col=22`
  - 凭证日期列 `col=18`
- 回填测试已通过：
  - Excel 三条测试数据匹配已生成表
  - 原始 NC 凭证号为 `00000339`, `00000340`, `00000341`
  - 程序已改为回填有效数字：`339`, `340`, `341`
  - 凭证号校验范围：`1 <= 凭证号 <= 9999`
- 已生成表存在历史重复时，程序已改为优先取凭证日期等于当天的记录。
  - 例如 `2286 + 上海...` 同时命中历史记录和当天记录时，正确选择当天 `2026-05-27` 的 `00000341`。
- JAB 空白蓝框/小浮窗已处理：
  - 识别条件：`SunAwtWindow`、无标题、小尺寸
  - `core/jab_operator.py` 中已加入隐藏逻辑。

### 当前命令

只规划/匹配，不点击生成：

```powershell
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py plan
```

真实生成并保存：

```powershell
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py generate --yes
```

限制只跑一批：

```powershell
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py generate --yes --max-batches 1
```

只测试切换到已生成/正式单据：

```powershell
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py switch-generated
```

在已生成列表回填凭证号：

```powershell
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py backfill
```

注意：PowerShell 换行续写要用反引号。不要写成：

```powershell
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py generate
  --max-batches 1
```

这种写法会把第二行当成新命令，导致 `一元运算符 "--" 后面缺少表达式`。

### 当前文件改动

- `core/data_handler.py`
  - 新增读取 `Sheet1` 的 `金额+对手方` 拼接 key。
  - 新增 `load_jab_batch_data()`。
  - 新增 `save_jab_results()`。
- `core/jab_operator.py`
  - 封装 JAB 表格读取、控件点击、窗口关闭、按键、表格多选。
  - 新增隐藏空白 AWT 浮窗逻辑。
  - `read_table_snapshot()` 支持读取额外列。
- `core/jab_batch_processor.py`
  - 新增批量生成流程。
  - 新增制单表匹配和保存验证。
  - 新增 `switch_to_generated_list()`。
  - 新增 `backfill_generated_vouchers()`。
  - 回填时凭证号去前导 0，并校验最大 `9999`。
  - 重复记录时优先按凭证日期列取当天记录。
- `tools/jab_batch.py`
  - 命令：
    - `plan`
    - `generate`
    - `switch-generated`
    - `backfill`
- `tools/test_voucher_selection.py`
  - 用于制单窗口行选择、跨行多选测试。
- `config.json`
  - 当前 Excel：
    `C:\Users\Queclink\Desktop\5.27凭证.xlsx`
  - `generated_voucher_col=22`
  - `generated_date_col=18`
  - `generated_voucher_max=9999`

### 仍需继续测试/完善

1. 完整闭环复测：
   - 待生成界面
   - `plan`
   - `generate --yes --max-batches 1`
   - 制单表保存验证
   - 自动切到已生成
   - `backfill`
   - 检查 Excel 第二列

2. 保存验证继续加固：
   - 如果制单窗口存在但制单表消失，判定异常。
   - 如果制单窗口关闭，回待生成表 F5 刷新，确认已保存记录不再存在。
   - 如果制单表存在且为空，并且本批数量等于制单表原数量，应判定正常。

3. 回填前自动切换：
   - 当前 `backfill` 假设已经在已生成/正式单据列表。
   - 后续可以加一个选项，让 `backfill` 自动先执行 `switch-generated`。

4. Excel 被占用处理：
   - 如果 Excel 打开，保存会报：
     `PermissionError: [Errno 13] Permission denied`
   - 后续可在写入前检测文件锁，给出更明确提示。

5. 长批量稳定性：
   - 已验证 3 条样本。
   - 还需要测试几十条以上，观察 JAB 读表速度、窗口关闭/刷新时序、NC 是否有锁或延迟。

## 架构演进建议

当前批量凭证生成已经验证可行，但代码仍偏“脚本编排”。后续如果还要加更多 NC 功能，建议把项目逐步拆成可复用的自动化框架。

### 已验证的业务闭环

2026-05-27 已完成一次 33 行 Excel 的端到端测试：

- 前 3 行回填凭证号。
- 第 5-8 行先小批量生成并回填。
- 第 9-10 行验证“制单窗口在但表空白”逻辑。
- 第 11-34 行按正确逻辑执行：
  - 待生成主表一次性选中 24 条。
  - 只点一次 `生成 -> 前台生成`。
  - 制单窗口内按 Excel 顺序保存。
  - 制单表行数从 24 递减到 0。
  - 待生成表从 246 行变成 222 行。
  - 已生成表回填凭证号到 Excel。
- Excel 第 25、26 行出现凭证号顺序 `370, 369`，这是按已生成表真实匹配结果回填，建议人工复核一次。

### 推荐优先级

1. 状态机：优先加入。
   - 当前流程天然是状态驱动：
     `待生成表 -> 已全量选中 -> 制单窗口 -> 保存中 -> 制单表空 -> 待生成复核 -> 已生成回填`
   - 状态机能明确每一步的允许转移、失败恢复和人工接管点。
   - 建议新增 `core/workflow_state.py` 或 `core/state_machine.py`。

2. 契约模式：优先加入。
   - 每个关键步骤都应该有前置/后置契约：
     - 进入生成前：当前表必须是待生成表，Excel 待处理行必须全部唯一匹配。
     - 点击前台生成后：必须出现制单窗口，并且制单表行数应大于 0。
     - 保存后：本批记录必须从制单表消失，或制单表为空。
     - 关闭制单后：本轮记录必须从待生成表消失。
     - 回填前：当前表必须是已生成表，凭证列可读。
   - 建议新增 `core/contracts.py`，统一抛出带上下文的异常。

3. 命令模式：适合加入。
   - 把具体 NC 操作封装成命令对象：
     - `SelectPendingRowsCommand`
     - `OpenFrontGenerateCommand`
     - `SelectVoucherRowsCommand`
     - `SaveVoucherBatchCommand`
     - `RefreshPendingCommand`
     - `SwitchGeneratedCommand`
     - `BackfillVouchersCommand`
   - 好处是日志、重试、dry-run、审计更统一。

4. 管道过滤：适合用于数据侧。
   - Excel 和 NC 表格匹配可以改成管道：
     - `LoadExcelRows`
     - `ParseConcatKey`
     - `DropFilledRows`
     - `ReadPendingSnapshot`
     - `MatchAmountPartner`
     - `ValidateUniqueMatches`
     - `BuildVoucherSavePlan`
     - `ReadGeneratedSnapshot`
     - `ResolveVoucherNumber`
   - 管道更适合纯数据转换，不建议把 GUI 点击也塞进管道。

5. 事件总线：可以加，但应轻量。
   - 当前可以先做进程内同步事件总线，不需要消息队列。
   - 事件示例：
     - `ExcelRowsLoaded`
     - `PendingRowsMatched`
     - `PendingRowsSelected`
     - `FrontGenerateOpened`
     - `VoucherBatchSaved`
     - `PendingVerificationPassed`
     - `GeneratedBackfilled`
   - 用途：
     - 日志
     - 审计记录
     - UI/进度显示
     - 后续报警/截图/人工确认

6. 总线/事件驱动：暂不建议重做成全事件驱动。
   - NC 自动化强依赖同步 UI 状态，完全事件驱动会让控制流变难排查。
   - 建议保留同步 workflow，事件总线只做观测和扩展点。

### 建议目录结构

```text
core/
  app.py                  # 应用入口/依赖组装
  workflow_state.py       # 状态机
  contracts.py            # 前置/后置契约
  events.py               # 事件定义和轻量事件总线
  commands.py             # NC 操作命令
  pipeline.py             # 数据管道
  models.py               # ExcelRow, PendingMatch, VoucherMatch 等 dataclass
  jab_operator.py         # JAB 低层能力
  repositories.py         # Excel/NC snapshot 读写
  workflows/
    voucher_generation.py # 当前凭证生成流程
```

### 迁移策略

- 不要一次性大重构。
- 第一步先补 `models.py`，把 dict 改成 dataclass，减少 key 写错。
- 第二步抽 `contracts.py`，把现在散落的 RuntimeError 改成可分类错误。
- 第三步抽状态机，只包住当前凭证生成流程。
- 第四步再加事件总线，把日志和进度从 workflow 中剥离。
- 第五步才考虑命令模式和数据管道拆分。

### WSL/路径建议

- 可以把项目主体放到 WSL 文件系统里，例如：
  `/home/queclink/projects/nc_auto_v2`
- 但 JAB 必须用 Windows Python 和 Windows UI，不能直接在 Linux Python 下跑。
- 推荐做法：
  - 代码仓库放 WSL，便于 git、测试、重构。
  - 运行 JAB 的命令仍调用 Windows Python：
    `/mnt/h/python脚本/.venv/nc_auto_v2/.venv-local/Scripts/python.exe`
  - 或保留当前 H 盘路径作为运行目录，WSL 只做开发和 git 管理。

## 当前结论

NC6.5 这个客户端虽然是老 Java/UClient，但 Java Access Bridge（JAB）已经验证可用。它不是完全只能靠坐标点。

已经确认：

- NC 使用的 Java 是：
  `C:\Users\Queclink\AppData\Local\UClient\share\java1.7.0_51-x64\bin\javaw.exe`
- NC Java 进程已经加载：
  - `JavaAccessBridge-64.dll`
  - `JAWTAccessBridge-64.dll`
  - `awt.dll`
- Windows 侧 JAB DLL 是：
  `C:\Users\Queclink\AppData\Local\UClient\share\java1.7.0_51-x64\bin\WindowsAccessBridge-64.dll`
- JAB 配置已经开启：
  `C:\Users\Queclink\AppData\Local\UClient\share\java1.7.0_51-x64\lib\accessibility.properties`

配置内容应为：

```properties
assistive_technologies=com.sun.java.accessibility.AccessBridge
screen_magnifier_present=true
```

用户级配置也存在：

```text
C:\Users\Queclink\.accessibility.properties
```

内容同样已启用。

## 已新增工具

### `tools/jab_probe.py`

JAB 探测脚本，支持：

- 加载老版 JAB 的 `Windows_run`
- 在同一线程运行消息泵
- 枚举顶层窗口、子窗口、隐藏窗口
- 输出控件树
- `--query` 过滤控件
- `--inspect-path` 检查某个 JAB path
- 调用 table API 读取表格行列与单元格
- `--actions-path` 查看控件动作
- `--do-action-path` 执行控件动作

注意：老版 JAB 的关键修复点是 `Windows_run()` 必须和消息泵在同一线程里运行。之前主线程调用 `Windows_run()`、另起线程 pump 消息时，`isJavaWindow()` 全部返回 false。

### `tools/run_jab_probe.bat`

一键运行 JAB 探测，并保存输出到：

```text
tools\jab_probe_output.txt
```

### `tools/query_jab.bat`

按关键词搜索 JAB 控件，避免中文路径和 cmd 引号问题。

用法示例：

```bat
H:\python脚本\.venv\nc_auto_v2\tools\query_jab.bat 生成
H:\python脚本\.venv\nc_auto_v2\tools\query_jab.bat "role='push button'"
H:\python脚本\.venv\nc_auto_v2\tools\query_jab.bat "role='table'"
```

### `tools/stop_nc_uclient.bat`

只清理 NC/UClient 相关进程：

```bat
wmic process where "name='javaw.exe' and commandline like '%%UClient%%'" call terminate
taskkill /f /im Uclient.exe
```

如果用户说关闭 NC 后账号还在登录，不一定是进程没关，可能只是 UClient 本地登录态/自动登录缓存。用这个脚本确认进程是否真的清干净。

## 重大 JAB 发现

### 1. 正式业务表格可以读

在正式“单据生成”业务界面，金额区域确实嵌套在 Swing 表格里。

结构大致是：

```text
scroll pane
  viewport
    table
```

关键 scroll pane：

```text
role='scroll pane'
bounds=34,203,1246,428
children=8
```

主数据 viewport：

```text
role='viewport'
bounds=95,232,1168,359
children=1
```

真正业务表格：

```text
role='table'
bounds=95,-758,2740,5412
children=6150
```

JAB table API 可读：

```text
TABLE rows=246 cols=25
```

金额列是 `col=4`，示例：

```text
cell[0,4] name='31,080.00'
cell[1,4] name='1,182,970.00'
cell[2,4] name='141,688.50'
cell[3,4] name='4,760.00'
cell[4,4] name='9,444.00'
```

这说明可以用 JAB 直接读 NC 表格金额，后续可以减少对坐标和 Excel `Sheet2` 的依赖。

### 2. 生成按钮可以找到

在当前业务界面，JAB 能找到“生成”按钮。

关键词搜索命令：

```bat
H:\python脚本\.venv\nc_auto_v2\tools\query_jab.bat 生成
```

关键输出：

```text
role='push button'
name='生成'
desc='生成'
states='enabled,visible,showing'
```

业务主窗口里的路径：

```text
hwnd=527570
path=0.0.1.0.0.0.0.2.0.0.0.0.0.0.0.6
```

同一工具栏还能看到：

```text
删除
查询
刷新
选择
生成
拉式取数
重新生成
联查
选项
```

注意：这个按钮的 `bounds` 有时是负坐标，例如：

```text
bounds=-31744,-31862,67,30
```

所以不要直接用 JAB 返回坐标点击。已经验证可以用 `doAccessibleActions` 直接触发按钮动作。

### 3. “生成 -> 前台生成”已可执行

测试时间：2026-05-26 19:07 CST。

先对“生成”按钮执行 JAB action：

```bat
cmd.exe /c "set PYTHONIOENCODING=utf-8&& py -3.11 H:\python脚本\.venv\nc_auto_v2\tools\jab_probe.py --all --children --do-action-path 0.0.1.0.0.0.0.2.0.0.0.0.0.0.0.6 --startup-wait 5 --dll C:\Users\Queclink\AppData\Local\UClient\share\java1.7.0_51-x64\bin\WindowsAccessBridge-64.dll"
```

关键结果：

```text
role='push button' name='生成' desc='生成'
ACTIONS count=1
DO_ACTION ok=True failure=-1
```

随后 JAB 可以看到弹出的菜单项：

```text
hwnd=1314006 class='SunAwtWindow'
path=0.0.1.0.0.1
role='menu item'
name='前台生成'
states='enabled,visible,showing'
bounds=2,32,111,23
```

再对“前台生成”菜单项执行 JAB action：

```bat
cmd.exe /c "set PYTHONIOENCODING=utf-8&& py -3.11 H:\python脚本\.venv\nc_auto_v2\tools\jab_probe.py --all --children --do-action-path 0.0.1.0.0.1 --startup-wait 2 --dll C:\Users\Queclink\AppData\Local\UClient\share\java1.7.0_51-x64\bin\WindowsAccessBridge-64.dll"
```

关键结果：

```text
role='menu item' name='前台生成'
ACTIONS count=1 names=['单击']
DO_ACTION action='单击' ok=True failure=-1
```

执行后进入了新的 Java 对话框：

```text
hwnd=7541120
class='SunAwtDialog'
title='制单'
```

在 `制单` 对话框中已看到：

```text
role='push button' name='常用凭证' desc='常用凭证'
role='push button' name='凭证' desc='凭证'
```

这说明 JAB 方案已经走通到“生成 -> 前台生成 -> 制单对话框”，不是只能读取控件。

## 当前坐标方案状态

当前主流程仍在用坐标：

- `first_amount_cell`
- `find_next_btn`
- `find_close_btn`
- `generate_btn`
- `front_generate`
- `voucher_num_box`
- `return_btn`

但现在可以考虑替换：

- `generate_btn`：优先用 JAB 找“生成”按钮并执行动作。
- `front_generate`：优先用 JAB 找“前台生成”菜单项并执行动作。
- NC 表格金额查找：优先用 JAB table API 遍历 `rows x col=4`。

暂时不建议一次性全替换，因为：

- 单元格 bounds 经常返回 `-1,-1,-1,-1`
- 表格能读，但还没验证是否能通过 JAB selection API 选择指定行
- 凭证号框还没完整探测

## 下一步建议

### TODO 1：把 JAB action 封装进主流程

目标：把已经验证通过的 JAB action 接到 `core/gui_operator.py` 里的 `do_generate()`。

已在 `tools/jab_probe.py` 实现并验证：

- `getAccessibleActions`
- `doAccessibleActions`

已经验证的路径：

```text
生成按钮:
0.0.1.0.0.0.0.2.0.0.0.0.0.0.0.6

前台生成菜单项:
0.0.1.0.0.1
```

主流程 `do_generate()` 可以从：

```python
self.click_pos("generate_btn", "生成按钮")
time.sleep(...)
self.click_pos("front_generate", "前台生成")
```

替换为：

```python
jab.click_button_by_name("生成")
jab.click_menu_item_by_name("前台生成")
```

或类似封装。不要长期硬编码 path，正式实现应动态搜索 role/name/desc，因为 hwnd/path 会随窗口重开变化。

### TODO 2：实现 JAB table 读取封装

目标：封装读取当前 NC 表格：

- 找到主业务 table
- 读取 `rowCount`
- 读取 `columnCount`
- 读取指定单元格文本
- 遍历 `col=4` 查找金额

金额格式要和现有 `format_amount()` 对齐：

- Excel 里 `31080` 对应 NC 表格 `'31,080.00'`
- 需要去逗号、处理小数。

### TODO 3：验证是否能选中指定行

优先级：

1. 试 JAB selection API：
   - `addAccessibleSelectionFromContext`
   - `clearAccessibleSelectionFromContext`
   - `selectAllAccessibleSelectionFromContext`
   - 或 table row selection API
2. 如果 selection API 不行：
   - 用表格可见区坐标 + 行高计算点击
   - JAB 负责算出目标行，坐标只负责点击当前可见行
3. 如果目标行不在可见区：
   - 用 JAB 读表定位
   - 用滚动条/分页/查找让目标行可见

### TODO 4：继续探测“制单”对话框

当前已经进入：

```text
SunAwtDialog title='制单'
```

下一步在这个对话框里运行：

```bat
tools\query_jab.bat 凭证
tools\query_jab.bat 确定
tools\query_jab.bat 保存
tools\query_jab.bat "role='push button'"
tools\query_jab.bat "role='text'"
```

目标：

- 搞清楚 `常用凭证` 和 `凭证` 两个按钮哪个对应现有流程。
- 看是否有确认按钮、保存按钮、凭证号输入框/显示框。
- 优先用 JAB 读取凭证号，减少坐标点击 + Ctrl+C。

### TODO 5：探测凭证界面和凭证号框

进入凭证界面后运行：

```bat
tools\query_jab.bat 凭证
tools\query_jab.bat 保存
tools\query_jab.bat "role='text'"
tools\query_jab.bat "role='push button'"
```

目标：

- 找到凭证号输入框/显示框
- 找到返回按钮
- 看是否能直接读凭证号，不再需要坐标点击 + Ctrl+C

## 风险和注意点

- JAB 能读控件，但有些控件 `bounds=-1,-1,-1,-1`，不能直接靠坐标。
- 同一界面可能有 `SunAwtFrame` 和 `SunAwtCanvas` 两个窗口副本，优先看 `SunAwtCanvas`，目前业务表格和生成按钮主要在它下面。
- JAB path 可能随界面/版本变化，不应长期硬编码。最终应按 role/name/desc 搜索控件。
- 中文输出需要 UTF-8，否则会乱码。`query_jab.bat` 已设置：

```bat
set PYTHONIOENCODING=utf-8
```

- NC/UClient 关闭后可能保留登录缓存，这不等于进程没关。用 `tools/stop_nc_uclient.bat` 判断/清理进程。

## 建议改造方向

短期最现实方案：

```text
JAB 读表格金额 + JAB 找生成按钮 + 现有键盘/坐标兜底
```

不要一上来完全废弃坐标。当前最稳路线是逐步替换：

1. 保留现有主流程。
2. 新增 `core/jab_operator.py`。
3. 先提供只读能力：读表格金额、找按钮。
4. 再提供动作能力：触发“生成”按钮。
5. 最后再考虑替换选行、凭证号读取等高风险步骤。
