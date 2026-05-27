# 更新日志

## 2026-05-27 - JAB 批量凭证生成主线成型

### JAB 查询入口验证

- 已验证单据生成页 JAB 查询入口：
  `0.0.1.0.0.0.0.2.0.0.0.0.0.0.0.2`
- 控件信息为 `查询` / `push button` / `单击`，可以打开查询窗口。
- 实测发现：JAB path 能打开查询窗口，但在同一个进程、同一个 Access Bridge 会话里继续点击查询窗口里的 `正式单据` / `确定` 不稳定，可能找不到控件或卡住。
- 默认策略暂时不启用 `jab_action`，`config.json` 默认保持 `open_query.method=hotkey`，也就是 F3 路径。
- 代码保留 `jab_action` 能力和 `hotkey` fallback，后续继续单独验证。
- 隐藏或非 visible 的 `SunAwtDialog` 查询窗口可能是残留，不应作为可操作窗口依据。

### 静态检查目标

- `.venv/bin/python -m json.tool config.json`
- `.venv/bin/ruff check .`
- `.venv/bin/ruff format --check .`
- `.venv/bin/python -m compileall -q core tools`
- `.venv/bin/basedpyright .`

### 主线变化

- Java Access Bridge（JAB）成为当前主线方案。
- 旧的 `pyautogui` 坐标点击、截图识别、固定坐标方案停止作为新功能方向。
- 旧坐标入口和旧 GUI 模块已删除，只保留 JAB 主线代码。
- 开发迁移到 WSL 仓库：
  `/home/queclink/project/nc_auto_v2`
- H 盘目录作为 Windows/JAB 运行镜像：
  `/mnt/h/python脚本/.venv/nc_auto_v2`
- 明确运行边界：
  - WSL 负责开发、git、文档、重构。
  - JAB 实际操作 NC 时必须调用 Windows Python。

### 新增命令

`tools/jab_batch.py` 新增/确认以下命令：

- `plan`
  - 读取 Excel `Sheet1`。
  - 跳过 C 列已有状态/凭证号的行。
  - 优先解析 A 列 `金额+对手方`，否则按 A 列金额、B 列对手方解析。
  - 读取 NC 待生成主表。
  - 按 `金额 + 对手方` 唯一匹配。
  - 只读规划，不点击 NC。

- `generate --yes`
  - 一次性选中 Excel 全部匹配到的 NC 待生成行。
  - 只点击一次 `生成 -> 前台生成`。
  - 进入同一个 `制单` 窗口。
  - 在制单窗口内按 Excel 顺序查找并保存。
  - 读取 Excel 时每行优先判断 A 列是不是拼接 key；不是拼接则按 A/B 拆分列读取。
  - 读取到 A 列拼接 key 后自动拆回 A/B。
  - 保存后写 Excel C 列 `已生成待回填`。
  - 关闭制单窗口后 F5 刷新待生成表并验证记录消失。

- `switch-generated`
  - 从待生成界面自动进入已生成/正式单据列表。
  - 流程为关闭制单窗口、F3 查询、选择 `正式单据`、点击 `确定`、读取表格确认。

- `backfill`
  - 在已生成列表中按 `金额 + 对手方` 匹配 Excel 中 C 列等于 `已生成待回填` 的行。
  - 历史重复时优先取凭证日期为当天的记录。
  - 读取凭证号列，去前导 0。
  - 校验 `1 <= 凭证号 <= 9999`。
  - 写回 Excel C 列。

- `split-keys`
  - 把 Excel A 列 `金额+对手方` 拆回 A 列金额、B 列对手方。
  - C 列承担状态/凭证号写入。

### 关键实现

- `core/jab_operator.py`
  - 封装 JAB 启动、表格读取、行选择、按钮/菜单动作、窗口关闭、F3/F5。
  - 支持读取待生成表、制单窗口表、已生成表。
  - 增加 `hide_blank_awt_windows()`，用于隐藏 JAB/Java 残留的空白 AWT 小窗。

- `core/jab_batch_processor.py`
  - 实现批量生成主流程。
  - 实现待生成主表全量选择。
  - 实现制单窗口内按 Excel 顺序保存。
  - 实现制单表目标行消失/行数减少/空表验证。
  - 实现待生成表 F5 刷新后消失验证。
  - 实现已生成表凭证号回填。

- `core/data_handler.py`
  - 新增 `load_jab_batch_data()`。
  - 新增 `parse_jab_concat_key()`。
  - 新增 `save_jab_results()`。
  - 新增 `split_jab_keys_to_columns()`。

### 已验证结果

2026-05-27 已完成 33 行端到端测试。

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

已验证 JAB 能力：

- 读取 NC 待生成主表。
- 多选待生成表行。
- 点击 `生成` 和 `前台生成`。
- 读取 `制单` 弹窗表格。
- 在制单表中跨行选择。
- 切到已生成/正式单据列表。
- 读取已生成表凭证号。
- 隐藏空白 AWT 小窗。

### 重要业务结论

- Excel 行顺序来自真实附件顺序，是业务顺序的唯一权威。
- 当前索引始终是 `金额 + 对手方`，不是单据号，也不是单金额。
- 生成阶段必须先在待生成主表一次性选中全部匹配行，再只点一次 `生成 -> 前台生成`。
- 不要每条 Excel 单独生成。
- 不要拆成多轮待生成表 `选几行 -> 生成 -> 前台生成`。
- 不要每保存一张就切已生成表查询；全部生成后统一切已生成表回填。
- 保存成功提示不能作为唯一证据，必须结合制单表消失/行数减少/空表和待生成表 F5 验证。

### 已发现坑

- JAB 不能由 WSL/Linux Python 直接运行，必须调用 Windows Python。
- JAB path 和 hwnd 不稳定，不要长期硬编码。
- JAB `bounds` 不可靠，可能返回负坐标或 `-1,-1,-1,-1`。
- 开启 JAB 后可能出现左上角空白蓝框/截图样遮挡窗口，通常是 `SunAwtWindow` 无标题小窗口，已通过 `hide_blank_awt_windows()` 处理。
- Excel/WPS 打开文件时，写入 C 列或自动拆分 A/B 时可能报 `PermissionError`。
- 多行批量保存时，凭证号顺序可能不等于 Excel 顺序。旧的 `bottom_up` 行号递减策略已保留但不再默认，当前默认 `batch_reverse_select`：按 Excel 顺序组批，反序加入 NC selection，用来验证 NC 疑似“后选中先发号”的规律。

真实案例：

```text
excel_rows=[25, 26]
voucher_rows=[1, 9]
Excel行25 -> 370
Excel行26 -> 369
```

45-54 后续实测中，递减制单行批次仍出现 Excel 行凭证号反向，因此默认改为反序选择批量保存；最终仍以已生成表真实匹配结果回填。

### 相关提交

```text
14417da Split JAB handoff into docs
e401fcc Expand JAB business handoff TODO
c3a5624 Document JAB handoff pitfalls
2f72704 Condense JAB handoff TODO
55b2354 Document handoff notes and pitfalls
b74f18d Add Excel key splitting command
a9f72df Add NC JAB voucher automation workflow
```

## 2026-05-27 - 删除旧坐标流程代码

删除不再属于当前方案的旧文件：

- `main.py`
- `collect_positions.py`
- `core/gui_operator.py`
- `core/test_helper.py`

清理内容：

- `config.json` 删除旧 `positions`、`timing`、`retry`、旧 Sheet2 金额匹配配置。
- `core/data_handler.py` 删除旧金额单索引、重复金额检查、旧进度文件和旧凭证写回方法，只保留 JAB Excel key 读写。
- `core/utils.py` 删除旧窗口激活、DPI、健康检查、紧急恢复等坐标流程工具，只保留配置读取和中断检测。
- `core/logger.py` 删除旧截图 recorder。

保留内容：

- JAB 主流程。
- JAB 探测工具。
- `pyautogui` 在 `core/jab_operator.py` 中仅用于发送 F3/F5 等键盘按键。
- `keyboard` 在 `core/utils.py` 中仅用于空格/ESC 中断检测。

## 2026-04-07 - 旧坐标方案优化记录

此节为历史记录。旧坐标方案已删除，不再作为当前主线维护。

### 优化内容

- 查找窗口输入从坐标点击改为 `Ctrl+F` 后直接输入。
- 保存凭证从坐标点击改为 `Ctrl+S`。
- 凭证界面等待时间增加为可配置。
- 凭证号校验逻辑从“处理当前行之前校验上一条”改为“处理当前行之后、写入 Excel 之前校验”。

这些变更涉及的文件后来已被 JAB 主线替换或清理，仅作为历史背景保留。
