# NC JAB 自动化交接

日期：2026-05-27

## 结论

当前只维护 Java Access Bridge（JAB）方案。以前的 `pyautogui` 坐标点击方案已经废弃，后续不要再基于坐标补功能。

主入口：

- `tools/jab_batch.py`
- `core/jab_operator.py`
- `core/jab_batch_processor.py`

运行必须用 Windows Python：

```powershell
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py <command>
```

WSL 仓库：`/home/queclink/project/nc_auto_v2`
Windows 运行目录：`H:\python脚本\.venv\nc_auto_v2`
当前 Excel：`C:\Users\Queclink\Desktop\5.27凭证.xlsx`

## 已实现

- `plan`：只读 Excel 和 NC 待生成表，按 `金额 + 对手方` 唯一匹配，不点击。
- `generate --yes`：一次性选中 Excel 全部匹配行，只点一次 `生成 -> 前台生成`，进入同一个 `制单` 窗口后按 Excel 顺序保存，并写 `已生成待回填`。
- `switch-generated`：从待生成界面自动进入已生成/正式单据列表。
- `backfill`：在已生成列表按 `金额 + 对手方` 回填凭证号，去前导 0，校验 `1-9999`。
- `split-keys`：把 Excel A 列 `金额+对手方` 拆到默认 C/D 两列，A/B 保持不动。

常用命令：

```powershell
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py plan
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py generate --yes
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py switch-generated
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py backfill
.\.venv-local\Scripts\python.exe .\tools\jab_batch.py split-keys
```

## 当前正确流程

1. Excel `Sheet1` A 列是 `金额+对手方`，B 列用于状态/凭证号。
2. 在 NC 待生成表读取全部数据。
3. 按 Excel 行顺序匹配 NC 行。
4. 一次性选中所有匹配到的待生成行。
5. 只点一次 `生成 -> 前台生成`。
6. 在 `制单` 弹窗内按 Excel 顺序逐条找行并保存。
7. 制单表为空或目标行消失后，关闭制单窗口。
8. 回待生成表 F5 刷新，验证本轮记录已消失。
9. 统一切到已生成表，按 `金额 + 对手方` 回填凭证号。

## 已验证

2026-05-27 已完成 33 行端到端测试：

- 待生成主表一次性选中 24 条。
- 只点一次 `生成 -> 前台生成`。
- 制单窗口出现 24 条，保存后递减到 0。
- 待生成表从 246 行变成 222 行。
- 已生成表凭证号回填成功。

## 关键列

- 待生成表：金额 `col=4`，对手方 `col=3`，选择列 `col=0`。
- 制单窗口：`SunAwtDialog title='制单'`，按整行文本匹配对手方。
- 已生成表：金额 `col=4`，对手方 `col=3`，凭证日期 `col=18`，凭证号 `col=22`。

## 踩坑

- 不要回到坐标方案；JAB 已能读表、选行、点按钮、切查询、回填凭证号。
- 不要拆成多轮 `选几行 -> 生成 -> 前台生成`；只选 1 行时 NC 界面可能变形。
- 凭证号顺序严格时，制单窗口内应一行一保存；多行批量保存可能出现 Excel 行 25 回填 370、行 26 回填 369 这种倒序。
- 制单窗口还在但表为空是正常完成状态；关闭窗口、F5 刷新待生成表再验证。
- 不要只信“保存成功”；以制单表目标行消失、行数减少、空表、待生成表刷新后消失为准。
- Excel/WPS 打开文件会导致写回 `PermissionError`，运行前关闭 Excel。
- PowerShell 参数不要直接换行；写一行，或用反引号续行。

## 下一步

1. 给 `backfill` 增加自动 `switch-generated`，避免人工先切界面。
2. 写入 Excel 前做文件锁检测。
3. 把当前 dict 数据整理为 `dataclass`，再加最小契约校验。
4. 状态机只用于约束流程阶段；事件总线只用于日志、审计、进度，不要把 GUI 主流程改成全异步。

## 最近提交

```text
55b2354 Document handoff notes and pitfalls
b74f18d Add Excel key splitting command
a9f72df Add NC JAB voucher automation workflow
```
