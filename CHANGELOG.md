# 更新日志

只记录影响维护判断的关键节点。具体实验流水账看 git 历史。

## 2026-06-01 - workflow 拆分和状态守卫

- 新增页面状态识别和状态守卫：`pending`、`generated`、`voucher_open`、`query_open`、`loading`、`error`。
- 页面状态识别同时检查父页面标签、按钮布局、表格数据特征；制单、查询按阻塞式子窗口处理。
- 增加状态事件和状态跳转记录。
- 新增 workflow 领域异常：`WorkflowStateError`、`TableMatchError`、`ContractViolation`、`JABControlNotFound`、`JABActionError`。
- 架构检查阻止 workflow 模块新增裸 `raise RuntimeError(...)`，避免业务失败原因继续混成一种异常。
- 新增 `ExcelLockedError`，Excel 拆分列、生成状态写入、凭证号回填写入遇到文件占用时统一报清楚。
- `backfill` 默认会从 `pending` 自动切到 `generated`；遇到制单、查询、加载、异常状态会停止，不再直接读表。
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
- 开启 JAB 后可能出现左上角空白蓝框/截图样遮挡窗口，通常是 `SunAwtWindow` 无标题小窗口，已通过 `hide_blank_awt_windows()` 处理。
- Excel/WPS 打开文件时，写入 C 列或自动拆分 A/B 时可能报 `PermissionError`。
- 多行批量保存时，凭证号顺序可能不等于 Excel 顺序。严格顺序场景默认使用 `single`。

真实反序案例：

```text
excel_rows=[25, 26]
voucher_rows=[1, 9]
Excel行25 -> 370
Excel行26 -> 369
```
