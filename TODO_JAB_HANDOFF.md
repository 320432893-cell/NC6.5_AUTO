# NC JAB 后续 TODO

日期：2026-06-01

当前项目说明在 `README.md`，已实现和验证记录在 `CHANGELOG.md`。本文只保留后续开发待办。

## 2026-06-01 当前进度和结论

- `JABBatchProcessor` 已从大单文件拆成装配入口：
  - `core/nc_state.py` / `core/nc_page_probe.py`：页面状态识别和页面特征探测。
  - `core/nc_pending_workflow.py`：待生成列表、生成入口、恢复当前制单窗口。
  - `core/nc_voucher_workflow.py`：制单窗口匹配、保存、关闭后验证。
  - `core/nc_switch_generated_workflow.py`：切换到已生成/正式单据列表。
  - `core/nc_backfill_workflow.py`：已生成列表凭证号回填。
  - `core/nc_table_matcher.py`：表格匹配和批次构造。
- 页面状态识别已落地为 `pending`、`generated`、`voucher_open`、`query_open`、`loading`、`error`，流程入口已加 `require_page_state()` 守卫。
- 生成、保存、切已生成、回填流程已拆出独立 workflow，并补了状态事件和状态跳转记录。
- 当前没有继续低风险拆分项；后续若继续改，优先做模型/契约/错误类型这类有设计影响的结构化演进。

### 当前不继续做的拆分

- 不继续拆 `NCVoucherWorkflow`，虽然文件仍较长，但它内部保存策略、匹配兜底、验证逻辑耦合较强，继续拆会扩大行为风险。
- 不把状态识别 wrapper 全部从 `JABBatchProcessor` 移除；当前 workflow 统一通过装配入口共享状态检测能力，先保持边界稳定。
- 不把 `__getattr__` 代理立即删掉；下一步若做强类型模型/显式依赖注入，再一并收紧。

## 2026-05-28 当前进度和结论

### `switch-generated` 当前正确链路

已验证从待生成界面切到已生成/正式单据列表的稳定链路：

1. 用 F3 打开查询窗口。
2. 查询窗口内用 JAB AccessibleAction 点击 `正式单据`。
3. 等 `目的业务日期` 条件出现，不用固定 sleep 盲点。
4. 用 JAB `setTextContents` 写入两个日期框：
   - 起始日期 path：`0.0.1.0.0.4.0.1.0.0.0.0.0.11.1.0.0`
   - 结束日期 path：`0.0.1.0.0.4.0.1.0.0.0.0.0.11.1.2.0`
5. 用 JAB AccessibleAction 点击 `确定`。
6. 读取表格验证切换结果。

关键点：

- 主查询入口继续走 F3；不要用 JAB action 点主界面查询按钮，之前会触发 Access Bridge 不稳定。
- `正式单据` 和 `确定` 已验证可走 JAB AccessibleAction，不应再走 `bounds` 坐标点击。
- `目的业务日期` 的条件操作符是 `介于`；只填第一个日期框等价于下限过滤，会带出其他日期。要限定当天，两个日期框都填同一天。
- `setTextContents` 对日期框写入有效，但 `getAccessibleTextInfo/getAccessibleTextRange` 可能读回空字符串；不能把读回空直接判为写入失败。
- F3 后必须用 path guard 等 `正式单据` path 出现；不要固定 sleep 后直接执行下一步。

### 性能现状

已把 `jab.startup_wait` 从 `2.0` 实验降到 `0.5`：

- `startup_wait=0.5`：已验证可用。
- `startup_wait=0.2`：不稳定，曾出现 path guard 等满失败。

已验证快路径 `switch-generated --perf --perf-label fast-guard-test`：

```text
switch_open_query:          0.394s
switch_run_steps:           4.001s
switch_generated_snapshot:  0.446s
switch_generated_total:     5.487s
```

本次快路径验证使用默认 `generated_date_value = date.today()`，即 `2026-05-28`，读到：

```text
rows=11
sample_voucher_count=11
```

### 本次踩坑

- 把 JAB `bounds` 当成底层动作是错误的。`bounds` 只是 JAB 给出的当前控件矩形，最后仍是屏幕坐标点击，窗口前后台、可见性、RemoteApp 映射都可能影响结果。
- `ok=True` 不能直接等同于业务动作成功。必须用后置状态验证，例如点击 `正式单据` 后检查 `目的业务日期` 是否出现。
- 隐藏或 `visible=False` 的查询窗口可能仍能被 JAB 枚举到，不能作为可操作窗口依据。
- 查询窗口右侧筛选区和左侧会计科目树必须严格隔离；自动化不要碰左侧会计科目。
- NC 查询条件区有些输入框视觉上像两行，实际 JAB/布局树里可能像同一个 div/容器里换行显示，仍属于同一个容器或同一行组。不能只按肉眼的“上一行/下一行”推断控件归属，必须结合 label、row 容器、role、bounds 和后置状态验证。
- UI 现象和日志不一致时，应暂停并对齐，不要继续叠加实验。

### 已完成/仍待办

1. `generated_date_value` 已支持显式命令参数 `--generated-date YYYY-MM-DD`。
   - 优先级：命令行、`config.json`、当天日期。

2. `switch_run_steps` 内部耗时已细分为正式单据、起始日期、结束日期、确定等步骤。

3. `backfill` 仍可增加“最近 N 行扫描 + 全表兜底”。
   - 在日期筛选后表格已变小的情况下，优先匹配当前结果集。
   - 如未找到，再全表兜底，避免漏回填。

4. README/CHANGELOG 已记录稳定操作手册和历史变更，后续新增命令或配置时继续同步维护。

## 0. 当前边界

- 当前主线是 Java Access Bridge（JAB）。
- 旧 `pyautogui` 坐标点击、截图识别、固定坐标方案已删除，不再扩展。
- WSL 仓库是源码源头：`/home/queclink/project/nc_auto_v2`。
- H 盘目录只作为 Windows/JAB 运行镜像：`/mnt/h/python脚本/.venv/nc_auto_v2`。
- JAB 实际连接 NC UI 时必须调用 Windows Python。

## 1. 开发/运行流程

1. 同步脚本增强
   - 已新增 `tools/sync_to_windows.sh`。
   - 支持一次同步、`--watch` 自动同步、`--dry-run` 预览。
   - 后续可加同步后自动冒烟：可选执行 `jab_batch.py plan`。

2. README 继续维护
   - README 已改为 JAB 主线。
   - 后续新命令、新配置、新流程优先更新 README。
   - 旧坐标方案只放历史说明，不再展开操作手册。

## 2. 当前功能补强

3. `backfill` 自动切换
   - 当前 `backfill` 假设界面已经在已生成/正式单据列表。
   - 增加参数或默认流程：先 `switch-generated`，再回填。
   - 切换失败时要明确当前界面状态，不要直接按待生成表列位误读。
   - 已验证单据生成页 JAB 查询入口：
     `0.0.1.0.0.0.0.2.0.0.0.0.0.0.0.2`。
   - 主查询入口默认继续走 F3。
   - 查询窗口内部已验证 `正式单据` / `确定` 走 JAB AccessibleAction。
   - 日期筛选走 JAB `setTextContents`，两个日期框都要填。

4. Excel 文件锁检测
   - 写入前检测文件是否被 WPS/Excel 占用。
   - 如果被占用，先报清楚。
   - 避免 NC 已生成成功，但 Excel C 列状态/凭证号写不进去。

5. 保存策略显式配置
   - 保留 `save_strategy=single|safe_batch_by_pending_row|bottom_up`。
   - `single`：当前正式策略，制单窗口一行一保存，凭证号顺序最稳。
   - 正式主线已收敛为：`single + jab_button + use_voucher_queue_cache`。
   - 队列缓存规则：保存并删除制单行 `d` 后，只把缓存中原行号 `> d` 的剩余记录减 1；原行号 `< d` 的记录不变。
   - `bottom_up`：旧策略，只合并制单行号严格递减的连续 Excel 行，已保留但不再默认。
   - `safe_batch_by_pending_row`：快速备选策略。只合并 Excel 顺序中待生成 NC 行号递增、且制单窗口行号递增的连续段；其它自动拆成单张。它整体线性，但不承诺凭证号严格按 Excel 递增。
   - 2026-05-28 已验证历史实验策略 `batch_reverse_select` 和 `batch` 没有继续维护价值，已删除。
   - `Ctrl+S` 保存已验证可用，但端到端闭环未快于 JAB 按钮；保留为实验开关，不作为默认主线。

6. 性能优化
   - 记录每张凭证耗时。
   - 拆分耗时：读表、选行、前台生成、保存、验证、写 Excel。
   - 减少固定 `sleep`，改成等待具体 UI 状态。
   - `switch-generated` 已开始使用 path guard；继续细分步骤耗时。
   - 缓存一次制单表快照，避免每保存一张都做过深控件遍历。
   - 速度目标可以逐步逼近 0.x 秒/张，但不能牺牲验证。
   - 下一步优先做 30-50 张压力样本，拆出生成、保存、切换已生成、回填四段耗时。

7. 异常恢复
   - 生成中断时明确当前 NC 串行状态：
     `单据生成业务页空闲`、`单据生成子窗口打开`、`未知弹窗/未知状态`。
   - Excel B 列写入失败时，不要误判 NC 失败。
   - 制单窗口空表时，按正常完成路径关闭并 F5 验证。
   - 第一版异常弹窗全部停止人工处理，不自动点击确定。
   - 失败后只做保守 cleanup：查询阶段可关闭/取消回业务页；生成菜单阶段可 ESC 回业务页；制单保存中或未知弹窗不自动关闭，保留现场。

## 3. 数据模型和契约

8. 新增 `core/models.py`
   - 把当前 dict 改成 dataclass。
   - 建议模型：
     - `ExcelVoucherItem`
     - `PendingTableRow`
     - `PendingMatch`
     - `VoucherTableRow`
     - `VoucherSaveMatch`
     - `GeneratedVoucherRow`
     - `BackfillUpdate`
   - 目标是减少 `item["row"]`、`row_data["voucher_text"]` 这类 key 写错。

9. 新增 `core/contracts.py`
   - 统一前置/后置契约。
   - 典型契约：
     - 进入生成前：Excel 待处理行必须全部可解析。
     - 选行前：每个 Excel 行必须唯一匹配待生成表。
     - 前台生成后：必须出现 `制单` 窗口。
     - 保存后：目标制单行必须消失、行数减少或制单表为空。
     - 关闭制单后：本轮记录必须从待生成表消失。
     - 回填前：当前表必须是已生成/正式单据表。
   - 契约失败要带上下文：Excel 行、金额、对手方、NC 行、当前窗口。

10. 统一错误类型
   - 替代散落的裸 `RuntimeError`。
   - 建议增加：
     - `WorkflowStateError`
     - `TableMatchError`
     - `ContractViolation`
     - `ExcelLockedError`
     - `JABControlNotFound`

## 4. 状态机

11. 加状态机，但不要过度设计
   - NC 本身按串行窗口模型处理：`主页 -> 业务页 -> 当前业务页唯一子窗口`。
   - 第一版不做主页跳转；默认用户已经进入 `单据生成` 业务页。
   - `单据生成` 业务页已通过 JAB 探测到稳定标识：选中的 `page tab` 名称/描述为 `单据生成`。
   - 查询也是业务页子窗口，不单独作为更高层级。
   - 生成下拉菜单属于业务页动作控件，不是子窗口；`前台生成` 成功后进入制单子窗口。
   - 第一版主状态：
     - `业务页空闲`：当前为 `单据生成`，且无子窗口。
     - `子窗口打开`：当前业务页打开了查询、制单、提示、确认或错误窗口。
     - `未知状态`：业务页或子窗口无法识别，停止。
   - 子窗口类型只在 `子窗口打开` 内识别：
     - 查询：只允许填条件、确定、取消/关闭。
     - 制单：只允许读制单表、保存、关闭。
     - 提示/确认/错误/未知：第一版全部停止人工处理。
   - 用途：
     - 限制非法跳转。
     - 失败时告诉用户卡在哪一步。
     - 支持人工接管后继续。
   - 不要把 GUI 操作改成全异步事件驱动；NC UI 更适合同步、串行状态机。
   - 每个主业务页动作前先确认 `业务页空闲`，每个子窗口动作前先确认当前子窗口类型。
   - 子窗口关闭后必须确认回到原业务页。

12. 状态持久化
   - 写 `logs/run_state.json`。
   - 记录本轮 Excel 行、NC 行、已保存行、已回填行。
   - 用于异常恢复和人工复核。
   - 第一版只记录状态和失败上下文，不自动恢复继续执行。
   - 失败 cleanup 只负责尽量回到业务页空闲，不负责继续业务：
     - 查询阶段失败：可取消/关闭查询。
     - 生成菜单阶段失败：可 ESC 回业务页。
     - 制单未保存或制单表为空：可尝试关闭。
     - 制单已部分保存、异常弹窗、未知窗口：不自动关闭。

## 5. 管道过滤

13. 数据侧用管道过滤
   - 适合 Excel/NC 数据转换。
   - 不适合 GUI 点击。

14. 生成 pipeline
   - `LoadExcelRows`
   - `ParseConcatKey`
   - `DropFilledRows`
   - `ReadPendingSnapshot`
   - `MatchAmountPartner`
   - `ValidateUniqueMatches`
   - `BuildSavePlan`

15. 回填 pipeline
   - `LoadExcelRows`
   - `ReadGeneratedSnapshot`
   - `MatchGeneratedRows`
   - `PreferToday`
   - `NormalizeVoucherNo`
   - `BuildBackfillUpdates`

## 6. 命令模式

16. 抽 NC 操作为命令对象
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

17. 命令模式目标
   - 统一日志。
   - 统一重试。
   - 支持 dry-run。
   - 审计每一步输入输出。

## 7. 事件总线

18. 新增轻量同步事件总线 `core/events.py`
   - 只做进程内同步发布。
   - 不引入消息队列。

19. 事件示例
   - `ExcelRowsLoaded`
   - `PendingRowsMatched`
   - `PendingRowsSelected`
   - `FrontGenerateOpened`
   - `VoucherRowsSaved`
   - `VoucherWindowEmpty`
   - `PendingVerificationPassed`
   - `GeneratedRowsMatched`
   - `ExcelBackfilled`

20. 事件总线用途
   - 日志。
   - 审计。
   - 进度显示。
   - 后续截图/报警/人工确认扩展点。

21. 不要做全事件驱动重写
   - 主流程保持同步 workflow。
   - 事件总线只做观测和扩展。

## 8. 审计和复核

22. 增加 run 日志/审计表
   - 每次运行生成 run id。
   - 记录 Excel 行、金额、对手方、待生成 NC 行、制单行、凭证号、状态。
   - 方便解释“为什么先 370 后 369”这类问题。

23. 回填后复核报表
   - 输出本轮成功、未找到、重复、凭证号异常。
   - 对凭证号倒序或非递增只做提示，不直接判失败。
   - 最终凭证号以已生成表真实结果为准。

## 9. JAB 已知坑提醒

- JAB 不能由 WSL/Linux Python 直接运行，实际控制 NC 时必须调用 Windows Python。
- JAB path 和 hwnd 不稳定，不要长期硬编码。
- JAB `bounds` 不可靠，不要用它恢复坐标点击。
- 隐藏或非 visible 的 `SunAwtDialog` 查询窗口可能是残留，不应作为可操作窗口依据。
- 左上角空白蓝框/截图样遮挡窗口通常是 `SunAwtWindow` 无标题小窗口，优先检查 `hide_blank_awt_windows()`。
- Excel/WPS 打开文件时可能导致写回 `PermissionError`。

## 10. 提交前静态检查目标

- `.venv/bin/python -m json.tool config.json`
- `.venv/bin/ruff check .`
- `.venv/bin/ruff format --check .`
- `.venv/bin/python -m compileall -q core tools`
- `.venv/bin/basedpyright .`
