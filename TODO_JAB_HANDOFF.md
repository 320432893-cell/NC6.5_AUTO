# NC JAB 后续 TODO

日期：2026-05-27

当前项目说明在 `README.md`，已实现和验证记录在 `CHANGELOG.md`。本文只保留后续开发待办。

## 0. 当前边界

- 当前主线是 Java Access Bridge（JAB）。
- 旧 `pyautogui` 坐标点击、截图识别、固定坐标方案不再扩展。
- WSL 仓库是源码源头：`/home/queclink/project/nc_auto_v2`。
- H 盘目录只作为 Windows/JAB 运行镜像：`/mnt/h/python脚本/.venv/nc_auto_v2`。
- JAB 实际连接 NC UI 时必须调用 Windows Python。

## 1. 开发/运行流程

1. 增加同步脚本
   - 新增 `tools/sync_to_windows.sh`。
   - 从 WSL 同步到 H 盘运行镜像。
   - 排除 `.git`、`.venv-local`、缓存、日志大文件。
   - 同步后可选跑 `jab_batch.py plan` 做冒烟。

2. README 继续维护
   - README 已改为 JAB 主线。
   - 后续新命令、新配置、新流程优先更新 README。
   - 旧坐标方案只放历史兼容/已废弃说明，不再展开操作手册。

3. 配置清理
   - `config.json` 里仍有 `positions`、`timing`、`retry` 等旧坐标配置。
   - 后续确认旧入口不再使用后，拆分为 `config.legacy.json` 或标注 deprecated。

## 2. 当前功能补强

4. `backfill` 自动切换
   - 当前 `backfill` 假设界面已经在已生成/正式单据列表。
   - 增加参数或默认流程：先 `switch-generated`，再回填。
   - 切换失败时要明确当前界面状态，不要直接按待生成表列位误读。

5. Excel 文件锁检测
   - 写入前检测文件是否被 WPS/Excel 占用。
   - 如果被占用，先报清楚。
   - 避免 NC 已生成成功，但 Excel B 列状态/凭证号写不进去。

6. 保存策略显式配置
   - 新增 `save_strategy=single|batch`。
   - `single`：制单窗口一行一保存，凭证号顺序最稳。
   - `batch`：制单窗口递增批量保存，速度更快但凭证号可能不递增。
   - 默认建议先用 `single`，等回填和审计稳定后再实验 `batch`。

7. 性能优化
   - 记录每张凭证耗时。
   - 拆分耗时：读表、选行、前台生成、保存、验证、写 Excel。
   - 减少固定 `sleep`，改成等待具体 UI 状态。
   - 缓存一次制单表快照，避免每保存一张都做过深控件遍历。
   - 速度目标可以逐步逼近 0.x 秒/张，但不能牺牲验证。

8. 异常恢复
   - 生成中断时明确当前阶段：
     `待生成表`、`制单窗口`、`已生成表`、`查询窗口`。
   - Excel B 列写入失败时，不要误判 NC 失败。
   - 制单窗口空表时，按正常完成路径关闭并 F5 验证。

## 3. 数据模型和契约

9. 新增 `core/models.py`
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

10. 新增 `core/contracts.py`
   - 统一前置/后置契约。
   - 典型契约：
     - 进入生成前：Excel 待处理行必须全部可解析。
     - 选行前：每个 Excel 行必须唯一匹配待生成表。
     - 前台生成后：必须出现 `制单` 窗口。
     - 保存后：目标制单行必须消失、行数减少或制单表为空。
     - 关闭制单后：本轮记录必须从待生成表消失。
     - 回填前：当前表必须是已生成/正式单据表。
   - 契约失败要带上下文：Excel 行、金额、对手方、NC 行、当前窗口。

11. 统一错误类型
   - 替代散落的裸 `RuntimeError`。
   - 建议增加：
     - `WorkflowStateError`
     - `TableMatchError`
     - `ContractViolation`
     - `ExcelLockedError`
     - `JABControlNotFound`

## 4. 状态机

12. 加状态机，但不要过度设计
   - 建议状态：
     `excel_loaded -> pending_snapshot_read -> pending_rows_selected -> front_generate_clicked -> voucher_window_opened -> voucher_saving -> voucher_empty -> pending_refreshed -> generated_opened -> backfilled`
   - 用途：
     - 限制非法跳转。
     - 失败时告诉用户卡在哪一步。
     - 支持人工接管后继续。
   - 不要把 GUI 操作改成全异步事件驱动；NC UI 更适合同步状态机。

13. 状态持久化
   - 写 `logs/run_state.json`。
   - 记录本轮 Excel 行、NC 行、已保存行、已回填行。
   - 用于异常恢复和人工复核。

## 5. 管道过滤

14. 数据侧用管道过滤
   - 适合 Excel/NC 数据转换。
   - 不适合 GUI 点击。

15. 生成 pipeline
   - `LoadExcelRows`
   - `ParseConcatKey`
   - `DropFilledRows`
   - `ReadPendingSnapshot`
   - `MatchAmountPartner`
   - `ValidateUniqueMatches`
   - `BuildSavePlan`

16. 回填 pipeline
   - `LoadExcelRows`
   - `ReadGeneratedSnapshot`
   - `MatchGeneratedRows`
   - `PreferToday`
   - `NormalizeVoucherNo`
   - `BuildBackfillUpdates`

## 6. 命令模式

17. 抽 NC 操作为命令对象
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

18. 命令模式目标
   - 统一日志。
   - 统一重试。
   - 支持 dry-run。
   - 审计每一步输入输出。

## 7. 事件总线

19. 新增轻量同步事件总线 `core/events.py`
   - 只做进程内同步发布。
   - 不引入消息队列。

20. 事件示例
   - `ExcelRowsLoaded`
   - `PendingRowsMatched`
   - `PendingRowsSelected`
   - `FrontGenerateOpened`
   - `VoucherRowsSaved`
   - `VoucherWindowEmpty`
   - `PendingVerificationPassed`
   - `GeneratedRowsMatched`
   - `ExcelBackfilled`

21. 事件总线用途
   - 日志。
   - 审计。
   - 进度显示。
   - 后续截图/报警/人工确认扩展点。

22. 不要做全事件驱动重写
   - 主流程保持同步 workflow。
   - 事件总线只做观测和扩展。

## 8. 审计和复核

23. 增加 run 日志/审计表
   - 每次运行生成 run id。
   - 记录 Excel 行、金额、对手方、待生成 NC 行、制单行、凭证号、状态。
   - 方便解释“为什么先 370 后 369”这类问题。

24. 回填后复核报表
   - 输出本轮成功、未找到、重复、凭证号异常。
   - 对凭证号倒序或非递增只做提示，不直接判失败。
   - 最终凭证号以已生成表真实结果为准。

## 9. JAB 已知坑提醒

- JAB 不能由 WSL/Linux Python 直接运行，实际控制 NC 时必须调用 Windows Python。
- JAB path 和 hwnd 不稳定，不要长期硬编码。
- JAB `bounds` 不可靠，不要用它恢复坐标点击。
- 左上角空白蓝框/截图样遮挡窗口通常是 `SunAwtWindow` 无标题小窗口，优先检查 `hide_blank_awt_windows()`。
- Excel/WPS 打开文件时可能导致写回 `PermissionError`。
