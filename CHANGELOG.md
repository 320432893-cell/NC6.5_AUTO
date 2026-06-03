# 更新日志

只记录影响维护判断的关键节点。具体实验流水账看 git 历史。

## 2026-06-03 - 收款单查询结果抽取骨架

- 新增 `ReceiptNCResultExtractor`，按 NC 结果表表头语义抽取 `单据日期`、`原币金额`、`客户`，不依赖固定视觉列位。
- 收款单结果抽取会跳过空行；坏日期、坏金额、空客户会返回结构化抽取问题。
- `tools/receipt_query_fill.py` 默认会先用 F3 打开收款单查询条件窗口；`--no-open-query` 仅用于窗口已打开时的诊断。
- `tools/receipt_query_fill.py` 新增 `--read-results`，可在填条件或 `--confirm` 查询后读取可见 NC 表格并打印前 20 条抽取结果/问题。
- 收款单 dry-run 查询后会把结果页每页条数改为 500，并按分页逐页读取；总数超过 500 时通过“下页”继续采集。
- “下页”优先走 JAB 动作接口，失败时按同一路径做 bounds 点击回退；dry-run 输出会携带分页报告。
- 分页 path 查找新增 `scope_hwnd` 保护：收款查询结果区的 `SunAwtCanvas` 标题为空，必须先锁定分页 label 所在窗口，再在同一 hwnd 内改每页条数/点下页，避免串到其他 Java Canvas。
- 收款 dry-run 默认改为只读取匹配所需的 NC 结果列，降低大结果表读取耗时；这只是降耗优化，不再作为 NC 重启根因记录。
- 收款分页加入稳定等待：查询后、改每页条数后、读表前、下页后都等待结果表稳定，避免 NC 结果表刷新未完成时继续读表/翻页/改分页导致状态混乱。
- 收款分页等待从纯 sleep 升级为稳定判定：分页 label 和结果表摘要连续一致后才改每页条数、读表或翻页。
- 收款分页固定等待收紧为 0.5s 级短等待，主要依赖稳定判定放行。
- 收款分页去掉重复稳定等待：默认跳过改每页条数前的完整稳定轮询，并复用“改页后/下页后”的稳定结果作为后续读表依据。
- 收款 dry-run 修复分页聚合：多页采集到的同结构结果表会全部进入索引抽取，并按单据号去重，不再只取最大页表。
- 收款 dry-run 输出新增 NC 摘要：金额范围、名称样本和名称金额命中诊断；日期只用于查询范围，不参与匹配诊断。
- 收款单金额解析支持 NC 显示的 `- 1,368.10` 和 `(1,368.10)` 负数格式。
- 收款单结果列位按 NC/JAB 0 基索引配置：名称列 `2`，金额列 `7`；`8` 是 Excel/人工 1 基口径误填，已加配置校验拦截。
- 已在真实 NC 跑通稳定判定分页：`659` 条记录按 `500 + 159` 两页采集完成，并全部进入 dry-run 索引抽取。

## 2026-06-02 - 模型、契约和收款单查询准备

- `ExcelVoucherItem`、`PendingMatch`、`GeneratedVoucherMatch`、`VoucherPendingMatch`、`VoucherSaveMatch`、`MatchIssue` 收口为 dataclass。
- 删除模型的字典访问兼容层，workflow 改为属性访问。
- `ExcelVoucherItem` 增加处理前契约检查：无解析错误时必须有正 Excel 行号、金额和对手方。
- `VoucherSaveMatch` 增加保存前契约检查：制单表索引、表行数、制单行号和单元格内容必须有效。
- `match_generated_voucher_table` 显式返回 `GeneratedVoucherMatch`，不再复用待生成匹配类型。
- 待生成表重复匹配定义为异常；`generate` 默认暂停在点击 NC 前，显式传 `--on-duplicate skip` 时写入异常行并跳过继续。
- `config.json` 新增 `receipt_entry`，记录收款单录入状态标签、财务组织清单和组织-账户映射，并纳入配置校验。
- 收款单 Excel 预处理已支持按银行映射主体，候选行默认限定最近 2 个月且跳过已有状态。
- 收款单匹配规则已落地：Excel `原始金额 + 银行来款名` 对齐 NC `原币金额 + 名称列`；客户名做归一化和包含匹配；日期只用于查询/候选范围，不参与匹配；重复命中不执行并报告“名称和金额相同，需人工确认”。
- 收款单查询窗口已枚举到右侧条件行，并将 `收款财务组织`、`单据日期`、`原币金额`、`客户` 的 JAB path 写入配置校验。
- 新增 `tools/receipt_query_fill.py`，已验证能在 `查询条件` 窗口写入 `A001` 和日期区间；默认不点确定，`--confirm` 才触发查询。
- 修复 `tools/jab_probe.py --inspect-path` 忽略 `--depth` 的问题；该问题曾导致右侧查询条件区被误判为空。

保留坑点：

- `--inspect-path` 的输出路径是相对被 inspect 的目标节点，不是完整窗口路径；配置 JAB path 前必须用完整 path 复验。
- `setTextContents` 对 NC 查询条件文本框可写入，但 JAB 文本读回可能为空，不能仅凭读回为空判失败。

## 2026-06-01 - workflow 拆分和状态守卫

- 新增页面状态识别和状态守卫：`pending`、`generated`、`voucher_open`、`query_open`、`loading`、`error`。
- 页面状态识别同时检查父页面标签、按钮布局、表格数据特征；制单、查询按阻塞式子窗口处理。
- 增加状态事件和状态跳转记录。
- 新增 workflow 领域异常：`WorkflowStateError`、`TableMatchError`、`ContractViolation`、`JABControlNotFound`、`JABActionError`。
- 架构检查阻止 workflow 模块新增裸 `raise RuntimeError(...)`，避免业务失败原因继续混成一种异常。
- 新增 `ExcelLockedError`，Excel 拆分列、生成状态写入、凭证号回填写入遇到文件占用时统一报清楚。
- `backfill` 默认会从 `pending` 自动切到 `generated`；遇到制单、查询、加载、异常状态会停止，不再直接读表。
- `backfill` 增加结构化审计事件 `backfill_audit`，记录 Excel 行、金额、对手方、NC 行、凭证号和失败状态。
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
