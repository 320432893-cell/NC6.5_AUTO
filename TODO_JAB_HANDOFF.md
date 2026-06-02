# NC JAB 后续 TODO

日期：2026-06-02

本文只记录后续要做的事和必须记住的坑。已完成的流水账看 git 历史，不在这里重复维护。

## 当前状态

- `JABBatchProcessor` 已收敛为装配入口。
- 主流程已拆成 `nc_pending_workflow`、`nc_voucher_workflow`、`nc_switch_generated_workflow`、`nc_backfill_workflow`、`nc_table_matcher`、`nc_state`。
- 页面状态守卫已覆盖 `pending`、`generated`、`voucher_open`、`query_open`、`loading`、`error`。
- workflow 已引入领域错误类型，并由架构检查阻止新增裸 `raise RuntimeError(...)`。
- Excel 写入锁已包装为 `ExcelLockedError`，覆盖拆分 A/B、写生成状态和回填凭证号。
- `backfill` 默认会从 `pending` 自动切到 `generated`；阻塞/异常状态会停止，避免按错误表格列位读取。
- `backfill` 已记录结构化审计事件 `backfill_audit`，包含 Excel 行、金额、对手方、NC 行、凭证号和失败状态。
- 关键匹配模型已收口为 dataclass：`ExcelVoucherItem`、`PendingMatch`、`GeneratedVoucherMatch`、`VoucherPendingMatch`、`VoucherSaveMatch`、`MatchIssue`。
- `ExcelVoucherItem` 和 `VoucherSaveMatch` 已有统一契约检查；失败会带 Excel 行、金额、对手方、NC 行或制单表位置。
- 待生成表重复匹配属于异常；`generate` 默认 `duplicate_match_policy=stop`，会在点击 NC 前暂停。临时允许跳过异常行时用 `--on-duplicate skip`。
- `receipt_entry` 已记录收款单录入状态标签、财务组织清单、组织-账户映射、Excel 预处理和 NC 已做过匹配规则。
- 收款单查询窗口已枚举到稳定条件行：`收款财务组织`、`单据日期`、`原币金额`、`客户`；`tools/receipt_query_fill.py` 已验证可写入 `A001` 和日期区间，不点确定也能成功填值。
- 当前没有继续低风险拆分项。后续再动结构，应优先做模型和契约，而不是机械拆文件。

## 待办

1. 收款单查询闭环
   - 在已打开的 `收款单录入 -> 查询条件` 窗口中，按主体分组填 `收款财务组织` 和最近 1-2 个月 `单据日期`。
   - 先人工确认一轮 `--confirm` 后列表返回正确，再接自动读表。
   - 读取结果表时使用列语义：`单据日期`、`客户`、`原币金额`；不要依赖视觉列宽。

2. Excel 回写
   - 按 Excel 候选行的 `到款日期 + 原始金额 + 银行来款名` 匹配 NC 的 `单据日期 + 原币金额 + 客户`。
   - 匹配唯一时写 `是否NC已做过`；未找到写未做过；重复命中按异常策略处理，默认暂停。

3. 契约检查
   - 继续观察真实 `plan` / `generate` / 收款查询输出，发现缺字段再补。
   - 已覆盖模型契约不要复制第二份定义，见 `core/models.py`。

## 保留坑点

- JAB 不能由 WSL/Linux Python 直接控制 NC，实际操作必须用 Windows Python。
- JAB path 和 hwnd 不稳定，不要长期硬编码。
- JAB `bounds` 不是底层动作，不要用它恢复坐标点击。
- `ok=True` 只代表 JAB 动作返回成功，业务上必须做后置状态验证。
- 控件探索工具本身也可能截断真相；`tools/jab_probe.py --inspect-path` 曾经固定只展开 1 层，导致右侧查询条件区被误判为空。遇到“控件不存在/为空”时，先核查 `--depth`、`--max-children`、窗口 title/class、visible/showing 过滤是否真的生效。
- `--inspect-path` 输出里的 `path=0...` 是相对目标节点的路径，不是完整窗口路径；写入配置前必须用完整 path 再只读验证一次。
- 主查询入口默认走 F3；用 JAB action 点主界面查询按钮曾触发 Access Bridge 不稳定。
- 查询窗口内部 `正式单据` / `确定` 可走 JAB AccessibleAction，日期框走 `setTextContents`。
- `目的业务日期` 是 `介于`，限定当天必须起止两个日期框都填同一天。
- `setTextContents` 对日期框有效，但 JAB 文本读取可能返回空，不能把读回空直接判定为写入失败。
- 收款单查询条件里的 `收款财务组织` 和 `单据日期` 也存在同样现象：`setTextContents` 成功后 JAB 读回可能为空，以界面状态或后续查询结果验证。
- 隐藏或 `visible=False` 的 `SunAwtDialog` 查询窗口可能是残留，不能作为可操作窗口依据。
- 查询窗口右侧筛选区和左侧会计科目树必须隔离，自动化不要碰左侧会计科目。
- NC 查询条件区视觉换行不等于 JAB 结构，定位要结合 label、role、row 容器、bounds 和后置状态验证。
- 左上角空白蓝框/截图样遮挡窗口通常是 `SunAwtWindow` 无标题小窗口，优先检查 `hide_blank_awt_windows()`。
- Excel/WPS 打开文件时可能导致写回失败；当前会包装成 `ExcelLockedError`，但 NC 已完成的业务动作无法自动撤销。
- 多行批量保存时，凭证号顺序可能不等于 Excel 顺序。严格顺序主线使用 `single`。

## 提交前检查

```bash
.venv/bin/python tools/check.py
```
