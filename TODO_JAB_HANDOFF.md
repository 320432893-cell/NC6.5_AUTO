# NC JAB 后续 TODO

日期：2026-06-01

本文只记录后续要做的事和必须记住的坑。已完成的流水账看 git 历史，不在这里重复维护。

## 当前状态

- `JABBatchProcessor` 已收敛为装配入口。
- 主流程已拆成 `nc_pending_workflow`、`nc_voucher_workflow`、`nc_switch_generated_workflow`、`nc_backfill_workflow`、`nc_table_matcher`、`nc_state`。
- 页面状态守卫已覆盖 `pending`、`generated`、`voucher_open`、`query_open`、`loading`、`error`。
- workflow 已引入领域错误类型，并由架构检查阻止新增裸 `raise RuntimeError(...)`。
- Excel 写入锁已包装为 `ExcelLockedError`，覆盖拆分 A/B、写生成状态和回填凭证号。
- 当前没有继续低风险拆分项。后续再动结构，应优先做模型和契约，而不是机械拆文件。

## 待办

1. `backfill` 自动切换
   - 当前假设界面已经在已生成/正式单据列表。
   - 可增加参数或默认流程：先 `switch-generated`，再回填。
   - 切换失败时必须报告当前页面状态，不能按错误表格列位继续读。

2. 数据模型
   - 用 dataclass 替代关键 dict。
   - 优先模型：`ExcelVoucherItem`、`PendingMatch`、`VoucherSaveMatch`、`GeneratedVoucherRow`、`BackfillUpdate`。
   - 目标是减少 `item["row"]`、`row_data["voucher_text"]` 这类 key 写错。

3. 契约检查
   - 继续补统一前置/后置检查。
   - 失败要带 Excel 行、金额、对手方、NC 行、当前窗口。

4. 审计和复核
   - 每次运行生成 run id。
   - 记录 Excel 行、金额、对手方、待生成 NC 行、制单行、凭证号、状态。
   - 回填后输出成功、未找到、重复、凭证号异常。

## 保留坑点

- JAB 不能由 WSL/Linux Python 直接控制 NC，实际操作必须用 Windows Python。
- JAB path 和 hwnd 不稳定，不要长期硬编码。
- JAB `bounds` 不是底层动作，不要用它恢复坐标点击。
- `ok=True` 只代表 JAB 动作返回成功，业务上必须做后置状态验证。
- 主查询入口默认走 F3；用 JAB action 点主界面查询按钮曾触发 Access Bridge 不稳定。
- 查询窗口内部 `正式单据` / `确定` 可走 JAB AccessibleAction，日期框走 `setTextContents`。
- `目的业务日期` 是 `介于`，限定当天必须起止两个日期框都填同一天。
- `setTextContents` 对日期框有效，但 JAB 文本读取可能返回空，不能把读回空直接判定为写入失败。
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
