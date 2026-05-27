# 更新日志

## 2026-04-07 - 优化查找和保存流程

### 优化内容

| 优化点 | 当前方案 | 优化方案 | 优势 |
|--------|---------|---------|------|
| 查找窗口输入 | 点击输入框坐标 → Ctrl+A → Delete → 输入 | Ctrl+F 后直接 Ctrl+A → Backspace → 输入 | 减少坐标采集，更稳定，焦点自动在输入框 |
| 查找窗口清空 | Delete 键 | Backspace 键 | 更符合用户习惯 |
| 保存凭证 | 点击保存按钮坐标 | Ctrl+S 快捷键 | 更快、更稳定、不受界面布局变化影响 |
| 凭证界面等待 | 0.2秒 | 0.5秒（可配置 `voucher_load_wait`） | 确保界面完全加载后再保存 |

### 修改文件

- `core/gui_operator.py`
  - `find_amount()`: 去掉点击输入框坐标，改用 Backspace 清空
  - `do_save()`: 改用 Ctrl+S 保存，增加界面加载等待时间
- `config.json`
  - 新增 `timing.voucher_load_wait: 0.5` 配置项
- `README.md`
  - 更新坐标采集说明（输入框和保存按钮无需采集）
  - 更新配置参数表格（新增 `voucher_load_wait`）

### 减少的坐标采集

- ~~`find_target_input`~~: 查找窗口输入框（焦点自动在输入框）
- ~~`save_btn` 坐标点击~~: 保存按钮（改用 Ctrl+S，但图像识别仍保留用于确认界面已打开）

### 凭证号校验逻辑修正

| 修改点 | 错误逻辑 | 正确逻辑 | 说明 |
|--------|---------|---------|------|
| 校验时机 | 处理当前行**之前**校验上一条 | 处理当前行**之后**、写入 Excel **之前**校验 | 避免第一行处理时检查不存在的"上一条" |
| 校验方式 | 从 NC 界面复制凭证号对比 | 从 Excel 中读取上一行凭证号对比 | 每次写入后立即保存 Excel，下一行从 Excel 读取对比 |
| 凭证号长度 | ≤ 4 位 | ≤ 6 位 | 修正为正确的长度限制 |
| 第一行处理 | 无特殊处理 | 检测到无上一行时跳过对比 | 第一行数据无需对比 |

### 修改的方法

- `gui_operator.py`
  - 删除 `ensure_previous_voucher_changed()` 方法
  - 删除 `self.last_success_voucher` 属性
  - `validate_voucher()`: 长度限制从 4 位改为 6 位
  - `process_one()`: 步骤从 7 步改为 6 步，在写入 Excel 前增加对比逻辑
- `data_handler.py`
  - 新增 `get_previous_voucher(current_row)` 方法：读取 Excel 中上一行的凭证号
