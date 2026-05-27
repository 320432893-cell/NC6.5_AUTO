# NC6.5 自动生成凭证脚本

用于在 NC6.5 中按 Excel 金额查找记录、生成凭证、保存并把凭证号写回 Excel。

## 环境准备

建议在 Windows 的 Python 环境中运行，不建议在 WSL/Linux 中运行，因为 `pyautogui`、`keyboard`、窗口激活等功能需要直接控制 Windows 桌面和 NC 窗口。

安装依赖：

```bash
python -m pip install pyautogui pyperclip openpyxl keyboard Pillow pygetwindow
```

如果电脑上有多个 Python，先确认正在使用的解释器：

```bash
python --version
python -m pip --version
```

## 文件结构

```text
nc_auto_v2/
├── config.json           # 配置：Excel路径、坐标、等待时间、重试次数等
├── collect_positions.py  # 坐标采集工具
├── main.py               # 主程序入口
├── core/
│   ├── __init__.py
│   ├── data_handler.py   # Excel 数据读写
│   ├── gui_operator.py   # NC GUI 操作
│   ├── logger.py         # 日志
│   ├── test_helper.py    # 测试工具
│   └── utils.py          # 工具函数
├── images/               # 截图素材，可选
└── logs/                 # 运行日志，自动生成
```

## 使用步骤

### 1. 准备 Excel

打开你的 Excel 文件：

- `Sheet1`：你的数据，A 列放金额，B 列留空，脚本会写入凭证号。
- `Sheet2`：从 NC 导出的财务数据，E 列是来源金额。

### 2. 修改配置

打开 `config.json`，修改 `excel_path` 为你的 Excel 文件路径。

其他参数一般不需要改，坐标由采集工具自动写入。

### 3. 采集坐标

```bash
python collect_positions.py
```

按提示操作，分四轮采集：

1. 主界面：来源金额列单元格、生成按钮。
2. 查找窗口：手动按 `Ctrl+F` 打开查找窗口，采集查找下一个、关闭。
3. 生成菜单：关闭查找窗口，点击生成按钮，采集前台生成选项。
4. 凭证界面：点击前台生成进入凭证界面，采集凭证号输入框、返回按钮。

每个位置把鼠标移过去按空格即可。

### 4. 运行测试

```bash
python main.py
```

首次运行会进入测试菜单：

- 选 1：鼠标会依次移到每个坐标位置画圈，确认位置是否正确。
- 选 2：测试查找窗口流程。
- 选 3：测试生成流程。
- 选 4：跳过测试直接运行。

建议首次选 1 确认坐标，再选 4 正式运行。

### 5. 正式运行

确认测试通过后，切换到 NC 窗口，脚本会自动开始处理。

运行过程中：

- 按空格可紧急停止，按 `ESC` 也可中断，已处理的数据不会丢失。
- 中断后再次运行会自动跳过已有凭证号的行。
- 日志保存在 `logs/` 目录。

## 配置说明

### config.json 参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `excel_path` | Excel 文件路径 | 需修改 |
| `sheet_my` | 你的数据 Sheet 名 | `Sheet1` |
| `sheet_finance` | 财务数据 Sheet 名 | `Sheet2` |
| `my_amount_col` | 你的金额列号 | `1`，A 列 |
| `my_voucher_col` | 凭证号写入列号 | `2`，B 列 |
| `finance_amount_col` | 财务金额列号 | `5`，E 列 |
| `has_header` | 是否有表头 | `true` |
| `debug_screenshots` | 是否保存调试截图 | `false` |
| `verify_amount` | 是否启用金额验证 | `false` |
| `verify_input` | 是否验证输入 | `true` |
| `ensure_window_focus` | 是否确保窗口焦点 | `true` |
| `use_mouse_movement` | 点击前是否先移动鼠标 | `true` |
| `health_check_interval` | 每处理多少条做一次健康检查 | `10` |
| `retry.max_retries` | 单条记录最大重试次数 | `2` |
| `retry.copy_retries` | 复制凭证号重试次数 | `3` |
| `retry.click_retries` | 点击验证失败后的重试次数 | `2` |

## 等待时间总览

### config.json 可配置等待

| 参数 | 默认值 | 代码位置 | 作用 |
|------|--------|----------|------|
| `timing.action_delay` | `0.2` 秒 | `core/gui_operator.py` | 每次点击坐标后等待，多个按钮点击都会用到。 |
| `timing.window_wait` | `0.8` 秒 | `config.json` | 预留的窗口弹出等待参数，当前代码没有直接使用。 |
| `timing.save_timeout` | `30` 秒 | `config.json` | 预留的保存超时参数，当前代码没有直接使用。 |
| `timing.page_timeout` | `15` 秒 | `wait_for_image()` | 等待页面或截图出现的默认超时时间。 |
| `timing.between_tasks` | `0.2` 秒 | `process_one()` | 每条记录处理成功后，进入下一条前等待。 |
| `timing.input_interval` | `0.02` 秒 | `config.json` | 预留的输入间隔参数，当前正式流程使用剪贴板粘贴，没有直接使用。 |
| `timing.mouse_move_duration` | `0.05` 秒 | `click_pos()` | 点击前鼠标移动到目标坐标的耗时。 |
| `timing.voucher_load_wait` | `0.5` 秒 | `do_save()` | 检测到凭证界面后，再额外等待界面加载完成，然后按 `Ctrl+S`。 |

### 正式流程固定等待

这些等待写在代码里，不能通过 `config.json` 直接修改。

| 等待时间 | 代码位置 | 触发场景 |
|----------|----------|----------|
| `0.1` 秒 | `pyautogui.PAUSE` | 每个 `pyautogui` 动作后的全局停顿。 |
| `2` 秒 | `wait_for_image()` | 如果 `images/` 中缺少目标截图，降级为固定等待 2 秒并认为成功。 |
| `0.3` 秒/轮 | `wait_for_image()` | 图像识别轮询间隔，直到识别成功或超时。 |
| `0.05` 秒 | `click_pos()` | 鼠标移动到坐标后、点击前的短暂停顿。 |
| `2` 秒超时 | `click_pos_with_verify()` | 点击后如需验证截图，单次等待验证截图出现的超时时间。 |
| `0.1` 秒 | `find_amount()` | 金额复制到剪贴板后等待。 |
| `0.05` 秒 | `find_amount()` | 查找框 `Ctrl+A` 后等待。 |
| `0.2` 秒 | `find_amount()` | 查找框 `Ctrl+V` 粘贴后等待。 |
| `0.5` 秒 | `find_amount()` | 点击“查找下一个”后等待结果定位。 |
| `0.6` 秒超时 | `verify_found_amount()` | 检测 `not_found.png` 的超时时间，用于判断“没有找到”。 |
| `0.5` 秒 | `do_generate()` | 点击“生成”后等待生成菜单展开。 |
| `0.2` 秒 | `do_save()` | 按 `Ctrl+S` 保存后等待。 |
| `0.1` 秒 | `copy_voucher_num()` | 清空剪贴板后等待。 |
| `0.2` 秒 | `copy_voucher_num()` | 点击凭证号框后等待。 |
| `0.1` 秒 | `copy_voucher_num()` | 凭证号框 `Ctrl+A` 后等待。 |
| `0.3` 秒 | `copy_voucher_num()` | `Ctrl+C` 后等待剪贴板更新。 |
| `0.5` 秒 | `copy_voucher_num()` | 凭证号无效或复制异常后的重试等待。 |
| `10` 秒超时 | `process_one()` | 异常恢复后等待主界面 `generate_btn.png` 出现。 |
| `1` 秒 | `process_one()` | 单条记录失败但未达到最大重试次数时，下一次重试前等待。 |
| `5 * 0.3` 秒 | `emergency_recovery()` | 异常恢复时连续按 5 次 `ESC`，每次间隔 0.3 秒。 |
| `5` 秒超时 | `health_check()` | 健康检查等待主界面 `generate_btn.png` 出现。 |
| `0.2 + 0.2` 秒 | `health_check()` | 健康检查时鼠标来回移动两次，每次 0.2 秒。 |
| `0.5` 秒 | `ensure_nc_active()` | NC 窗口重新激活后等待。 |
| `3` 秒 | `main.py` | 正式开始前倒计时，3、2、1 每秒等待一次。 |

### 测试模式固定等待

| 等待时间 | 代码位置 | 触发场景 |
|----------|----------|----------|
| `3` 秒 | `test_all_positions()` | 坐标测试开始前等待。 |
| `0.5` 秒 | `test_all_positions()` | 鼠标移动到每个坐标的耗时。 |
| `8 * 0.1` 秒 | `test_all_positions()` | 围绕坐标画圈，每圈 8 个点，每点移动 0.1 秒。 |
| `0.2` 秒 | `test_all_positions()` | 画圈后鼠标回到坐标中心。 |
| `2` 秒 | `test_all_positions()` | 每个坐标展示完成后等待人工确认。 |
| `3` 秒 | `test_find_window()` | 查找窗口测试开始前等待。 |
| `0.5` 秒 | `test_find_window()` | 点击来源金额单元格后等待。 |
| `1` 秒 | `test_find_window()` | 按 `Ctrl+F` 后等待查找窗口出现。 |
| `0.02` 秒/字符 | `test_find_window()` | 测试输入 `123.45` 的字符间隔。 |
| `0.5` 秒 | `test_find_window()` | 输入后等待。 |
| `0.5` 秒 | `test_find_window()` | 点击“查找下一个”后等待。 |
| `3` 秒 | `test_generate_flow()` | 生成流程测试开始前等待。 |
| `0.5` 秒 | `test_generate_flow()` | 点击“生成”后等待菜单展开。 |

## 等待时间怎么调

- NC 响应慢、点击后界面还没切换：优先调大 `timing.action_delay`。
- 查找、返回、凭证界面截图经常超时：调大 `timing.page_timeout`。
- 凭证界面刚出现就保存失败：调大 `timing.voucher_load_wait`。
- 每条记录之间想更稳一点：调大 `timing.between_tasks`。
- 鼠标移动太快导致点击不稳：调大 `timing.mouse_move_duration`。
- 需要修改固定等待时，直接改上表对应的代码位置。

## 适配其他场景

如果要操作其他字段或其他流程：

1. 在 `config.json` 的 `positions` 中添加新坐标 key。
2. 在 `collect_positions.py` 的 `POSITIONS_TO_COLLECT` 中添加对应描述。
3. 在 `core/gui_operator.py` 中添加新的操作方法。
4. 修改 `process_one()` 的流程。

## 常见问题

Q: 坐标不准怎么办？

A: 重新运行 `collect_positions.py` 采集。

Q: 中途失败了怎么办？

A: 直接重新运行 `main.py`，会自动跳过已处理的行。

Q: 金额重复怎么办？

A: 脚本会标记重复记录，需要手动在 NC 中处理。

Q: 系统响应慢怎么办？

A: 先调大 `config.json` 中的 `timing.action_delay`、`timing.page_timeout`、`timing.voucher_load_wait`。

## 快速步骤

1. 确认 Windows Python 可用：`python --version`。
2. 安装依赖：`python -m pip install pyautogui pyperclip openpyxl keyboard Pillow pygetwindow`。
3. 修改 `config.json` 里的 `excel_path`。
4. 运行 `python collect_positions.py` 采集坐标。
5. 运行 `python main.py`，先选测试 1 验证坐标，再正式跑。
