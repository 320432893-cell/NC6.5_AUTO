# 生命周期：T0 兼容壳（删除条件：现场人员全部改用 tools/receipt_detail_entry.py）
# 覆盖的业务阶段：收款单自制录入-明细主行/手续费行现场试写
# 依赖的服务/环境：Windows Python、NC 收款单自制录入界面、Java Access Bridge
# 运行方式：python tools/tmp_receipt_detail_main_line_run.py

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools.receipt_detail_entry import main  # noqa: E402


if __name__ == "__main__":
    print("提示：请改用正式入口 tools/receipt_detail_entry.py；本脚本仅保留短期兼容。")
    raise SystemExit(main())
