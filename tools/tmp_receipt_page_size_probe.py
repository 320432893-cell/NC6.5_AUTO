# 职责: T0 在当前收款单查询结果页直接触发正式 set_receipt_page_size，并输出 path/前后页大小报告
# 不做什么: 不打开查询条件，不改查询条件，不保存单据，不写 Excel
# 允许依赖层: core JAB/config、tools.receipt_query_pagination
# 谁不应该 import: 正式流程、core 模块和测试不应 import 本临时探针
# 生命周期: T0 临时探针（删除条件：查询结果页每页500触发机制完成现场复核并沉淀到正式模块/文档）

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_query_pagination import set_receipt_page_size  # noqa: E402


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default="config.json")
    args = parser.parse_args(argv)

    config = load_config(args.config)
    query_cfg = config["receipt_entry"]["query"]
    jab = JABOperator(config)
    try:
        jab.ensure_started()
        report = set_receipt_page_size(jab, query_cfg)
    finally:
        jab.close()
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    return 0 if report.get("page_size_ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
