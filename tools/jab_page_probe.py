import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.nc_page_probe import NCPageProbe  # noqa: E402
from core.utils import load_config  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description="只读枚举 NC/JAB 页面状态特征")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--rows", type=int, default=5)
    parser.add_argument("--cols", type=int, default=25)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        report = NCPageProbe(jab, cfg.get("jab_batch", {})).build_report(
            max_rows=args.rows,
            max_cols=args.cols,
        )
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)


def print_text(report):
    blockers = report["blocking_child_windows"]
    print("blocking_child_windows:", len(blockers))
    for item in blockers:
        print(
            f"  {item['title']!r} {item['class']!r} "
            f"hwnd={item['hwnd']} visible={item['visible']}"
        )

    print("parent_markers:", len(report["parent_markers"]))
    for item in report["parent_markers"][:20]:
        print(
            f"  path={item['path']} name={item['name']!r} role={item['role']!r} "
            f"showing={item['showing']} win={item['window_title']!r}/{item['window_class']!r}"
        )

    print("watched_controls:", len(report["watched_controls"]))
    for item in report["watched_controls"][:80]:
        print(
            f"  path={item['path']} name={item['name']!r} desc={item['description']!r} "
            f"role={item['role']!r} showing={item['showing']} "
            f"win={item['window_title']!r}/{item['window_class']!r}"
        )

    print("table_signatures:", len(report["table_signatures"]))
    for table in report["table_signatures"]:
        print(
            f"  table={table['table_index']} win={table['window_title']!r}/"
            f"{table['window_class']!r} rows={table['row_count']} cols={table['col_count']}"
        )
        print(f"    date_col[{table['date_col']}]: {table['date_values']}")
        print(f"    voucher_col[{table['voucher_col']}]: {table['voucher_values']}")
        for row in table["sample_rows"][:3]:
            print(f"    row {row['row_index']}: {row['cells']}")


if __name__ == "__main__":
    main()
