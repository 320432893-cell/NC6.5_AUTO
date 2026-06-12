import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.archive.receipt_text_controls_dump import collect_text_controls  # noqa: E402


def main():
    parser = argparse.ArgumentParser(
        description="Dump current text controls with a temporary JAB depth."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--depth", type=int, default=50)
    parser.add_argument("--query", required=True)
    parser.add_argument("--class-name", default="SunAwtCanvas")
    parser.add_argument("--all-controls", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        jab.max_depth = args.depth
        jab.ensure_started()
        rows = collect_text_controls(
            jab, args.class_name, args.query, args.all_controls
        )
    finally:
        jab.close()

    print(json.dumps(rows, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
