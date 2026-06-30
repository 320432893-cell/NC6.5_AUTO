# lifecycle: tool（分层身份漂移闸原型;2026-06-30 advisory,大清扫档）
# 作用：用 fan-in 报晋升候选——tools/ 模块被≥2 个生产模块 import(当库用了,该进 core/)。
#   这是分层倒置(业务逻辑住 tools/)的闸,import-linter 只守"方向"、不报"该晋升"。
# 不做"降级候选":试过——"core 业务模块只被 tools/ 入口调"是正常的(入口→core),
#   硬报全是误报,该不该出 core 是判断、机器报不出。
# 不做什么：不改文件;只铺候选,晋升是带闸的人裁(机器报、人判)。
# 运行：python tools/layer_drift.py

import ast
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def collect():
    mods = {}  # name -> ('core'|'tools', path)
    for d in ("core", "tools"):
        for p in sorted((ROOT / d).rglob("*.py")):
            if "archive/" in str(p) or "/.venv/" in str(p):
                continue
            mods[".".join(p.relative_to(ROOT).with_suffix("").parts)] = (d, p)
    return mods


def importers_of(mods, include_tests=True):
    imp = defaultdict(set)  # module -> set(importing modules)
    srcs = list(mods.items())
    if include_tests:
        for p in sorted((ROOT / "tests").rglob("*.py")):
            srcs.append((f"tests::{p.stem}", ("tests", p)))
    for name, (_d, p) in srcs:
        try:
            tree = ast.parse(p.read_text(encoding="utf-8"))
        except SyntaxError:
            continue
        for n in ast.walk(tree):
            tgt = set()
            if isinstance(n, ast.ImportFrom) and n.module:
                if n.module in mods:
                    tgt.add(n.module)
                for a in n.names:
                    if f"{n.module}.{a.name}" in mods:
                        tgt.add(f"{n.module}.{a.name}")
            elif isinstance(n, ast.Import):
                for a in n.names:
                    if a.name in mods:
                        tgt.add(a.name)
            for t in tgt:
                if t != name:
                    imp[t].add(name)
    return imp


def main():
    mods = collect()
    imp = importers_of(mods)
    promote = []
    for name, (d, _p) in mods.items():
        prod = {c for c in imp.get(name, set()) if not c.startswith("tests::")}
        if d == "tools" and len(prod) >= 2:
            promote.append((name, len(prod)))
    print(f"模块 {len(mods)}")
    print(f"== 晋升候选(tools/ 被当库用,该进 core) {len(promote)} ==")
    for name, n in sorted(promote, key=lambda x: -x[1]):
        print(f"  {name}  被 {n} 个生产模块 import")
    if not promote:
        print("  (无——tools/ 都是薄入口,没业务库)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
