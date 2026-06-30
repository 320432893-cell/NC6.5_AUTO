# lifecycle: tool（import 环检测闸原型;2026-06-30 advisory）
# 作用：建模块级 import 图(含函数内延迟 import),找强连通分量>1=循环依赖簇。
#   import-linter independence 只查指定模块两两独立,查不到全图任意环;本闸补这个洞
#   (如 query_fill↔query_cli 那条延迟 import 的环)。
# 不做什么：不改文件;只报环、不判谁该让步(边界划错=人裁)。
# 运行：python tools/import_cycles.py [core tools]

import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCAN = sys.argv[1:] or ["core", "tools"]
sys.setrecursionlimit(5000)


def modname(p):
    return ".".join(p.relative_to(ROOT).with_suffix("").parts)


def build():
    mods = {}
    for d in SCAN:
        for p in sorted((ROOT / d).rglob("*.py")):
            if "archive/" in str(p) or "/.venv/" in str(p):
                continue
            mods[modname(p)] = p
    edges = {m: set() for m in mods}
    for m, p in mods.items():
        tree = ast.parse(p.read_text(encoding="utf-8"))
        for n in ast.walk(tree):
            if isinstance(n, ast.ImportFrom) and n.module:
                if n.module in mods:
                    edges[m].add(n.module)
                for a in n.names:
                    cand = f"{n.module}.{a.name}"
                    if cand in mods:
                        edges[m].add(cand)
            elif isinstance(n, ast.Import):
                for a in n.names:
                    if a.name in mods:
                        edges[m].add(a.name)
        edges[m].discard(m)
    return mods, edges


def tarjan(mods, edges):
    idx = {}
    low = {}
    on = {}
    stack = []
    counter = [0]
    sccs = []

    def sc(v):
        idx[v] = low[v] = counter[0]
        counter[0] += 1
        stack.append(v)
        on[v] = True
        for w in edges.get(v, ()):
            if w not in idx:
                sc(w)
                low[v] = min(low[v], low[w])
            elif on.get(w):
                low[v] = min(low[v], idx[w])
        if low[v] == idx[v]:
            comp = []
            while True:
                w = stack.pop()
                on[w] = False
                comp.append(w)
                if w == v:
                    break
            if len(comp) > 1:
                sccs.append(comp)

    for v in mods:
        if v not in idx:
            sc(v)
    return sccs


def main():
    mods, edges = build()
    sccs = tarjan(mods, edges)
    print(f"模块 {len(mods)} | 循环依赖簇 {len(sccs)}\n")
    for comp in sccs:
        print("  环:", " ↔ ".join(sorted(comp)))
    if not sccs:
        print("  (无环)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
