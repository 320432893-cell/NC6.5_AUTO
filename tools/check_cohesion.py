# 职责：拆分判据——算模块内函数调用图的连通分量,判"该不该拆"(行数是信号、连通度是裁判)
# 不做什么：不拆、不改码;只分析 + 报告。见 .ai-config rules/engineering/code.index.md §4
# 允许依赖层：标准库;读单个 .py
# 谁不应该 import：CLI 入口,不被业务 import

import ast
import collections
import sys
from pathlib import Path

BLOB_RATIO = 0.80  # 最大连通分量占比 ≥ 此值 = 内聚单块,判不拆


def components(path):
    src = Path(path).read_text(encoding="utf-8")
    tree = ast.parse(src)
    funcs = {
        n.name: n
        for n in tree.body
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
    }
    names = set(funcs)
    adj = collections.defaultdict(set)
    for nm, node in funcs.items():
        for x in ast.walk(node):
            if (
                isinstance(x, ast.Name)
                and isinstance(x.ctx, ast.Load)
                and x.id in names
                and x.id != nm
            ):
                adj[nm].add(x.id)
                adj[x.id].add(nm)
    seen, comps = set(), []
    for n in names:
        if n in seen:
            continue
        stack, comp = [n], set()
        while stack:
            x = stack.pop()
            if x in seen:
                continue
            seen.add(x)
            comp.add(x)
            stack += list(adj[x] - seen)
        comps.append(comp)
    comps.sort(key=len, reverse=True)
    return names, comps


def main():
    if len(sys.argv) < 2:
        print("用法: check_cohesion.py <模块.py> [模块.py ...]")
        sys.exit(2)
    for path in sys.argv[1:]:
        names, comps = components(path)
        if not names:
            print(f"{path}: 无顶层函数")
            continue
        biggest = len(comps[0])
        ratio = biggest / len(names)
        verdict = (
            "不拆(内聚单块,硬拆=制造循环依赖)"
            if ratio >= BLOB_RATIO
            else f"可拆({len(comps)} 个子簇,按簇切+import-linter锁)"
        )
        print(
            f"{path}: {len(names)} 函数 → {len(comps)} 连通分量, "
            f"最大块 {biggest}({ratio:.0%}) → {verdict}"
        )
        for c in comps[:6]:
            print(
                f"    [{len(c)}] {', '.join(sorted(c)[:5])}{'...' if len(c) > 5 else ''}"
            )


if __name__ == "__main__":
    main()
