# lifecycle: tool（可达性闸原型;2026-06-30 起 advisory,验证能否自动逮语义死码）
# 作用：从真入口(各 main / 模块级代码 / 测试引用)出发,按名字级调用图求可达闭包,
#   报"从入口走不到的函数/方法"。补 vulture 的盲区——vulture 把"任何地方出现该名字"
#   都算用到(死代理引用底层、孤岛互调都被判可达),本闸只认"调用者自身也可达"才算活。
# 不做什么：不删码、不改文件;名字级近似(忽略类区分/动态分发),故 advisory、人核后再删。
# 已知误报源(人核时排除):①re-export 门面 `import X as X`(本闸不把 import 别名计为使用)
#   ②反射/注册表字符串分发的 handler ③只被外部进程/GUI 按名调的入口。
# 已验证(2026-06-30):dunder 修正后逮到 click_control_near_label 等多 agent 审计漏掉的死对。
# 运行：python tools/reachability_probe.py [core tools]

import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCAN_DIRS = sys.argv[1:] or ["core", "tools"]


def refs_in(node):
    """节点子树里所有被引用的名字(Call 的函数名 + 任何 Name/Attribute)——
    含回调按名传递(after_field=self._foo)、装饰器等,贴近 vulture 的"使用"口径。"""
    out = set()
    for n in ast.walk(node):
        if isinstance(n, ast.Name):
            out.add(n.id)
        elif isinstance(n, ast.Attribute):
            out.add(n.attr)
    return out


def collect():
    funcs = {}  # name -> [(file, lineno)]
    edges = {}  # name -> set(被它引用的名字)
    seeds = set()  # 入口直接引用的名字(模块级 + __main__ + 测试)
    for d in SCAN_DIRS:
        for path in sorted((ROOT / d).rglob("*.py")):
            if "archive/" in str(path) or "/.venv/" in str(path):
                continue
            try:
                tree = ast.parse(path.read_text(encoding="utf-8"))
            except SyntaxError:
                continue
            _walk_module(tree, str(path.relative_to(ROOT)), funcs, edges, seeds)
    # 测试目录:整体当入口(测试能触达任何被测函数)
    for path in sorted((ROOT / "tests").rglob("*.py")):
        try:
            tree = ast.parse(path.read_text(encoding="utf-8"))
        except SyntaxError:
            continue
        seeds |= refs_in(tree)
    return funcs, edges, seeds


def _walk_module(tree, relpath, funcs, edges, seeds):
    def register(fn):
        name = fn.name
        funcs.setdefault(name, []).append((relpath, fn.lineno))
        edges.setdefault(name, set())
        edges[name] |= refs_in(fn)
        # 装饰器引用也算入口侧使用
        for dec in fn.decorator_list:
            seeds.update(refs_in(dec))

    def walk_body(body, top):
        for stmt in body:
            if isinstance(stmt, (ast.FunctionDef, ast.AsyncFunctionDef)):
                register(stmt)
            elif isinstance(stmt, ast.ClassDef):
                for sub in stmt.body:
                    if isinstance(sub, (ast.FunctionDef, ast.AsyncFunctionDef)):
                        register(sub)
                    else:
                        seeds.update(refs_in(sub))  # 类体语句(属性赋值等)
            else:
                # 模块级语句(import/赋值/if __name__ 块等)= 入口
                if top:
                    seeds.update(refs_in(stmt))

    walk_body(tree.body, top=True)


def main():
    funcs, edges, seeds = collect()
    # dunder(__init__/__del__/__getattr__…)由运行时隐式调用,不靠名字引用→恒入口
    seeds |= {n for n in funcs if n.startswith("__") and n.endswith("__")}
    reachable = set()
    frontier = set(seeds) & set(funcs)
    reachable |= frontier
    while frontier:
        nxt = set()
        for name in frontier:
            for callee in edges.get(name, ()):
                if callee in funcs and callee not in reachable:
                    nxt.add(callee)
        reachable |= nxt
        frontier = nxt
    dead = sorted(set(funcs) - reachable)
    print(
        f"扫描 {SCAN_DIRS} + tests/  |  函数/方法 {len(funcs)}  入口种子 {len(seeds & set(funcs))}"
    )
    print(f"从入口不可达(候选死码) {len(dead)} 个:\n")
    for name in dead:
        locs = ", ".join(f"{f}:{ln}" for f, ln in funcs[name])
        print(f"  {name}  ({locs})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
