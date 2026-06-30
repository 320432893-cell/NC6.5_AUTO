# lifecycle: tool（名字图健康闸原型;2026-06-30 advisory）
# 作用：扫"名字图破口"——让 grep/可达性/重构失准的两类点:
#   ①撞名:同一函数/类名在多文件定义(裸名 grep 命中混、可达性闸合并节点会少报)。
#   ②动态分发断点:getattr(非字面)/importlib/eval/`import X as X` 门面——名字图在此断,
#     静态工具跟丢,需改名或补 `# reach: dispatch/facade` 标记。
# 不做什么：不改文件;近名只做前缀包含的轻提示,可能噪;故 advisory、人核。
# 运行：python tools/name_health.py [core tools]

import ast
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCAN = sys.argv[1:] or ["core", "tools"]
# 每个 CLI/入口自带的样板名,跨文件同名属正常,不算撞名
BOILERPLATE = {"main", "run", "build_parser", "parse_args"}


def iter_py():
    for d in SCAN:
        for p in sorted((ROOT / d).rglob("*.py")):
            if "archive/" in str(p) or "/.venv/" in str(p):
                continue
            yield p


def main():
    defs = defaultdict(list)  # name -> [file:line]
    dispatch = []  # (file:line, 类型, 片段)
    for p in iter_py():
        rel = str(p.relative_to(ROOT))
        try:
            tree = ast.parse(p.read_text(encoding="utf-8"))
        except SyntaxError:
            continue
        # 撞名只看模块级 def/class(类方法同名属正常,mixin 代理/通用方法名不算)
        for n in tree.body:
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                if n.name in BOILERPLATE:
                    continue
                if not (n.name.startswith("__") and n.name.endswith("__")):
                    defs[n.name].append(f"{rel}:{n.lineno}")
        for n in ast.walk(tree):
            # 动态分发断点
            if isinstance(n, ast.Call) and isinstance(n.func, ast.Name):
                if n.func.id in ("getattr", "eval") and len(n.args) >= 2:
                    if n.func.id == "getattr" and isinstance(n.args[1], ast.Constant):
                        pass  # getattr 字面属性=可查,放过
                    else:
                        dispatch.append((f"{rel}:{n.lineno}", n.func.id, ""))
            if isinstance(n, ast.Attribute) and n.attr == "import_module":
                dispatch.append((f"{rel}:{n.lineno}", "importlib", ""))
            if isinstance(n, ast.ImportFrom):
                for a in n.names:
                    if a.asname and a.asname == a.name:  # import X as X 门面
                        dispatch.append((f"{rel}:{n.lineno}", "facade", a.name))

    collisions = {k: v for k, v in defs.items() if len(v) > 1}
    names = sorted(defs)
    nearby = []
    for nm in names:
        for other in names:
            if other != nm and other.startswith(nm + "_"):
                nearby.append((nm, other))

    print(f"扫描 {SCAN} | 符号 {len(defs)}\n")
    print(f"== 撞名(同名多处定义) {len(collisions)} ==")
    for k in sorted(collisions):
        print(f"  {k}: {', '.join(collisions[k])}")
    print(f"\n== 动态分发断点 {len(dispatch)}(改名或补 # reach: 标记) ==")
    for loc, kind, frag in dispatch:
        print(f"  {loc}  [{kind}]{(' ' + frag) if frag else ''}")
    print(f"\n== 近名前缀(轻提示,易混) {len(nearby)} ==")
    for a, b in nearby[:30]:
        print(f"  {a}  ⊂  {b}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
