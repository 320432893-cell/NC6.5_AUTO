# 职责：命名约定的机器兜底(让命名信号不撒谎)——见 .ai-config rules/engineering/naming.index.md
# 不做什么：不改名、不删码;只检查 + 报告。退非0=有硬违规。
# 允许依赖层：标准库;扫 core/tools 的 .py(排除 archive)
# 谁不应该 import：CLI 入口,不被业务 import

import ast
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DIRS = ("core", "tools")
STATUS_PREFIX = ("_deprecated", "_legacy", "_stub", "_wip", "experimental_", "probe_")


def py_files():
    for d in DIRS:
        for f in (ROOT / d).rglob("*.py"):
            if "archive/" not in str(f.relative_to(ROOT)):
                yield f


def check():
    hard = []   # 硬违规:跨模块 import 单下划线私有名
    notes = []  # 提示:状态命名(供 review/grep)
    for f in py_files():
        rel = f.relative_to(ROOT)
        try:
            tree = ast.parse(f.read_text(encoding="utf-8"))
        except SyntaxError:
            continue
        for node in ast.walk(tree):
            # 硬:from core/tools.X import _private(单下划线,非 dunder)
            if isinstance(node, ast.ImportFrom) and node.module and \
                    node.module.split(".")[0] in DIRS:
                for a in node.names:
                    if re.match(r"_[^_]", a.name):
                        hard.append(f"{rel}:{node.lineno}  跨模块 import 私有 `{a.name}`"
                                    f"(来自 {node.module})——私有不应外露")
            # 提示:状态命名的顶层 def/class
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                if node.name.startswith(STATUS_PREFIX):
                    notes.append(f"{rel}:{node.lineno}  {node.name}  (状态命名,grep 可查)")
    return hard, notes


def main():
    hard, notes = check()
    if notes:
        print("[check_naming] 状态命名(提示,非违规):")
        for n in notes:
            print("  " + n)
    if hard:
        print("[check_naming] 硬违规:")
        for h in hard:
            print("  " + h)
        print(f"[check_naming] {len(hard)} 处硬违规 → 红")
        sys.exit(1)
    print("[check_naming] 命名约定通过")


if __name__ == "__main__":
    main()
