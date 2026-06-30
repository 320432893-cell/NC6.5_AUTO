import argparse
import importlib.util
import subprocess
import sys
from pathlib import Path


# 三档按「速度 × 改起来影响半径」分:fast 每次写完、stage 阶段闭包、cleanup 大扫除。
FAST_CHECKS = (  # 秒级·每次写完都跑
    ("json", [sys.executable, "-m", "json.tool", "config.json"]),
    ("config", [sys.executable, "tools/validate_config.py", "config.json"]),
    ("ruff", [".venv/bin/ruff", "check", "."]),
    ("format", [".venv/bin/ruff", "format", "--check", "."]),
    ("compile", [sys.executable, "-m", "compileall", "-q", "core", "tools"]),
)

STAGE_CHECKS = (  # 地基/高影响·快·阶段闭包跑(改了后续影响大,早逮)
    ("import-linter", [".venv/bin/lint-imports", "--config", ".importlinter"]),
    ("import-cycles", [sys.executable, "tools/import_cycles.py"]),
    ("architecture", [sys.executable, "tools/check_architecture.py"]),
    ("naming", [sys.executable, "tools/check_naming.py"]),
    ("detect-secrets", [sys.executable, "tools/check_detect_secrets.py"]),
)

CLEANUP_CHECKS = (  # 全量/慢/涌现·大扫除才跑(攒一批才显形)
    (
        "vulture",
        [
            sys.executable, "-m", "vulture",
            "core", "tools", "tests", ".vulture_whitelist.py",
            "--exclude", "*/archive/*",
            "--min-confidence", "60",
        ],
    ),
    ("reachability", [sys.executable, "tools/reachability_probe.py"]),
    ("name-health", [sys.executable, "tools/name_health.py"]),
    ("layer-drift", [sys.executable, "tools/layer_drift.py"]),
    ("radon", [sys.executable, "-m", "radon", "cc", "core", "tools", "-s", "-a"]),
    ("basedpyright", [".venv/bin/basedpyright", "."]),
    (
        "semgrep",
        [
            ".venv/bin/semgrep", "scan", "--config", ".semgrep.yml",
            "--error", "--severity", "ERROR", "core", "tools", "tests",
        ],
    ),
    ("pip-audit", [".venv/bin/pip-audit", "--local", "--progress-spinner", "off"]),
    ("pytest", [sys.executable, "-m", "pytest", "-q"]),
)

RULE_TOOL_CONTRACT_CHECKS = (
    ("check-tool-contract", [sys.executable, "tools/check.py", "--list"]),
)


def gate_runnable(command):
    """这闸的二进制装了没(能不能跑)——只判可运行,不判绿不绿。返回 (ok, 说明)。"""
    head = command[0]
    if head == sys.executable:
        if len(command) >= 3 and command[1] == "-m":
            return importlib.util.find_spec(command[2]) is not None, f"-m {command[2]}"
        if len(command) >= 2 and command[1].endswith(".py"):
            return Path(command[1]).exists(), command[1]
        return True, "python"
    return Path(head).exists(), head


def doctor():
    """闸健康:遍历所有闸,配了却没装的=假闸(摆设),红。"""
    dead = []
    print("[doctor] 闸健康(能不能跑,不判绿)")
    for name, command in FAST_CHECKS + STAGE_CHECKS + CLEANUP_CHECKS:
        ok, detail = gate_runnable(command)
        print(f"  {'OK' if ok else '假闸·DEAD':<11} {name}  [{detail}]")
        if not ok:
            dead.append(name)
    if dead:
        print(f"\n假闸 {len(dead)}(配在 check.py 却没装,等于摆设):{', '.join(dead)}")
        raise SystemExit(1)
    print("\n全部闸可运行")


def run_check(name, command):
    ok, _ = gate_runnable(command)
    if not ok:
        print(f"[skip] {name} 假闸·没装,跳过(跑 check.py --doctor 看全部)")
        return
    print(f"[check] {name}")
    result = subprocess.run(command, check=False, capture_output=True, text=True)
    if result.returncode != 0:
        if result.stdout:
            print(result.stdout, end="")
        if result.stderr:
            print(result.stderr, end="", file=sys.stderr)
        raise SystemExit(result.returncode)


def main():
    parser = argparse.ArgumentParser(description="Run project quality checks")
    parser.add_argument(
        "profile",
        nargs="?",
        default="all",
        choices=(
            "fast", "stage", "cleanup",
            "all", "changed", "deep", "rule-tool-contracts",
        ),
        help="fast 每次写完 / stage 阶段闭包 / cleanup 大扫除(全量)。",
    )
    parser.add_argument("--list", action="store_true", help="Print configured checks.")
    parser.add_argument(
        "--doctor", action="store_true", help="闸健康:验每个闸的二进制装没装(假闸=红)。"
    )
    args = parser.parse_args()

    if args.doctor:
        doctor()
        return

    stage = FAST_CHECKS + STAGE_CHECKS
    cleanup = stage + CLEANUP_CHECKS
    profiles = {
        "fast": FAST_CHECKS,
        "stage": stage,
        "cleanup": cleanup,
        "changed": stage,  # 切片/阶段闭包别名
        "deep": cleanup,  # 大扫除别名
        "all": cleanup,
        "rule-tool-contracts": RULE_TOOL_CONTRACT_CHECKS,
    }
    checks = profiles[args.profile]
    if args.list:
        for name, command in checks:
            print(f"{name}: {' '.join(command)}")
        return

    for name, command in checks:
        run_check(name, command)
    print("[check] all checks passed")


if __name__ == "__main__":
    main()
