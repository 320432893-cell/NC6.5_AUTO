import argparse
import subprocess
import sys


BASE_CHECKS = (
    ("json", [sys.executable, "-m", "json.tool", "config.json"]),
    ("config", [sys.executable, "tools/validate_config.py", "config.json"]),
    ("ruff", [".venv/bin/ruff", "check", "."]),
    ("format", [".venv/bin/ruff", "format", "--check", "."]),
    ("compile", [sys.executable, "-m", "compileall", "-q", "core", "tools"]),
    ("basedpyright", [".venv/bin/basedpyright", "."]),
    ("architecture", [sys.executable, "tools/check_architecture.py"]),
    ("pytest", [sys.executable, "-m", "pytest", "-q"]),
)

AUDIT_CHECKS = (
    (
        "semgrep",
        [
            ".venv/bin/semgrep",
            "scan",
            "--config",
            ".semgrep.yml",
            "--error",
            "--severity",
            "ERROR",
            "core",
            "tools",
            "tests",
        ],
    ),
    ("import-linter", [".venv/bin/lint-imports", "--config", ".importlinter"]),
    ("detect-secrets", [sys.executable, "tools/check_detect_secrets.py"]),
    (
        "pip-audit",
        [".venv/bin/pip-audit", "--local", "--progress-spinner", "off"],
    ),
)

DEEP_CHECKS = (
    ("radon", [sys.executable, "-m", "radon", "cc", "core", "tools", "-s", "-a"]),
    (
        "vulture",
        [
            sys.executable, "-m", "vulture",
            "core", "tools", "tests", ".vulture_whitelist.py",
            "--exclude", "*/archive/*",
            "--min-confidence", "60",
        ],
    ),
)

RULE_TOOL_CONTRACT_CHECKS = (
    ("check-tool-contract", [sys.executable, "tools/check.py", "--list"]),
)


def run_check(name, command):
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
        choices=("all", "changed", "audit", "deep", "rule-tool-contracts"),
        help="Check profile. 'changed' is the slice-closure alias for all local gates.",
    )
    parser.add_argument("--list", action="store_true", help="Print configured checks.")
    args = parser.parse_args()

    profiles = {
        "all": BASE_CHECKS + AUDIT_CHECKS,
        "changed": BASE_CHECKS + AUDIT_CHECKS,
        "audit": AUDIT_CHECKS,
        "deep": DEEP_CHECKS,
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
