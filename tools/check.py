import argparse
import subprocess
import sys


CHECKS = (
    ("json", [sys.executable, "-m", "json.tool", "config.json"]),
    ("config", [sys.executable, "tools/validate_config.py", "config.json"]),
    ("ruff", [".venv/bin/ruff", "check", "."]),
    ("format", [".venv/bin/ruff", "format", "--check", "."]),
    ("compile", [sys.executable, "-m", "compileall", "-q", "core", "tools"]),
    ("basedpyright", [".venv/bin/basedpyright", "."]),
    ("architecture", [sys.executable, "tools/check_architecture.py"]),
    ("pytest", [sys.executable, "-m", "pytest", "-q"]),
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
    parser.parse_args()

    for name, command in CHECKS:
        run_check(name, command)
    print("[check] all checks passed")


if __name__ == "__main__":
    main()
