# 职责：运行 detect-secrets 扫描并把发现项转换为闭包门禁失败
# 不做什么：不生成或更新 secrets baseline，不自动清理疑似密钥
# 允许依赖层：标准库和 detect-secrets CLI 模块
# 谁不应该 import：业务代码和 NC/JAB 自动化流程不应 import

import json
import subprocess
import sys


EXCLUDE_FILES = r"(^\.git/|^\.venv/|^\.ruff_cache/|^\.pytest_cache/|^\.import_linter_cache/|^logs/|__pycache__)"


def main():
    command = [
        sys.executable,
        "-m",
        "detect_secrets",
        "scan",
        "--all-files",
        "--exclude-files",
        EXCLUDE_FILES,
    ]
    result = subprocess.run(command, check=False, capture_output=True, text=True)
    if result.returncode != 0:
        if result.stdout:
            print(result.stdout, end="")
        if result.stderr:
            print(result.stderr, end="", file=sys.stderr)
        raise SystemExit(result.returncode)
    payload = json.loads(result.stdout or "{}")
    findings = payload.get("results") or {}
    if findings:
        print("detect-secrets found possible secrets:")
        for path, items in sorted(findings.items()):
            lines = ", ".join(str(item.get("line_number")) for item in items)
            print(f"- {path}: lines {lines}")
        raise SystemExit(1)
    print("detect-secrets scan passed")


if __name__ == "__main__":
    main()
