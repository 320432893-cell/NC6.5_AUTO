# 职责：现场入口启动前准备 Java Access Bridge 环境
# 不做什么：不读取 NC 控件，不替代 JABOperator 的正式加载/健康检查

from __future__ import annotations

import os
from pathlib import Path
import platform
import subprocess


def prepare_java_access_bridge() -> dict:
    if os.name != "nt":
        return {
            "ok": False,
            "reason": "Java Access Bridge 必须用 Windows Python 运行。",
        }

    jabswitch = find_jabswitch()
    if not jabswitch:
        return {
            "ok": False,
            "reason": (
                "未找到 jabswitch.exe。请确认 NC/UClient 自带 JRE 已安装，"
                "常见路径为 %LOCALAPPDATA%\\UClient\\share\\java1.7.0_51-x64\\bin。"
            ),
        }

    result = subprocess.run(
        [str(jabswitch), "-enable"],
        check=False,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        return {
            "ok": False,
            "reason": (
                f"jabswitch -enable 失败，退出码 {result.returncode}。"
                f"{result.stderr or result.stdout}"
            ),
            "jabswitch": str(jabswitch),
        }
    return {"ok": True, "jabswitch": str(jabswitch)}


def find_jabswitch() -> Path | None:
    for directory in candidate_jre_bin_dirs():
        candidate = directory / "jabswitch.exe"
        if candidate.exists():
            return candidate
    return None


def candidate_jre_bin_dirs() -> list[Path]:
    candidates: list[Path] = []
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidates.append(
            Path(local_app_data) / "UClient" / "share" / "java1.7.0_51-x64" / "bin"
        )
    java_home = os.environ.get("JAVA_HOME")
    if java_home:
        candidates.append(Path(java_home) / "bin")
    for item in os.environ.get("PATH", "").split(os.pathsep):
        if item:
            candidates.append(Path(item))
    return candidates


def uclient_access_bridge_dll_patterns() -> list[str]:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        return []
    dll_name = (
        "WindowsAccessBridge-64.dll"
        if platform.architecture()[0] == "64bit"
        else "WindowsAccessBridge-32.dll"
    )
    return [str(Path(local_app_data) / "UClient" / "share" / "*" / "bin" / dll_name)]
