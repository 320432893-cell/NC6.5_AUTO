# 职责：现场入口启动前准备 Java Access Bridge 环境
# 不做什么：不读取 NC 控件，不替代 JABOperator 的正式加载/健康检查

from __future__ import annotations

import getpass
import json
import os
from pathlib import Path
import platform
import subprocess
from datetime import datetime

from core.paths import logs_dir


JAB_SETUP_VERSION = 1
REQUIRED_ASSISTIVE_TECHNOLOGY = "com.sun.java.accessibility.AccessBridge"
JAB_SETUP_STATE_NAME = "jab_setup_state.json"


def access_bridge_dll_name() -> str:
    return (
        "WindowsAccessBridge-64.dll"
        if platform.architecture()[0] == "64bit"
        else "WindowsAccessBridge-32.dll"
    )


def uclient_jre_bin_dirs() -> list[Path]:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        return []
    root = Path(local_app_data) / "UClient" / "share"
    if not root.exists():
        return []
    dll_name = access_bridge_dll_name()
    bins = [path for path in root.glob("*/bin") if path.is_dir()]
    return sorted(
        bins,
        key=lambda path: (
            not (path / dll_name).exists(),
            "x64" not in str(path).lower()
            if platform.architecture()[0] == "64bit"
            else "x64" in str(path).lower(),
            str(path).lower(),
        ),
    )


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


def ensure_jab_setup_once() -> dict:
    if os.name != "nt":
        return {
            "ok": False,
            "changed": False,
            "restart_required": False,
            "reason": "Java Access Bridge 必须用 Windows Python 运行。",
        }

    jabswitch = find_jabswitch()
    accessibility_path = accessibility_properties_path()
    context = {
        "version": JAB_SETUP_VERSION,
        "user": current_user(),
        "jabswitch": str(jabswitch) if jabswitch else "",
        "accessibility_path": str(accessibility_path),
    }

    config_ok = accessibility_properties_ok(accessibility_path)
    state = read_jab_setup_state()
    if jabswitch and config_ok and state_matches_context(state, context):
        return {
            "ok": True,
            "changed": False,
            "restart_required": False,
            "skipped": True,
            "reason": "JAB 已完成首次配置。",
            **context,
        }

    result = {
        "ok": False,
        "changed": False,
        "restart_required": True,
        "skipped": False,
        "reason": "",
        "jabswitch_returncode": None,
        "jabswitch_stdout": "",
        "jabswitch_stderr": "",
        "accessibility_changed": False,
        "accessibility_backup": None,
        **context,
    }

    if not jabswitch:
        result["reason"] = (
            "未找到 jabswitch.exe。请确认 NC/UClient 自带 JRE 已安装。"
        )
        return result

    process = subprocess.run(
        [str(jabswitch), "-enable"],
        check=False,
        capture_output=True,
        text=True,
    )
    result["jabswitch_returncode"] = process.returncode
    result["jabswitch_stdout"] = (process.stdout or "").strip()
    result["jabswitch_stderr"] = (process.stderr or "").strip()
    if process.returncode != 0:
        result["reason"] = (
            f"jabswitch -enable 失败，退出码 {process.returncode}。"
            f"{process.stderr or process.stdout}"
        )
        return result

    changed, backup = ensure_accessibility_properties(accessibility_path)
    result["accessibility_changed"] = changed
    result["accessibility_backup"] = str(backup) if backup else None
    result["changed"] = changed or not state_matches_context(state, context)
    result["ok"] = True
    result["reason"] = "JAB 首次配置已确认。"
    write_jab_setup_state({**context, "configured": True, "updated_at": now_text()})
    return result


def find_jabswitch() -> Path | None:
    for directory in candidate_jre_bin_dirs():
        candidate = directory / "jabswitch.exe"
        if candidate.exists():
            return candidate
    return None


def accessibility_properties_path() -> Path:
    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        return Path(userprofile) / ".accessibility.properties"
    home_drive = os.environ.get("HOMEDRIVE", "")
    home_path = os.environ.get("HOMEPATH", "")
    if home_drive and home_path:
        return Path(home_drive + home_path) / ".accessibility.properties"
    return Path.home() / ".accessibility.properties"


def accessibility_properties_ok(path: Path) -> bool:
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return False
    has_bridge = False
    has_magnifier = False
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("assistive_technologies="):
            technologies = [
                item.strip()
                for item in stripped.split("=", 1)[1].split(",")
                if item.strip()
            ]
            has_bridge = REQUIRED_ASSISTIVE_TECHNOLOGY in technologies
        elif stripped == "screen_magnifier_present=true":
            has_magnifier = True
    return has_bridge and has_magnifier


def ensure_accessibility_properties(path: Path) -> tuple[bool, Path | None]:
    existing = ""
    if path.exists():
        existing = path.read_text(encoding="utf-8", errors="replace")

    lines = existing.splitlines()
    found_assistive = False
    found_magnifier = False
    changed = False
    output = []

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("assistive_technologies="):
            found_assistive = True
            key, value = line.split("=", 1)
            technologies = [item.strip() for item in value.split(",") if item.strip()]
            if REQUIRED_ASSISTIVE_TECHNOLOGY not in technologies:
                technologies.append(REQUIRED_ASSISTIVE_TECHNOLOGY)
                changed = True
            output.append(f"{key}={','.join(technologies)}")
        elif stripped.startswith("screen_magnifier_present="):
            found_magnifier = True
            if stripped != "screen_magnifier_present=true":
                changed = True
            output.append("screen_magnifier_present=true")
        else:
            output.append(line)

    if not found_assistive:
        output.append(f"assistive_technologies={REQUIRED_ASSISTIVE_TECHNOLOGY}")
        changed = True
    if not found_magnifier:
        output.append("screen_magnifier_present=true")
        changed = True

    normalized = "\n".join(output).rstrip() + "\n"
    if normalized != existing:
        changed = True
    if not changed:
        return False, None

    path.parent.mkdir(parents=True, exist_ok=True)
    backup = None
    if path.exists():
        backup = path.with_name(
            f"{path.name}.bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        backup.write_text(existing, encoding="utf-8")
    path.write_text(normalized, encoding="utf-8")
    return True, backup


def jab_setup_state_path() -> Path:
    return logs_dir() / JAB_SETUP_STATE_NAME


def read_jab_setup_state() -> dict:
    try:
        with jab_setup_state_path().open("r", encoding="utf-8") as file:
            data = json.load(file)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def write_jab_setup_state(data: dict) -> None:
    path = jab_setup_state_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(f"{path.suffix}.tmp")
    tmp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp_path.replace(path)


def state_matches_context(state: dict, context: dict) -> bool:
    return bool(state.get("configured")) and all(
        str(state.get(key) or "") == str(value or "")
        for key, value in context.items()
    )


def current_user() -> str:
    return os.environ.get("USERNAME") or getpass.getuser()


def now_text() -> str:
    return datetime.now().isoformat(timespec="seconds")


def candidate_jre_bin_dirs() -> list[Path]:
    candidates: list[Path] = []
    candidates.extend(uclient_jre_bin_dirs())
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
    dll_name = access_bridge_dll_name()
    return [str(Path(local_app_data) / "UClient" / "share" / "*" / "bin" / dll_name)]
