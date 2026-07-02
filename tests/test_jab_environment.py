import json
from pathlib import Path

from core import jab_environment


def test_ensure_jab_setup_once_skips_when_state_and_config_match(tmp_path, monkeypatch):
    jre_bin = tmp_path / "UClient" / "share" / "java1.7.0_51-x64" / "bin"
    jre_bin.mkdir(parents=True)
    jabswitch = jre_bin / "jabswitch.exe"
    jabswitch.write_text("", encoding="utf-8")
    accessibility = tmp_path / "user" / ".accessibility.properties"
    accessibility.parent.mkdir()
    accessibility.write_text(
        "assistive_technologies=com.sun.java.accessibility.AccessBridge\n"
        "screen_magnifier_present=true\n",
        encoding="utf-8",
    )
    state_dir = tmp_path / "runtime"
    state_dir.mkdir()
    (state_dir / "jab_setup_state.json").write_text(
        json.dumps(
            {
                "version": jab_environment.JAB_SETUP_VERSION,
                "user": "tester",
                "jabswitch": str(jabswitch),
                "accessibility_path": str(accessibility),
                "configured": True,
            }
        ),
        encoding="utf-8",
    )
    calls = []

    monkeypatch.setattr(jab_environment.os, "name", "nt")
    monkeypatch.setattr(jab_environment, "find_jabswitch", lambda: jabswitch)
    monkeypatch.setattr(
        jab_environment, "accessibility_properties_path", lambda: accessibility
    )
    monkeypatch.setattr(jab_environment, "current_user", lambda: "tester")
    monkeypatch.setattr(jab_environment, "logs_dir", lambda: state_dir)
    monkeypatch.setattr(
        jab_environment.subprocess,
        "run",
        lambda *args, **kwargs: calls.append(args) or None,
    )

    result = jab_environment.ensure_jab_setup_once()

    assert result["ok"] is True
    assert result["skipped"] is True
    assert result["changed"] is False
    assert calls == []


def test_ensure_jab_setup_once_repairs_config_and_writes_state(tmp_path, monkeypatch):
    jabswitch = tmp_path / "jre" / "bin" / "jabswitch.exe"
    jabswitch.parent.mkdir(parents=True)
    jabswitch.write_text("", encoding="utf-8")
    accessibility = tmp_path / "user" / ".accessibility.properties"
    accessibility.parent.mkdir()
    accessibility.write_text("screen_magnifier_present=false\n", encoding="utf-8")
    state_dir = tmp_path / "runtime"
    state_dir.mkdir()

    class Result:
        returncode = 0
        stdout = "enabled"
        stderr = ""

    monkeypatch.setattr(jab_environment.os, "name", "nt")
    monkeypatch.setattr(jab_environment, "find_jabswitch", lambda: jabswitch)
    monkeypatch.setattr(
        jab_environment, "accessibility_properties_path", lambda: accessibility
    )
    monkeypatch.setattr(jab_environment, "current_user", lambda: "tester")
    monkeypatch.setattr(jab_environment, "logs_dir", lambda: state_dir)
    monkeypatch.setattr(jab_environment.subprocess, "run", lambda *args, **kwargs: Result())

    result = jab_environment.ensure_jab_setup_once()

    assert result["ok"] is True
    assert result["skipped"] is False
    assert result["accessibility_changed"] is True
    text = accessibility.read_text(encoding="utf-8")
    assert "assistive_technologies=com.sun.java.accessibility.AccessBridge" in text
    assert "screen_magnifier_present=true" in text
    state = json.loads((state_dir / "jab_setup_state.json").read_text(encoding="utf-8"))
    assert state["configured"] is True
    assert state["jabswitch"] == str(jabswitch)

