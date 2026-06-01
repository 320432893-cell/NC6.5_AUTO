from tools import check_architecture


def test_workflow_runtime_error_raise_is_rejected(tmp_path, monkeypatch):
    core = tmp_path / "core"
    core.mkdir()
    workflow = core / "nc_demo_workflow.py"
    workflow.write_text(
        "def run():\n    raise RuntimeError('ambiguous workflow failure')\n",
        encoding="utf-8",
    )

    monkeypatch.setattr(check_architecture, "ROOT", tmp_path)
    monkeypatch.setattr(check_architecture, "CORE", core)

    errors = []
    check_architecture._check_workflow_domain_errors(errors)

    assert errors == [
        "core/nc_demo_workflow.py:2 raises RuntimeError; "
        "use core.errors domain exceptions in workflow modules"
    ]


def test_workflow_domain_error_raise_is_allowed(tmp_path, monkeypatch):
    core = tmp_path / "core"
    core.mkdir()
    workflow = core / "nc_demo_workflow.py"
    workflow.write_text(
        "from core.errors import WorkflowStateError\n\n"
        "def run():\n"
        "    raise WorkflowStateError('bad state')\n",
        encoding="utf-8",
    )

    monkeypatch.setattr(check_architecture, "ROOT", tmp_path)
    monkeypatch.setattr(check_architecture, "CORE", core)

    errors = []
    check_architecture._check_workflow_domain_errors(errors)

    assert errors == []
