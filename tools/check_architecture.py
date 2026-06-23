import argparse
import ast
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CORE = ROOT / "core"
TOOLS = ROOT / "tools"

MAX_PROCESSOR_LINES = 320
MAX_WORKFLOW_LINES = 950
WORKFLOW_GLOB = "nc_*_workflow.py"


def check_architecture():
    errors = []
    _check_line_counts(errors)
    _check_workflow_import_boundaries(errors)
    _check_workflow_domain_errors(errors)
    _check_pyautogui_boundary(errors)
    _check_processor_api(errors)
    return errors


def _check_line_counts(errors):
    processor = CORE / "jab_batch_processor.py"
    processor_lines = _line_count(processor)
    if processor_lines > MAX_PROCESSOR_LINES:
        errors.append(
            f"{processor.relative_to(ROOT)} has {processor_lines} lines; "
            f"limit is {MAX_PROCESSOR_LINES}"
        )

    for path in CORE.glob(WORKFLOW_GLOB):
        lines = _line_count(path)
        if lines > MAX_WORKFLOW_LINES:
            errors.append(
                f"{path.relative_to(ROOT)} has {lines} lines; "
                f"limit is {MAX_WORKFLOW_LINES}"
            )


def _check_workflow_import_boundaries(errors):
    workflow_modules = {path.stem for path in CORE.glob(WORKFLOW_GLOB)}
    allowed_cross_imports = set()
    for path in CORE.glob(WORKFLOW_GLOB):
        tree = _parse(path, errors)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if not isinstance(node, ast.ImportFrom):
                continue
            module = node.module or ""
            if not module.startswith("core."):
                continue
            imported = module.removeprefix("core.")
            if imported in workflow_modules and imported not in allowed_cross_imports:
                errors.append(
                    f"{path.relative_to(ROOT)} imports workflow module {module}; "
                    "workflow dependencies should go through the processor wiring"
                )


def _check_workflow_domain_errors(errors):
    for path in CORE.glob(WORKFLOW_GLOB):
        tree = _parse(path, errors)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if not isinstance(node, ast.Raise) or node.exc is None:
                continue
            if _is_runtime_error_raise(node.exc):
                errors.append(
                    f"{path.relative_to(ROOT)}:{node.lineno} raises RuntimeError; "
                    "use core.errors domain exceptions in workflow modules"
                )


def _is_runtime_error_raise(exc):
    if isinstance(exc, ast.Call):
        exc = exc.func
    return isinstance(exc, ast.Name) and exc.id == "RuntimeError"


def _check_pyautogui_boundary(errors):
    allowed = {
        CORE / "jab_operator.py",
        Path(__file__),
    }
    for path in list(CORE.glob("*.py")) + list(TOOLS.glob("*.py")):
        if path in allowed:
            continue
        text = path.read_text(encoding="utf-8")
        if "pyautogui" in text:
            errors.append(
                f"{path.relative_to(ROOT)} references pyautogui; keep GUI input "
                "inside core/jab_operator.py"
            )


def _check_processor_api(errors):
    public_methods = _class_methods(
        CORE / "jab_batch_processor.py", "JABBatchProcessor"
    )
    expected = {
        "backfill_generated_vouchers",
        "close",
        "detect_page_state",
        "detect_voucher_window_state",
        "dry_run",
        "finish_run_state",
        "generate_and_backfill",
        "generate_and_collect_saved",
        "generate_and_save",
        "load_pending_items",
        "match_current_table",
        "normalize_generated_voucher",
        "parse_optional_decimal",
        "record_event",
        "record_transition",
        "require_page_state",
        "resume_current_voucher_window",
        "switch_to_generated_list",
        "wait_for_page_state",
    }
    helper_methods = {
        "choose_main_signature_table",
        "collect_page_controls",
        "collect_window_controls",
        "describe_signature_table",
        "is_generated_signature",
        "looks_loading",
        "read_page_table_signatures",
        "sample_table_col",
        "table_match_ratio",
    }
    unexpected = public_methods - expected - helper_methods
    if unexpected:
        errors.append(
            f"JABBatchProcessor exposes unexpected public methods: {sorted(unexpected)}"
        )


def _class_methods(path, class_name):
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == class_name:
            return {
                child.name
                for child in node.body
                if isinstance(child, ast.FunctionDef) and not child.name.startswith("_")
            }
    return set()


def _parse(path, errors):
    try:
        return ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    except SyntaxError as exc:
        errors.append(f"{path.relative_to(ROOT)} is not valid Python: {exc}")
        return None


def _line_count(path):
    return len(path.read_text(encoding="utf-8").splitlines())


def main():
    parser = argparse.ArgumentParser(description="Check nc_auto_v2 architecture rules")
    parser.parse_args()

    errors = check_architecture()
    if errors:
        print("architecture check failed:")
        for error in errors:
            print(f"- {error}")
        raise SystemExit(1)
    print("architecture check passed")


if __name__ == "__main__":
    main()
