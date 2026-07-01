# Read-only monitor for NC receipt engine runtime state.
#
# It does not import or call JAB, does not send keys, and does not touch NC UI.
# It only tails runtime logs/state files and optionally snapshots Windows tasklist.

from __future__ import annotations

import argparse
import csv
import io
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.paths import runtime_dir as engine_runtime_dir  # noqa: E402


DEFAULT_WSL_GUI_RUNTIME = Path(
    "/mnt/h/python脚本/采购对账桌面版/data/nc_runtime"
)
DEFAULT_WINDOWS_GUI_RUNTIME = Path(
    r"H:\python脚本\采购对账桌面版\data\nc_runtime"
)

JAB_LOAD_PATTERN = re.compile(r"JAB 已加载|JAB loaded|WindowsAccessBridge", re.I)
JAB_LOAD_LINE_PATTERN = re.compile(
    r".*(?:JAB 已加载|JAB loaded|WindowsAccessBridge).*",
    re.I,
)
FATAL_PATTERNS = (
    re.compile(r"Traceback \(most recent call last\)"),
    re.compile(r"Windows fatal exception", re.I),
    re.compile(r"access violation", re.I),
    re.compile(r"segmentation fault", re.I),
    re.compile(r"initializeAccessBridge returned false", re.I),
    re.compile(r"Fatal Python error", re.I),
    re.compile(r"崩溃|闪退|致命"),
)
THREAD_ALIVE_PATTERN = re.compile(r'"thread_alive"\s*:\s*true', re.I)
ASYNC_TIMEOUT_PATTERN = re.compile(
    r"action_returned_within_timeout[^,\n\r]*false|did not return within|async.*超时",
    re.I,
)
ASYNC_TIMEOUT_LINE_PATTERN = re.compile(
    r".*(?:action_returned_within_timeout[^,\n\r]*false|did not return within|async.*超时).*",
    re.I,
)
ACTION_LINE_PATTERN = re.compile(
    r".*(?:JAB 执行动作|do_action|action_returned_within_timeout|failure=).*",
    re.I,
)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Read-only monitor for NC receipt runtime logs. "
            "Safe to run while the GUI engine is operating."
        )
    )
    parser.add_argument(
        "--runtime-dir",
        default=None,
        help="Runtime dir containing logs/. Defaults to NC_RUNTIME_DIR or GUI runtime.",
    )
    parser.add_argument(
        "--logs-dir",
        default=None,
        help="Direct logs dir override. Usually <runtime-dir>/logs.",
    )
    parser.add_argument("--interval", type=float, default=1.0)
    parser.add_argument(
        "--duration",
        type=float,
        default=0.0,
        help="Seconds to run. 0 means run until Ctrl+C.",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="Optional JSONL output path for snapshots.",
    )
    parser.add_argument(
        "--no-tasklist",
        action="store_true",
        help="Do not snapshot Windows tasklist.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print unchanged heartbeat snapshots too.",
    )
    parser.add_argument(
        "--stress",
        action="store_true",
        help="Detailed long-run diagnostics: durations, stalls, memory, and summary.",
    )
    parser.add_argument(
        "--stall-seconds",
        type=float,
        default=20.0,
        help="In --stress mode, warn when run_state has not changed for this many seconds.",
    )
    return parser.parse_args(argv)


def resolve_logs_dir(args):
    if args.logs_dir:
        return Path(args.logs_dir)
    if args.runtime_dir:
        return Path(args.runtime_dir) / "logs"
    env_runtime = os.environ.get("NC_RUNTIME_DIR")
    if env_runtime:
        return Path(env_runtime) / "logs"
    if DEFAULT_WINDOWS_GUI_RUNTIME.exists():
        return DEFAULT_WINDOWS_GUI_RUNTIME / "logs"
    if DEFAULT_WSL_GUI_RUNTIME.exists():
        return DEFAULT_WSL_GUI_RUNTIME / "logs"
    return engine_runtime_dir() / "logs"


def read_json(path):
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        return None
    except json.JSONDecodeError as exc:
        return {"_read_error": f"JSONDecodeError:{exc}"}
    except OSError as exc:
        return {"_read_error": f"OSError:{exc}"}


def safe_file_size(path):
    try:
        return path.stat().st_size
    except FileNotFoundError:
        return 0
    except OSError:
        return 0


def read_text_from_offset(path, offset, max_bytes=512_000):
    try:
        size = path.stat().st_size
        if offset > size:
            offset = 0
        read_start = offset
        skipped_bytes = 0
        if size - read_start > max_bytes:
            skipped_bytes = size - read_start - max_bytes
            read_start = size - max_bytes
        with path.open("rb") as fh:
            fh.seek(read_start)
            data = fh.read()
        return (
            data.decode("utf-8", errors="replace"),
            size,
            read_start,
            size,
            skipped_bytes,
        )
    except FileNotFoundError:
        return "", 0, offset, offset, 0
    except OSError as exc:
        return f"<read-error {type(exc).__name__}: {exc}>", 0, offset, offset, 0


def windows_tasklist_snapshot():
    cmd = ["cmd.exe", "/c", "tasklist /FO CSV /NH"]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="mbcs" if os.name == "nt" else None,
            errors="replace",
            timeout=3,
            check=False,
        )
    except Exception as exc:
        return {"ok": False, "reason": f"{type(exc).__name__}: {exc}"}
    if proc.returncode != 0:
        return {"ok": False, "reason": proc.stderr.strip() or proc.stdout.strip()}
    rows = []
    for row in csv.reader(io.StringIO(proc.stdout)):
        if len(row) < 5:
            continue
        image, pid, session_name, session_no, mem = row[:5]
        image_l = image.lower()
        if not any(
            marker in image_l
            for marker in ("java", "uclient", "python", "采购对账", "nc")
        ):
            continue
        rows.append(
            {
                "image": image,
                "pid": pid,
                "session": session_name,
                "mem": mem,
                "mem_kb": parse_tasklist_mem_kb(mem),
            }
        )
    java_items = [item for item in rows if "java" in item["image"].lower()]
    python_items = [item for item in rows if "python" in item["image"].lower()]
    return {
        "ok": True,
        "count": len(rows),
        "items": rows[:80],
        "java_count": len(java_items),
        "python_count": len(python_items),
        "java_pids": [item["pid"] for item in java_items],
        "python_pids": [item["pid"] for item in python_items],
        "java_mem_kb": sum(item.get("mem_kb") or 0 for item in java_items),
        "python_mem_kb": sum(item.get("mem_kb") or 0 for item in python_items),
    }


def parse_tasklist_mem_kb(text):
    digits = re.sub(r"[^\d]", "", str(text or ""))
    return int(digits) if digits else 0


def summarize_run_state(state):
    if not isinstance(state, dict):
        return {"present": False}
    return {
        "present": True,
        "run_id": state.get("run_id") or "",
        "status": state.get("status") or "",
        "stage": state.get("stage") or "",
        "step_index": state.get("step_index"),
        "total_steps": state.get("total_steps"),
        "current": state.get("current") or {},
        "counts": state.get("counts") or {},
        "last_error": state.get("last_error") or state.get("error") or "",
        "event_count": len(state.get("events") or []),
        "last_event": (state.get("events") or [{}])[-1],
        "events": state.get("events") or [],
        "read_error": state.get("_read_error") or "",
    }


def compact_snippet(text, limit=220):
    text = " ".join(str(text or "").replace("\r", "").split())
    if len(text) <= limit:
        return text
    return text[: limit - 3] + "..."


def format_mem(kb):
    try:
        kb = int(kb or 0)
    except (TypeError, ValueError):
        kb = 0
    if kb <= 0:
        return "0MB"
    return f"{kb / 1024:.1f}MB"


def regex_line_hits(pattern, text, limit=8):
    hits = []
    for line in str(text or "").splitlines():
        if pattern.search(line):
            hits.append(compact_snippet(line))
            if len(hits) >= limit:
                break
    return hits


def detect_stdout_signals(text):
    fatal_hits = []
    for pattern in FATAL_PATTERNS:
        for match in pattern.finditer(text):
            start = max(match.start() - 120, 0)
            end = min(match.end() + 200, len(text))
            fatal_hits.append(text[start:end].replace("\r", "")[:500])
            if len(fatal_hits) >= 5:
                break
        if len(fatal_hits) >= 5:
            break
    return {
        "jab_load_count": len(JAB_LOAD_PATTERN.findall(text)),
        "jab_load_lines": regex_line_hits(JAB_LOAD_LINE_PATTERN, text),
        "fatal_hit_count": len(fatal_hits),
        "fatal_hits": fatal_hits,
        "thread_alive_true_count": len(THREAD_ALIVE_PATTERN.findall(text)),
        "async_timeout_hint_count": len(ASYNC_TIMEOUT_PATTERN.findall(text)),
        "async_timeout_lines": regex_line_hits(ASYNC_TIMEOUT_LINE_PATTERN, text),
        "action_lines": regex_line_hits(ACTION_LINE_PATTERN, text),
    }


SIGNAL_KEYS = (
    "jab_load_count",
    "fatal_hit_count",
    "thread_alive_true_count",
    "async_timeout_hint_count",
)


def make_snapshot(logs_dir, include_tasklist=True, stdout_offset=0):
    run_state_path = logs_dir / "run_state.json"
    stdout_path = logs_dir / "last_engine_stdout.txt"
    state = read_json(run_state_path)
    stdout_text, stdout_size, read_start, next_offset, skipped = read_text_from_offset(
        stdout_path, stdout_offset
    )
    snapshot = {
        "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
        "logs_dir": str(logs_dir),
        "run_state": summarize_run_state(state),
        "stdout": {
            "path": str(stdout_path),
            "size": stdout_size,
            "read_start": read_start,
            "next_offset": next_offset,
            "new_bytes": max(stdout_size - stdout_offset, 0),
            "skipped_bytes": skipped,
            **detect_stdout_signals(stdout_text),
        },
    }
    if include_tasklist:
        snapshot["tasklist"] = windows_tasklist_snapshot()
    return snapshot


def event_identity(event):
    event = event or {}
    return (
        event.get("ts"),
        event.get("name"),
        event.get("excel_row"),
        event.get("stage"),
        event.get("ok"),
        event.get("outcome"),
        event.get("error") or event.get("reason"),
    )


def new_run_events(snapshot, previous, limit=6):
    events = (snapshot.get("run_state") or {}).get("events") or []
    if not events or not previous:
        return []
    previous_events = (previous.get("run_state") or {}).get("events") or []
    previous_ids = {event_identity(item) for item in previous_events}
    added = [item for item in events if event_identity(item) not in previous_ids]
    return added[-limit:]


def format_event(event):
    parts = [str(event.get("name") or "event")]
    if event.get("excel_row") is not None:
        parts.append(f"row={event.get('excel_row')}")
    if event.get("ok") is not None:
        parts.append(f"ok={event.get('ok')}")
    if event.get("outcome"):
        parts.append(f"outcome={event.get('outcome')}")
    reason = event.get("error") or event.get("reason")
    if reason:
        parts.append(f"reason={compact_snippet(reason, 100)}")
    return " ".join(parts)


def process_delta(current_tasklist, previous_tasklist):
    if not current_tasklist.get("ok") or not previous_tasklist.get("ok"):
        return {}
    delta = {}
    for key in ("java_pids", "python_pids"):
        current = set(current_tasklist.get(key) or [])
        previous = set(previous_tasklist.get(key) or [])
        added = sorted(current - previous)
        removed = sorted(previous - current)
        if added or removed:
            delta[key] = {"added": added, "removed": removed}
    return delta


def add_detail_lines(lines, snapshot, previous):
    stdout = snapshot.get("stdout") or {}
    tasklist = snapshot.get("tasklist") or {}
    prev_tasklist = (previous or {}).get("tasklist") or {}
    for event in new_run_events(snapshot, previous):
        lines.append(f"    event: {format_event(event)}")
    proc_delta = process_delta(tasklist, prev_tasklist)
    for key, delta in proc_delta.items():
        label = "javaPid" if key == "java_pids" else "pythonPid"
        lines.append(
            f"    {label}: +{delta.get('added') or []} -{delta.get('removed') or []}"
        )
    for line in stdout.get("jab_load_lines") or []:
        lines.append(f"    jab: {line}")
    for line in stdout.get("async_timeout_lines") or []:
        lines.append(f"    async: {line}")
    for line in stdout.get("action_lines") or []:
        if line in (stdout.get("jab_load_lines") or []):
            continue
        lines.append(f"    action: {line}")
    for hit in stdout.get("fatal_hits") or []:
        lines.append(f"    fatal: {compact_snippet(hit)}")


def state_signature(snapshot):
    state = snapshot.get("run_state") or {}
    active_error = state.get("last_error") if state.get("status") == "running" else ""
    return (
        state.get("run_id"),
        state.get("status"),
        state.get("stage"),
        state.get("step_index"),
        state.get("total_steps"),
        (state.get("current") or {}).get("excel_row"),
        state.get("event_count"),
        active_error,
    )


class StressTracker:
    def __init__(self, stall_seconds=20.0):
        self.stall_seconds = float(stall_seconds or 20.0)
        self.first_seen = time.monotonic()
        self.last_change_at = self.first_seen
        self.last_signature = None
        self.run_started_at = {}
        self.row_started_at = {}
        self.stage_started_at = {}
        self.row_outcomes = {}
        self.stage_max_seconds = {}
        self.signal_totals = dict.fromkeys(SIGNAL_KEYS, 0)
        self.max_java_mem_kb = 0
        self.max_python_mem_kb = 0
        self.java_pid_changes = 0
        self.python_pid_changes = 0
        self.last_tasklist = {}

    def observe(self, snapshot, previous):
        now = time.monotonic()
        lines = []
        signature = state_signature(snapshot)
        if signature != self.last_signature:
            self.last_change_at = now
            self.last_signature = signature
        state = snapshot.get("run_state") or {}
        run_id = state.get("run_id") or ""
        row = (state.get("current") or {}).get("excel_row")
        stage = state.get("stage") or ""
        if run_id and run_id not in self.run_started_at:
            self.run_started_at[run_id] = now
            lines.append(f"    stress: run-start run={run_id}")
        if run_id and row is not None:
            row_key = (run_id, row)
            if row_key not in self.row_started_at:
                self.row_started_at[row_key] = now
                lines.append(f"    stress: row-start row={row}")
            stage_key = (run_id, row, stage)
            if stage and stage_key not in self.stage_started_at:
                self.stage_started_at[stage_key] = now
        for event in new_run_events(snapshot, previous, limit=20):
            self.observe_event(lines, run_id, event, now)
        self.observe_tasklist(lines, snapshot.get("tasklist") or {}, previous)
        stdout = snapshot.get("stdout") or {}
        for key in SIGNAL_KEYS:
            self.signal_totals[key] += int(stdout.get(key) or 0)
        stalled = now - self.last_change_at
        if run_id and state.get("status") == "running" and stalled >= self.stall_seconds:
            lines.append(
                "    warn: run_state-no-change "
                f"{stalled:.1f}s stage={stage or '-'} row={row or '-'}"
            )
        return lines

    def observe_event(self, lines, run_id, event, now):
        name = str(event.get("name") or "")
        row = event.get("excel_row")
        if row is not None and name == "row-done":
            key = (run_id, row)
            elapsed = None
            if key in self.row_started_at:
                elapsed = now - self.row_started_at[key]
            outcome = event.get("outcome") or ""
            self.row_outcomes[row] = outcome
            elapsed_text = f" seconds={elapsed:.1f}" if elapsed is not None else ""
            lines.append(f"    stress: row-done row={row} outcome={outcome}{elapsed_text}")
        if name in {"header-verify-summary", "post-query-done"}:
            ok = event.get("ok")
            lines.append(f"    stress: checkpoint {name} row={row or '-'} ok={ok}")
        if name.endswith("failed") or event.get("outcome") == "failed":
            lines.append(f"    warn: event-failure {format_event(event)}")

    def observe_tasklist(self, lines, tasklist, previous):
        if not tasklist.get("ok"):
            return
        self.max_java_mem_kb = max(
            self.max_java_mem_kb, int(tasklist.get("java_mem_kb") or 0)
        )
        self.max_python_mem_kb = max(
            self.max_python_mem_kb, int(tasklist.get("python_mem_kb") or 0)
        )
        prev_tasklist = (previous or {}).get("tasklist") or self.last_tasklist or {}
        delta = process_delta(tasklist, prev_tasklist)
        if "java_pids" in delta:
            self.java_pid_changes += 1
        if "python_pids" in delta:
            self.python_pid_changes += 1
        if tasklist.get("java_mem_kb"):
            lines.append(
                "    mem: "
                f"java={format_mem(tasklist.get('java_mem_kb'))} "
                f"python={format_mem(tasklist.get('python_mem_kb'))}"
            )
        self.last_tasklist = tasklist

    def summary_lines(self):
        counts = {}
        for outcome in self.row_outcomes.values():
            counts[outcome or "-"] = counts.get(outcome or "-", 0) + 1
        return [
            "[monitor-summary] "
            f"rows={counts} "
            f"signals={self.signal_totals} "
            f"java_pid_changes={self.java_pid_changes} "
            f"python_pid_changes={self.python_pid_changes} "
            f"max_java_mem={format_mem(self.max_java_mem_kb)} "
            f"max_python_mem={format_mem(self.max_python_mem_kb)}"
        ]


def compact_line(snapshot, previous=None, verbose=False, stress_tracker=None):
    state = snapshot["run_state"]
    stdout = snapshot["stdout"]
    totals = snapshot.get("stdout_since_start") or {}
    tasklist = snapshot.get("tasklist") or {}
    changed = verbose
    fields = [
        snapshot["ts"],
        f"run={state.get('run_id') or '-'}",
        f"status={state.get('status') or '-'}",
        f"stage={state.get('stage') or '-'}",
        f"step={state.get('step_index')}/{state.get('total_steps')}",
        f"row={(state.get('current') or {}).get('excel_row', '-')}",
        f"newBytes={stdout.get('new_bytes')}",
        f"JAB+={totals.get('jab_load_count', 0)}",
        f"fatal+={totals.get('fatal_hit_count', 0)}",
        f"threadAlive+={totals.get('thread_alive_true_count', 0)}",
        f"asyncHints+={totals.get('async_timeout_hint_count', 0)}",
    ]
    if tasklist.get("ok"):
        fields.append(f"java={tasklist.get('java_count')}")
        fields.append(f"python={tasklist.get('python_count')}")
        if tasklist.get("java_pids"):
            fields.append(f"javaPid={','.join(tasklist.get('java_pids') or [])}")
        if stress_tracker:
            fields.append(f"javaMem={format_mem(tasklist.get('java_mem_kb'))}")
    if state.get("last_error") and state.get("status") == "running":
        fields.append(f"err={state.get('last_error')[:120]}")
    elif state.get("last_error"):
        fields.append(f"lastErr={state.get('last_error')[:120]}")
    if previous:
        prev_state = previous["run_state"]
        prev_stdout = previous["stdout"]
        for key in ("run_id", "status", "stage", "step_index"):
            if state.get(key) != prev_state.get(key):
                changed = True
        if (
            state.get("status") == "running"
            and state.get("last_error") != prev_state.get("last_error")
        ):
            changed = True
        if stdout.get("new_bytes"):
            changed = True
        for key in SIGNAL_KEYS:
            if stdout.get(key):
                changed = True
        if stdout.get("size") != prev_stdout.get("size"):
            changed = True
    else:
        changed = True
    if not changed:
        return None
    lines = [" | ".join(str(item) for item in fields)]
    add_detail_lines(lines, snapshot, previous)
    if stress_tracker:
        lines.extend(stress_tracker.observe(snapshot, previous))
    return "\n".join(lines)


def main(argv=None):
    args = parse_args(argv)
    logs_dir = resolve_logs_dir(args)
    out_fh = None
    if args.out:
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_fh = out_path.open("a", encoding="utf-8")
    print(f"[monitor] logs_dir={logs_dir}")
    stdout_offset = safe_file_size(logs_dir / "last_engine_stdout.txt")
    print(f"[monitor] baseline_stdout_size={stdout_offset}")
    if out_fh:
        print(f"[monitor] jsonl={out_fh.name}")
    started = time.monotonic()
    previous = None
    stdout_totals = dict.fromkeys(SIGNAL_KEYS, 0)
    stress_tracker = StressTracker(args.stall_seconds) if args.stress else None
    try:
        while True:
            snapshot = make_snapshot(
                logs_dir,
                include_tasklist=not args.no_tasklist,
                stdout_offset=stdout_offset,
            )
            stdout_offset = snapshot["stdout"]["next_offset"]
            for key in SIGNAL_KEYS:
                stdout_totals[key] += snapshot["stdout"].get(key, 0)
            snapshot["stdout_since_start"] = dict(stdout_totals)
            line = compact_line(
                snapshot,
                previous,
                verbose=args.verbose or args.stress,
                stress_tracker=stress_tracker,
            )
            if line:
                print(line, flush=True)
            if out_fh:
                out_fh.write(json.dumps(snapshot, ensure_ascii=False) + "\n")
                out_fh.flush()
            previous = snapshot
            if args.duration and time.monotonic() - started >= args.duration:
                return 0
            time.sleep(max(float(args.interval or 1.0), 0.2))
    except KeyboardInterrupt:
        print("[monitor] stopped")
        return 130
    finally:
        if stress_tracker:
            for line in stress_tracker.summary_lines():
                print(line)
        if out_fh:
            out_fh.close()


if __name__ == "__main__":
    raise SystemExit(main())
