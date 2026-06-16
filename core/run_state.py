import json
import time
from datetime import datetime
from pathlib import Path

from core.paths import clear_stop_flag, logs_dir


class RunStateRecorder:
    """Write the latest workflow state for interruption review."""

    def __init__(self, command=None, config=None, path=None):
        self.started_perf = time.perf_counter()
        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.path = Path(path) if path else logs_dir() / "run_state.json"
        self.path.parent.mkdir(parents=True, exist_ok=True)
        # 新一轮运行开始,清掉上次残留的外部停止标志(见 ENGINE_CONTRACT.md §1.4)
        clear_stop_flag()
        self.events_limit = 200
        self.data = {
            "run_id": self.run_id,
            "command": command or "",
            "status": "running",
            "stage": "initialized",
            "step_index": 0,
            "total_steps": 0,
            "started_at": datetime.now().isoformat(timespec="seconds"),
            "updated_at": "",
            "ended_at": "",
            "elapsed_s": 0.0,
            "error": "",
            "last_error": "",
            "excel_path": (config or {}).get("excel_path", ""),
            "sheet": (config or {}).get("sheet_my", ""),
            "counts": {},
            "current": {},
            "events": [],
        }
        self.write()

    def set_stage(self, stage, step_index=None, total_steps=None, **fields):
        self.data["stage"] = stage
        if step_index is not None:
            self.data["step_index"] = step_index
        if total_steps is not None:
            self.data["total_steps"] = total_steps
        if fields:
            self.data["current"].update(fields)
        self.write()

    def update_counts(self, **counts):
        self.data["counts"].update(counts)
        self.write()

    def event(self, name, **fields):
        events = self.data["events"]
        events.append(
            {
                "ts": datetime.now().isoformat(timespec="seconds"),
                "elapsed_s": self.elapsed_s(),
                "name": name,
                **fields,
            }
        )
        if len(events) > self.events_limit:
            del events[: len(events) - self.events_limit]
        # 运行中实时暴露最近一次错误,供 GUI 进度面板即时显示(不必等收尾)
        if fields.get("error"):
            self.data["last_error"] = str(fields["error"])
        self.write()

    def finish(self, status, error=None):
        self.data["status"] = status
        self.data["ended_at"] = datetime.now().isoformat(timespec="seconds")
        self.data["elapsed_s"] = self.elapsed_s()
        if error:
            self.data["error"] = str(error)
            self.data["last_error"] = str(error)
        self.write()

    def elapsed_s(self):
        return round(time.perf_counter() - self.started_perf, 3)

    def write(self):
        self.data["updated_at"] = datetime.now().isoformat(timespec="seconds")
        self.data["elapsed_s"] = self.elapsed_s()
        tmp_path = self.path.with_suffix(".json.tmp")
        with tmp_path.open("w", encoding="utf-8") as file:
            json.dump(self.data, file, ensure_ascii=False, indent=2, default=str)
            file.write("\n")
        tmp_path.replace(self.path)
