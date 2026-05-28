import json
import time
from datetime import datetime
from pathlib import Path


class RunStateRecorder:
    """Write the latest workflow state for interruption review."""

    def __init__(self, command=None, config=None, path=None):
        self.started_perf = time.perf_counter()
        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.path = Path(path or "logs/run_state.json")
        self.path.parent.mkdir(exist_ok=True)
        self.events_limit = 200
        self.data = {
            "run_id": self.run_id,
            "command": command or "",
            "status": "running",
            "stage": "initialized",
            "started_at": datetime.now().isoformat(timespec="seconds"),
            "updated_at": "",
            "ended_at": "",
            "elapsed_s": 0.0,
            "error": "",
            "excel_path": (config or {}).get("excel_path", ""),
            "sheet": (config or {}).get("sheet_my", ""),
            "counts": {},
            "current": {},
            "events": [],
        }
        self.write()

    def set_stage(self, stage, **fields):
        self.data["stage"] = stage
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
        self.write()

    def finish(self, status, error=None):
        self.data["status"] = status
        self.data["ended_at"] = datetime.now().isoformat(timespec="seconds")
        self.data["elapsed_s"] = self.elapsed_s()
        if error:
            self.data["error"] = str(error)
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
