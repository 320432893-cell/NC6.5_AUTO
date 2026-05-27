import json
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path


class PerfRecorder:
    def __init__(self, enabled=False, label=None):
        self.enabled = enabled
        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.label = label or self.run_id
        self.path = None
        if self.enabled:
            log_dir = Path("logs")
            log_dir.mkdir(exist_ok=True)
            self.path = log_dir / f"perf_{self.label}_{self.run_id}.jsonl"
            self.event(
                "perf_start",
                run_id=self.run_id,
                label=self.label,
                path=str(self.path),
            )

    def event(self, name, **fields):
        if not self.enabled or self.path is None:
            return
        payload = {
            "ts": datetime.now().isoformat(timespec="milliseconds"),
            "run_id": self.run_id,
            "event": name,
            **fields,
        }
        with self.path.open("a", encoding="utf-8") as file:
            file.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")

    @contextmanager
    def span(self, name, **fields):
        if not self.enabled:
            yield
            return

        start = time.perf_counter()
        error = None
        try:
            yield
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
            raise
        finally:
            elapsed = time.perf_counter() - start
            self.event(
                name,
                elapsed_s=round(elapsed, 6),
                error=error,
                **fields,
            )
