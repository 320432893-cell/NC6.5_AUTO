import time
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_probe import (  # noqa: E402
    configure_jab,
    enum_windows,
    hwnd_int,
    load_access_bridge,
    run_windows_access_bridge,
)


def main():
    dll, path = load_access_bridge(
        r"C:\Users\Queclink\AppData\Local\UClient\share\java1.7.0_51-x64\bin\WindowsAccessBridge-64.dll"
    )
    configure_jab(dll)
    print(f"dll={path}")
    stop_pump = None
    pump_thread = None
    if hasattr(dll, "initializeAccessBridge"):
        print("init_mode=initializeAccessBridge")
        print(f"init1={bool(dll.initializeAccessBridge())}")
        time.sleep(1)
        print(f"init2={bool(dll.initializeAccessBridge())}")
    else:
        print("init_mode=Windows_run")
        stop_pump = threading.Event()
        pump_thread = threading.Thread(
            target=run_windows_access_bridge, args=(dll, stop_pump), daemon=True
        )
        pump_thread.start()
        time.sleep(1)

    rows = []
    for hwnd, title, class_name, _pid, visible in enum_windows(include_children=False):
        if class_name.startswith("SunAwt") or title == "Yonyou UClient":
            rows.append(
                {
                    "hwnd": hwnd_int(hwnd),
                    "title": title,
                    "class": class_name,
                    "visible": bool(visible),
                    "is_java": bool(dll.isJavaWindow(hwnd)),
                }
            )
    print(rows)
    if stop_pump:
        stop_pump.set()
    if pump_thread:
        pump_thread.join(timeout=1)


if __name__ == "__main__":
    main()
