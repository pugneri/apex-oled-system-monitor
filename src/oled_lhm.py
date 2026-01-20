"""
Apex OLED System Monitor (LibreHardwareMonitor + SteelSeries GameSense)

Displays CPU / GPU / RAM information on the OLED screen of SteelSeries keyboards
(e.g., Apex 5) using:
- LibreHardwareMonitor Remote Web Server (data.json) as the sensor source
- SteelSeries GG GameSense HTTP API for OLED output

Features:
- 3 rotating pages (CPU, GPU, RAM)
- Configurable refresh interval and page rotation
- Dynamic GameSense port detection (coreProps.json changes every GG start)
- Auto rebind if SteelSeries GG restarts
- Optional LHM auto-start and readiness wait
- Single-instance protection (Windows mutex)
- Optional log file for troubleshooting

Author: (Istvan Pugner)
License: MIT
"""

from __future__ import annotations

import datetime
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Dict, Iterable, Optional

import psutil
import requests
import schedule

# ============================================================
# CONFIG
# ============================================================

# Update interval for values (seconds)
UPDATE_SECONDS = 3

# Page rotation interval (seconds)
PAGE_SECONDS = 12

# If True, value updates will not force refresh when page changes (advanced)
PAGE_LOCK = False

# Temperature formatting (no space): 52°C
USE_DEGREE_SYMBOL = True
DEGREE_SYMBOL = "°"  # If OLED shows a box, try "º" or set USE_DEGREE_SYMBOL=False

# LibreHardwareMonitor Remote Web Server endpoint
LHM_DATAJSON_URL = "http://localhost:8085/data.json"

# Optionally start LibreHardwareMonitor if it is not running/ready
# (Set False if you already start LHM via Task Scheduler)
AUTO_START_LHM = False
LHM_EXE = r"C:\Tools\LHM\LibreHardwareMonitor.exe"
LHM_ARGS: list[str] = []  # Usually empty
WAIT_FOR_LHM_SECONDS = 30

# SteelSeries GG / GameSense
WAIT_FOR_GAMESENSE_SECONDS = 120
AUTO_REBIND_ON_FAIL = True
REBIND_COOLDOWN_SECONDS = 10

# Prefer GPU load value:
# - "d3d": uses "D3D 3D" %
# - "core": uses "GPU Core" %
GPU_LOAD_SOURCE = "d3d"

# RAM progress bar
BAR_WIDTH = 16  # total bar chars between brackets, total line length = 18
BAR_FILLED = "#"
BAR_EMPTY = "-"

# Logging (useful for EXE builds with no console)
LOG_TO_FILE = True
LOG_FILE = r"C:\Tools\OLED\oled_lhm.log"

# GameSense "game" + events (3 pages)
GAME = "LHM_OLED"
CPU_EVENT = "CPU_PAGE"
GPU_EVENT = "GPU_PAGE"
RAM_EVENT = "RAM_PAGE"

# Possible locations of coreProps.json (varies by GG/Engine version)
COREPROPS_PATHS = [
    r"C:\ProgramData\SteelSeries\SteelSeries Engine 3\coreProps.json",
    r"C:\ProgramData\SteelSeries\GG\coreProps.json",
    r"C:\ProgramData\SteelSeries\SteelSeries GG\coreProps.json",
]

# ============================================================
# Single instance protection (Windows mutex)
# ============================================================

def ensure_single_instance(mutex_name: str = "Global\\LHM_OLED_GAMESENSE_MONITOR") -> None:
    """
    Prevents running multiple instances at the same time.

    Requires pywin32. If pywin32 is missing, this check is skipped.
    """
    try:
        import win32api
        import win32event
        import winerror
    except Exception:
        return

    mutex = win32event.CreateMutex(None, False, mutex_name)
    last_error = win32api.GetLastError()
    if last_error == winerror.ERROR_ALREADY_EXISTS:
        sys.exit(0)

    # Keep mutex referenced to avoid GC
    globals()["_SINGLE_INSTANCE_MUTEX"] = mutex


ensure_single_instance()


# ============================================================
# Logging
# ============================================================

def log(message: str) -> None:
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {message}"
    print(line)
    if LOG_TO_FILE:
        try:
            os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass


# ============================================================
# GameSense dynamic port handling
# ============================================================

_current_base: Optional[str] = None
_last_base_read: float = 0.0
_last_rebind: float = 0.0
_counter: int = 0


def read_coreprops_address() -> Optional[str]:
    """
    Reads the dynamic GameSense HTTP API address from coreProps.json.
    Returns the base URL (e.g., "http://127.0.0.1:60578") or None.
    """
    for p in COREPROPS_PATHS:
        try:
            cp = Path(p)
            if not cp.exists():
                continue
            data = json.loads(cp.read_text(encoding="utf-8"))
            addr = data.get("address")
            if addr:
                return f"http://{addr}"
        except Exception:
            continue
    return None


def is_gamesense_alive(base: str) -> bool:
    """
    Tests whether GameSense HTTP API is reachable.
    """
    try:
        r = requests.post(f"{base}/game_heartbeat", json={"game": "PING"}, timeout=1.2)
        return r.status_code in (200, 204)
    except Exception:
        return False


def get_current_base(force: bool = False) -> Optional[str]:
    """
    Returns a working GameSense base URL.
    Automatically adapts to GG restarts and dynamic port changes.
    """
    global _current_base, _last_base_read
    now = time.time()

    # Avoid re-reading coreProps too frequently
    if not force and _current_base and (now - _last_base_read) < 3:
        return _current_base

    base = read_coreprops_address()
    _last_base_read = now

    if base and is_gamesense_alive(base):
        if base != _current_base:
            log(f"GameSense base changed -> {base}")
        _current_base = base
        return _current_base

    # keep last known base if still valid, otherwise None
    return _current_base


def gamesense_ok() -> bool:
    base = get_current_base()
    return bool(base and is_gamesense_alive(base))


def safe_post(path: str, payload: Dict[str, Any]) -> Optional[requests.Response]:
    """
    POST to GameSense API using the currently detected base URL.
    If it fails, optionally rebind/re-register.
    """
    global _last_rebind
    base = get_current_base()
    if not base:
        return None

    try:
        r = requests.post(f"{base}{path}", json=payload, timeout=2)
        if r.status_code >= 400 and AUTO_REBIND_ON_FAIL:
            now = time.time()
            if now - _last_rebind > REBIND_COOLDOWN_SECONDS:
                _last_rebind = now
                log(f"POST {path} failed ({r.status_code}) -> rebind")
                try_rebind()
        return r
    except Exception as e:
        if AUTO_REBIND_ON_FAIL:
            now = time.time()
            if now - _last_rebind > REBIND_COOLDOWN_SECONDS:
                _last_rebind = now
                log(f"POST {path} exception ({e}) -> refresh base + rebind")
                get_current_base(force=True)
                try_rebind()
        return None


def register_game() -> None:
    safe_post(
        "/game_metadata",
        {
            "game": GAME,
            "game_display_name": "LHM OLED Monitor",
            "developer": "local",
            "deinitialize_timer_length_ms": 60000,
        },
    )


def bind_screen(event_name: str) -> None:
    safe_post(
        "/bind_game_event",
        {
            "game": GAME,
            "event": event_name,
            "icon_id": 1,
            "handlers": [
                {
                    "device-type": "screened-128x40",
                    "zone": "one",
                    "mode": "screen",
                    "datas": [
                        {
                            "lines": [
                                {"has-text": True, "context-frame-key": "l1"},
                                {"has-text": True, "context-frame-key": "l2"},
                            ]
                        }
                    ],
                }
            ],
        },
    )


def bind_all() -> None:
    bind_screen(CPU_EVENT)
    bind_screen(GPU_EVENT)
    bind_screen(RAM_EVENT)


def try_rebind() -> None:
    try:
        register_game()
        bind_all()
        log("Rebind done.")
    except Exception as e:
        log(f"Rebind failed: {e}")


def send_screen(event_name: str, line1: str, line2: str) -> None:
    global _counter
    safe_post(
        "/game_event",
        {
            "game": GAME,
            "event": event_name,
            "data": {
                "value": _counter,
                "frame": {"l1": (line1 or "")[:18], "l2": (line2 or "")[:18]},
            },
        },
    )
    _counter += 1


def heartbeat() -> None:
    safe_post("/game_heartbeat", {"game": GAME})


# ============================================================
# LibreHardwareMonitor auto-start / readiness
# ============================================================

def is_lhm_ready(url: str) -> bool:
    try:
        r = requests.get(url, timeout=1.5)
        return r.status_code == 200 and "Children" in r.text
    except Exception:
        return False


def start_lhm_if_needed() -> None:
    if not AUTO_START_LHM:
        return
    if is_lhm_ready(LHM_DATAJSON_URL):
        return

    if not os.path.exists(LHM_EXE):
        log(f"LHM exe not found: {LHM_EXE}")
        return

    try:
        log("Starting LibreHardwareMonitor...")
        subprocess.Popen([LHM_EXE, *LHM_ARGS], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        log(f"Failed to start LHM: {e}")


def wait_for_lhm(timeout_s: int) -> bool:
    start = time.time()
    while time.time() - start < timeout_s:
        if is_lhm_ready(LHM_DATAJSON_URL):
            log("LHM ready.")
            return True
        time.sleep(1)
    log("LHM not ready (timeout).")
    return False


# ============================================================
# Formatting / rendering
# ============================================================

def fmt_pct(v: Optional[float]) -> str:
    return "?" if v is None else f"{int(round(v))}%"


def fmt_temp(v: Optional[float]) -> str:
    if v is None:
        return "?"
    t = int(round(v))
    if USE_DEGREE_SYMBOL:
        return f"{t}{DEGREE_SYMBOL}C"
    return f"{t}C"


def fmt_gb(v: Optional[float]) -> str:
    return "?" if v is None else f"{v:.1f}G"


def fmt_ghz_from_mhz(mhz: Optional[float]) -> str:
    return "?" if mhz is None else f"{mhz / 1000.0:.2f}GHz"


def progress_bar(used: Optional[float], total: Optional[float]) -> str:
    """
    Returns an 18-character bar: [################] or [----...]
    """
    if used is None or total is None or total <= 0:
        return "[" + (BAR_EMPTY * BAR_WIDTH) + "]"
    frac = max(0.0, min(1.0, used / total))
    filled = int(round(frac * BAR_WIDTH))
    filled = max(0, min(BAR_WIDTH, filled))
    return "[" + (BAR_FILLED * filled) + (BAR_EMPTY * (BAR_WIDTH - filled)) + "]"


# ============================================================
# LHM JSON parsing (tailored to typical LHM node structure)
# ============================================================

def _to_float_best(val: Any) -> Optional[float]:
    """
    Converts strings like '52,6 °C' or '6,3 %' to float (52.6 / 6.3).
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)

    s = str(val).strip()
    num = ""
    for ch in s:
        if ch.isdigit() or ch in ".-,":
            num += ch
        elif num:
            break
    if not num:
        return None
    return float(num.replace(",", "."))


def walk_json(node: Any) -> Iterable[Dict[str, Any]]:
    if isinstance(node, dict):
        yield node
        for v in node.values():
            yield from walk_json(v)
    elif isinstance(node, list):
        for it in node:
            yield from walk_json(it)


def lhm_read_metrics() -> Dict[str, Optional[float]]:
    """
    Reads:
      CPU:
        - Core (Tctl/Tdie) temperature
        - CPU Total load
        - Cores (Average) clock (MHz)
      GPU:
        - GPU Core temperature (fallback: GPU Hot Spot)
        - D3D 3D load (fallback: GPU Core %)
        - GPU Memory Used/Total (MB)
    """
    try:
        r = requests.get(LHM_DATAJSON_URL, timeout=2)
        r.raise_for_status()
        data = r.json()
    except Exception:
        return {
            "cpu_temp": None,
            "cpu_load": None,
            "cpu_clock_mhz": None,
            "gpu_temp": None,
            "gpu_load": None,
            "vram_used_mb": None,
            "vram_total_mb": None,
        }

    cpu_temp = cpu_load = cpu_clock_mhz = None
    gpu_temp = gpu_hotspot = None
    gpu_core_load = d3d_3d = None
    vram_used_mb = vram_total_mb = None

    for n in walk_json(data):
        text = n.get("Text")
        val = n.get("Value")
        if text is None or val is None:
            continue

        lname = str(text).strip().lower()
        sval = str(val).strip().lower()

        # CPU
        if cpu_temp is None and lname == "core (tctl/tdie)" and "c" in sval:
            cpu_temp = _to_float_best(val)
        if cpu_load is None and lname == "cpu total" and "%" in sval:
            cpu_load = _to_float_best(val)
        if cpu_clock_mhz is None and lname == "cores (average)" and "mhz" in sval:
            cpu_clock_mhz = _to_float_best(val)

        # GPU temp
        if gpu_temp is None and lname == "gpu core" and "c" in sval:
            gpu_temp = _to_float_best(val)
        if gpu_hotspot is None and lname == "gpu hot spot" and "c" in sval:
            gpu_hotspot = _to_float_best(val)

        # GPU load
        if gpu_core_load is None and lname == "gpu core" and "%" in sval:
            gpu_core_load = _to_float_best(val)
        if d3d_3d is None and lname == "d3d 3d" and "%" in sval:
            d3d_3d = _to_float_best(val)

        # VRAM
        if vram_used_mb is None and lname == "gpu memory used" and "mb" in sval:
            vram_used_mb = _to_float_best(val)
        if vram_total_mb is None and lname == "gpu memory total" and "mb" in sval:
            vram_total_mb = _to_float_best(val)

    # Select GPU load source
    if GPU_LOAD_SOURCE.lower() == "core":
        gpu_load = gpu_core_load or d3d_3d
    else:
        gpu_load = d3d_3d or gpu_core_load

    # Prefer core temp, fallback hotspot
    gpu_temp_final = gpu_temp or gpu_hotspot

    return {
        "cpu_temp": cpu_temp,
        "cpu_load": cpu_load,
        "cpu_clock_mhz": cpu_clock_mhz,
        "gpu_temp": gpu_temp_final,
        "gpu_load": gpu_load,
        "vram_used_mb": vram_used_mb,
        "vram_total_mb": vram_total_mb,
    }


def mb_to_gb(mb: Optional[float]) -> Optional[float]:
    return None if mb is None else mb / 1024.0


# ============================================================
# Runtime state
# ============================================================

MET: Dict[str, Optional[float]] = {
    "cpu_clock_mhz": None,
    "cpu_load": None,
    "cpu_temp": None,
    "gpu_load": None,
    "gpu_temp": None,
    "vram_used_gb": None,
    "vram_total_gb": None,
    "ram_used_gb": None,
    "ram_total_gb": None,
}

PAGE = 0


def update_metrics() -> None:
    m = lhm_read_metrics()

    MET["cpu_clock_mhz"] = m["cpu_clock_mhz"]
    MET["cpu_load"] = m["cpu_load"]
    MET["cpu_temp"] = m["cpu_temp"]

    MET["gpu_load"] = m["gpu_load"]
    MET["gpu_temp"] = m["gpu_temp"]

    MET["vram_used_gb"] = mb_to_gb(m["vram_used_mb"])
    MET["vram_total_gb"] = mb_to_gb(m["vram_total_mb"])

    # RAM from OS (accurate physical RAM, GiB)
    try:
        vm = psutil.virtual_memory()
        MET["ram_total_gb"] = vm.total / (1024**3)
        MET["ram_used_gb"] = (vm.total - vm.available) / (1024**3)
    except Exception:
        MET["ram_total_gb"] = None
        MET["ram_used_gb"] = None

    if PAGE_LOCK:
        render_page()


def render_page() -> None:
    if PAGE == 0:
        # CPU page:
        # Line 1: CPU clock speed in GHz
        # Line 2: CPU load + temperature
        l1 = f"CPU {fmt_ghz_from_mhz(MET['cpu_clock_mhz'])}"
        l2 = f"CPU {fmt_pct(MET['cpu_load'])} {fmt_temp(MET['cpu_temp'])}"
        send_screen(CPU_EVENT, l1, l2)

    elif PAGE == 1:
        # GPU page:
        # Line 1: GPU load + temperature
        # Line 2: VRAM total/used
        l1 = f"GPU {fmt_pct(MET['gpu_load'])} {fmt_temp(MET['gpu_temp'])}"
        l2 = f"VRAM {fmt_gb(MET['vram_total_gb'])}/{fmt_gb(MET['vram_used_gb'])}"
        send_screen(GPU_EVENT, l1, l2)

    else:
        # RAM page:
        # Line 1: RAM total/used
        # Line 2: RAM usage bar
        l1 = f"RAM {fmt_gb(MET['ram_total_gb'])}/{fmt_gb(MET['ram_used_gb'])}"
        l2 = progress_bar(MET["ram_used_gb"], MET["ram_total_gb"])
        send_screen(RAM_EVENT, l1, l2)


def rotate_page() -> None:
    global PAGE
    PAGE = (PAGE + 1) % 3
    if not PAGE_LOCK:
        render_page()


def watchdog() -> None:
    """
    Periodic health checks:
    - LHM: if not ready, optionally start it
    - GameSense: if not reachable, rebind
    """
    if not is_lhm_ready(LHM_DATAJSON_URL):
        log("LHM not ready -> check/start")
        start_lhm_if_needed()

    if AUTO_REBIND_ON_FAIL and not gamesense_ok():
        log("GameSense not OK -> rebind")
        get_current_base(force=True)
        try_rebind()


# ============================================================
# Main
# ============================================================

def wait_for_gamesense(timeout_s: int) -> bool:
    log("Waiting for GameSense...")
    start = time.time()
    while time.time() - start < timeout_s:
        base = get_current_base(force=True)
        if base:
            log(f"GameSense ready at {base}")
            return True
        time.sleep(1)
    log("GameSense not ready (timeout).")
    return False


def main() -> None:
    log("Starting OLED app...")

    start_lhm_if_needed()
    wait_for_lhm(WAIT_FOR_LHM_SECONDS)

    wait_for_gamesense(WAIT_FOR_GAMESENSE_SECONDS)

    # Register and bind events
    register_game()
    bind_all()
    log("GameSense registered & bound.")

    # Initial draw
    send_screen(CPU_EVENT, "LHM OLED", "starting...")
    update_metrics()
    render_page()

    # Schedulers
    schedule.every(5).seconds.do(heartbeat)
    schedule.every(UPDATE_SECONDS).seconds.do(update_metrics)
    schedule.every(UPDATE_SECONDS).seconds.do(render_page)
    schedule.every(PAGE_SECONDS).seconds.do(rotate_page)
    schedule.every(10).seconds.do(watchdog)

    while True:
        schedule.run_pending()
        time.sleep(0.2)


if __name__ == "__main__":
    main()

