# Apex OLED System Monitor (LibreHardwareMonitor + GameSense)

This project displays live system information (CPU, GPU, RAM) on the OLED screen of SteelSeries keyboards (e.g. Apex 5) using:

* **LibreHardwareMonitor** as sensor source (via HTTP API)
* **SteelSeries GG GameSense SDK** for OLED output

## Features

* 3-page rotating display:

  * **CPU page**: clock speed, load, temperature
  * **GPU page**: load, temperature, VRAM usage
  * **RAM page**: total/used memory + usage bar
* Automatic page rotation
* Configurable refresh interval
* Automatic reconnect if SteelSeries GG restarts
* Automatic handling of **dynamic GameSense port**
* Optional LibreHardwareMonitor auto-start
* Windows mutex protection (prevents double start)
* Optional EXE build (no Python needed on target machine)
* Log file for debugging

---

## Requirements

### External software

* **SteelSeries GG** (must be running)
* **LibreHardwareMonitor**

  * Enable: `Options → Remote Web Server → Run`

### If running from Python source

* Python 3.10+
* Install dependencies:

```bash
pip install -r requirements.txt
```

### If running the compiled EXE

* Python is **NOT required**
* Only these are needed:

  * SteelSeries GG
  * LibreHardwareMonitor

---

## Installation

### 1) Clone repository

```bash
git clone https://github.com/YOURNAME/apex-oled-system-monitor.git
cd apex-oled-system-monitor
```

### 2) Install Python dependencies (if using script)

```bash
pip install -r requirements.txt
```

### 3) Configure paths

Edit in `src/oled_lhm.py`:

```python
LHM_EXE = r"C:\Tools\LHM\LibreHardwareMonitor.exe"
```

And verify:

```python
LHM_DATAJSON_URL = "http://localhost:8085/data.json"
```

---

## Running

### Run from Python:

```bash
python src/oled_lhm.py
```

### Build EXE:

```bash
pip install pyinstaller
cd src
pyinstaller --onefile --noconsole --name LHM_OLED oled_lhm.py
```

Result:

```
dist/LHM_OLED.exe
```

---

## Autostart (recommended)

Use **Windows Task Scheduler**:

* Trigger: `At startup`
* Delay: `20–30 seconds`
* Action: start `LHM_OLED.exe`
* Setting: "If task is already running → Do not start a new instance"

---

## Repo Structure

```
apex-oled-system-monitor/
 ├── src/
 │    └── oled_lhm.py
 ├── README.md
 ├── requirements.txt
 └── .gitignore
```
---
## Notes
* GameSense uses a **dynamic local port**. This project automatically detects it.
* If SteelSeries GG restarts, the script automatically rebinds.
* LibreHardwareMonitor must be running with Remote Web Server enabled.
---
## License
MIT 

