# TrackerSleep

**A tool for tracking sleep schedules** — designed for cases of chronic sleep disruption, as well as medical and/or research tracking.

[🇷🇺 Русская версия](https://github.com/Argonstudio/Tracker_sleep/tree/russian-version)

---

## ✨ Features

The program takes the following parameters into account:

| Parameter | Description |
|-----------|-------------|
| **Bedtime** | Date (day:month:year) and exact time |
| **Wake-up time** | Date (day:month:year) and exact time |
| **Sleep duration** | Total hours of sleep |
| **Wake duration** | Total hours awake |
| **Shift** | How much later or earlier bedtime was compared to the previous day |
| **Day** | Actual length of the sleep–wake cycle (subjective day) |
| **Comment** | Optional note |

It also:

- Builds an Excel spreadsheet and draws a chart/graph.
- Tracks average, minimum, and maximum values for:
  
  - sleep duration
  - wake duration
  - bedtime shift
  - day length

---

## 🧠 Two Versions

### 1. Main version (automatic)
- Designed for internet geeks.
- Automatically counts any time outside Google Chrome lasting longer than **3 hours** as sleep.
- Based on pauses in browser history and Windows program logs.
- Automatically calculates sleep periods and presents them to the user — incorrect entries can be deleted.
- Calculation starts from the last date in the Excel file.

### 2. Second version (`tracker_sleep1.py`) — manual
- Allows manual date/time entry and adding comments.
- To use it: delete the main version and rename this file to `tracker_sleep.py`.

---

## ⚙️ Installation

1. Create a folder.
2. Download:
3. 
   - the desired program version file
   - the `.bat` file (launches the program with a double click)
   - `requirements.txt`
     
4. Open **Windows PowerShell** (Python must be installed).
5. Run the following commands:

```powershell
cd desired_folder
python -m venv .venv
./.venv/Scripts/Activate.ps1
pip install -r requirements.txt
