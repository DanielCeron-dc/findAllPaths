# CPM / Project Network Path Finder

This script reads a CPM/activity network from an Excel file and prints **all paths** from start activities to end activities, including each path duration and the **critical path** (longest duration).

## Requirements
- Python 3.10+

## Setup
```bash
# from this folder
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Windows (PowerShell):
```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Run

### Basic path enumeration (`cpm_paths.py`)
```bash
python cpm_paths.py GPCrashTimes.xlsx
```
Excel input format: `Activity | Activity time | Predecessor 1..N`.

### Crashing analysis (`cpm_crashing.py`)
```bash
python cpm_crashing.py CrashingWorkshop.xlsx
```
Excel input format:
`Activity | Normal time | Crash time | Normal Cost | Crash Cost | Predecessor 1..N`

This program prints:
- All paths with **normal** times and slacks (vs the longest normal path).
- All paths with **crash (intensive)** times and slacks (vs the longest crash path).
- The **optimal project time** (= length of the critical path under crash times).
- The **optimal project cost** (cheapest schedule that still meets the optimal time, found via LP).
- Per-activity recommended schedule (`tiempo óptimo sugerido`, días desintensificados, savings).
- All paths re-evaluated with the suggested optimal times.

## Notes
- The script assumes the network is a DAG (no cycles). If there are cycles, DFS may recurse indefinitely.
- End activities are those with no outgoing edges (not listed as a predecessor for any other activity).
