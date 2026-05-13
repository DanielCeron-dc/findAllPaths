# CPM / Project Network Crashing Analyzer

Reads a CPM/activity network from an Excel file, enumerates all paths from start to end activities, and runs a full crashing analysis (normal vs. intensive times, optimal time and cost).

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

### Crashing analysis (`cpm_crashing.py`)
```bash
python cpm_crashing.py Taller3.xlsx
```

Excel input format:
`Activity | Normal time | Crash time | Normal Cost | Crash Cost | Predecessor 1..N`

See `Taller3.xlsx` for a working example.

This program prints:
- All paths with **normal** times and slacks (vs the longest normal path).
- All paths with **crash (intensive)** times and slacks (vs the longest crash path).
- The **optimal project time** (= length of the critical path under crash times).
- The **optimal project cost** (cheapest schedule that still meets the optimal time, found via LP).
- Per-activity recommended schedule (suggested optimal time, days de-intensified, savings).
- All paths re-evaluated with the suggested optimal times.

### Word report (`generate_report.py`)
```bash
python generate_report.py Taller3.xlsx
```
Generates `Reporte_Taller3.docx` with the full analysis in tabular form.

## Notes
- The script assumes the network is a DAG (no cycles). If there are cycles, DFS may recurse indefinitely.
- End activities are those with no outgoing edges (not listed as a predecessor for any other activity).
