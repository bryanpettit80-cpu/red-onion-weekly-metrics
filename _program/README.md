# Red Onion Program Notes

This folder contains the code and technical files for the Red Onion weekly metrics runner.

## Setup

From this `_program` folder:

```powershell
python -m pip install -e ".[dev]"
```

The operator launcher creates a machine-local virtual environment at:

```text
%LOCALAPPDATA%\RedOnionMetrics\.venv
```

## Validation

Run these checks before committing code changes:

```powershell
python -m pytest -q
python -m py_compile red_onion_weekly_metrics.py
```

From the repository root, this compile check is also valid:

```powershell
python -m py_compile _program\red_onion_weekly_metrics.py
```

## Runtime Folders

In the deployed Dropbox layout, runtime folders are rooted one level above the `Red Onion Weekly Metrics Automation` repository:

- `..\01 Daily Reports - Drop Here`
- `..\02 Finished Reports`
- `..\03 Archive`

When the repository is used outside that named Dropbox folder, the same numbered folders default to the repository root. The outer launcher passes `-OperationsRoot` explicitly.

The master workbook is rebuilt from active daily reports plus archived daily reports found under `03 Archive`.

## Master Workbook Management Layer

The master workbook separates performance level from momentum:

- Recent Momentum compares the latest complete week with up to four prior complete weeks across check average, wine percentage, rate of sale, and ticket time.
- 8-Week Direction compares the most recent four complete weeks with the preceding four. Eight usable server weeks with at least 100 guests in each four-week block are labeled `Full`; six or seven may be labeled `Developing` when both blocks meet the configured week and guest thresholds.
- Incomplete latest weeks are `Not Scored` and generate no server action. A low current sample also keeps Recent Momentum at `Not Scored`, while qualified longer-term context may still display.
- Performance level compares the latest values with optional management targets, falling back to the store rolling baseline.
- Average rank movement is capped at a one-point modifier on each horizon so it does not double-count the underlying metrics.
- Prominent server actions require both the current-week volume thresholds and enough prior full-week history.

The management `Server Scorecard` shows the action, current sample, performance, both trend horizons, and the exact weeks and guests used. The hidden `Server Week-over-Week Detail` tab remains an audit view of adjacent-week changes; it is not the management coaching trend.

`Management Setup` targets, owner names, and manual `Action Board` fields are read from the existing master before regeneration. The new workbook is written to a temporary file and atomically replaces the prior master only after validation succeeds.

Technical calculation and raw-detail sheets remain in the workbook but are hidden by default. Do not delete them from the generator; they provide auditability and chart sources.

## Maintenance Notes

- Keep dependencies small and listed in `pyproject.toml` and `requirements.txt`.
- Do not add external services or credentials for the weekly run.
- Do not commit customer-facing workbooks, Toast source files, or archive contents unless explicitly requested.
