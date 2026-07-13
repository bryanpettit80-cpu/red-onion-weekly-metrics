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

The command-line defaults are rooted one level above `_program`:

- `..\Daily Reports`
- `..\Output`
- `..\Archive - Old Files`

The master workbook is rebuilt from active daily reports plus archived daily reports found under `Archive - Old Files`.

## Master Workbook Management Layer

The master workbook separates performance level from momentum:

- Momentum compares the latest full week with up to four prior full weeks across check average, wine percentage, rate of sale, and ticket time.
- Performance level compares the latest values with optional management targets, falling back to the store rolling baseline.
- Average rank movement is capped at a one-point modifier so it does not double-count the underlying metrics.
- Prominent server actions require both the current-week volume thresholds and enough prior full-week history.

`Management Setup` targets, owner names, and manual `Action Board` fields are read from the existing master before regeneration. The new workbook is written to a temporary file and atomically replaces the prior master only after validation succeeds.

Technical calculation and raw-detail sheets remain in the workbook but are hidden by default. Do not delete them from the generator; they provide auditability and chart sources.

## Maintenance Notes

- Keep dependencies small and listed in `pyproject.toml` and `requirements.txt`.
- Do not add external services or credentials for the weekly run.
- Do not commit customer-facing workbooks, Toast source files, or archive contents unless explicitly requested.
