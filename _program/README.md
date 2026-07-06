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

## Maintenance Notes

- Keep dependencies small and listed in `pyproject.toml` and `requirements.txt`.
- Do not add external services or credentials for the weekly run.
- Do not commit customer-facing workbooks, Toast source files, or archive contents unless explicitly requested.
