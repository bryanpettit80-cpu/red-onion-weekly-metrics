# Red Onion Weekly Metrics

This folder is the weekly Red Onion metrics workspace. It turns Toast/Marketing Vitals daily `.xls` reports into weekly public snapshots and the internal master workbook.

## Folder Layout

```text
Red Onion Metrics\
  Daily Reports\
  Output\
  Archive - Old Files\
  _program\
  Run Weekly Snapshot.cmd
```

- `Daily Reports`: drop the current Toast daily `.xls` files here.
- `Output`: generated weekly workbooks appear here.
- `Archive - Old Files`: processed source files and old local files are preserved here.
- `_program`: code, config, dependencies, tests, and the internal PowerShell runner. Operators should not need this folder.
- `Run Weekly Snapshot.cmd`: the only root file operators need to run.

## Weekly Run

1. Save the current Toast/Marketing Vitals files into `Daily Reports`.
2. Double-click `Run Weekly Snapshot.cmd`.
3. Open the generated workbooks in `Output`.

The report files should be named like:

```text
Daily Report - TM (Auto-Run) - Marketing Vitals - 06-10-2026.xls
```

The program reads the business date inside the workbook. The filename date is only a fallback.

## What The Run Does

- Uses Red Onion's existing Tuesday-Sunday operating week. Mondays are closed and excluded from weekly public snapshots.
- Stops if `Daily Reports` is empty.
- Stops if `Daily Reports` contains files from more than one Tuesday-Sunday operating week and lists the files to fix.
- Creates or replaces:
  - `Output\Check_Wine_RVA<week_end>.xlsx`
  - `Output\Check_Wine_VB<week_end>.xlsx`
  - `Output\Red_Onion_Server_Master.xlsx`
- The master workbook includes rising/falling server trends, store trend summaries, and an all-stores group trend summary.
- Moves successfully processed source files to:

```text
Archive - Old Files\processed-daily-reports\week-ending-YYYY-MM-DD\
```

The master workbook reads archived daily reports plus the active files in `Daily Reports`, so moving processed source files does not remove historical rows from the master workbook.

## If Something Fails

If parsing or workbook creation fails, the source files stay in `Daily Reports` so they can be fixed and rerun.

Close any open output workbook before rerunning. Excel can block replacement while a workbook is open.

If Python is missing, install Python 3.9 or newer and select `Add python.exe to PATH`.

## Configuration

Configuration lives in `_program\red_onion_config.json`. It controls location short codes, minimum guest counts, dashboard trend eligibility, display aliases, and public/dashboard exclusions.

## What Is Not Redundant

- `Run Weekly Snapshot.cmd` is the operator launcher to double-click.
- `_program\Run-WeeklySnapshot.ps1` is the internal runner used by the launcher.
- `_program\red_onion_weekly_metrics.py` is the actual workbook-building program.
