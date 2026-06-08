# Red Onion Weekly Metrics Automation

This program turns the daily Toast/Marketing Vitals `.xls` files into:

- Public weekly check average and wine percentage snapshots for Richmond and Virginia Beach.
- An internal master workbook with daily detail, weekly server metrics, location metrics, and trend summaries.

The assumed working folder is Dropbox:

```text
C:\Users\<your user>\Dropbox\Red Onion Metrics
```

## Folder Layout

The Dropbox folder should look like this:

```text
Red Onion Metrics\
  Run-WeeklySnapshot.ps1
  red_onion_weekly_metrics.py
  red_onion_config.json
  requirements.txt
  README.md
  source_daily_reports\
  outputs\
  sample_outputs\
```

Use the folders this way:

- `source_daily_reports`: put all raw daily Toast/Marketing Vitals `.xls` files here.
- `outputs`: generated weekly public files and the internal master workbook appear here.
- `sample_outputs`: reference examples only. Do not put new raw daily reports here.

## First-Time Setup On A Computer

Each computer needs:

1. Dropbox desktop installed and synced.
2. Python 3.9 or newer installed.
3. Internet access for the first run so Python packages can install.

When installing Python, check the box for `Add python.exe to PATH`.

Dependency files should not sync through Dropbox. The runner creates a computer-specific environment here:

```text
%LOCALAPPDATA%\RedOnionMetrics\.venv
```

That keeps Dropbox clean and lets the same synced project run on different computers.

## Weekly Run Steps

1. Save the raw daily `.xls` files into:

```text
C:\Users\<your user>\Dropbox\Red Onion Metrics\source_daily_reports
```

2. Keep historical raw files in `source_daily_reports` if you want the master workbook to keep long-term history.

3. Open PowerShell.

4. Run:

```powershell
cd "$env:USERPROFILE\Dropbox\Red Onion Metrics"
.\Run-WeeklySnapshot.ps1
```

If Windows blocks script execution, run this instead:

```powershell
cd "$env:USERPROFILE\Dropbox\Red Onion Metrics"
powershell -ExecutionPolicy Bypass -File ".\Run-WeeklySnapshot.ps1"
```

5. Open the generated files in:

```text
C:\Users\<your user>\Dropbox\Red Onion Metrics\outputs
```

## Source File Rules

The program reads files in `source_daily_reports` named like:

```text
Daily Report - TM (Auto-Run) - Marketing Vitals - 06-08-2026.xls
```

The filename must start with `Daily Report` and end with `.xls`.

The weekly public snapshot uses the latest report date found in the source folder and builds the prior 7-day window from the available files. The master workbook rebuilds from every raw daily report found in `source_daily_reports`.

## Outputs

The run creates or replaces:

- `outputs\Check_Wine_RVA<week_end>.xlsx`
- `outputs\Check_Wine_VB<week_end>.xlsx`
- `outputs\Red_Onion_Server_Master.xlsx`

Close any open output workbook before rerunning. Excel lock files can prevent the script from replacing an open workbook.

## Public Snapshot Rules

The public files show:

- Store
- Server/display name
- Check Average
- Wine %

Gross sales, guest count, and wine sales remain in the workbook for formulas but are hidden from view.

Public snapshot highlighting:

- Top check average is green.
- Top wine percentage is yellow.
- Bottom three check averages are light red.

Public exclusions are controlled by `red_onion_config.json`. Excluded names are removed only from the public snapshots. The internal master workbook still includes them.

## Master Workbook Logic

The master workbook is rebuilt from every raw daily report found in `source_daily_reports`. It includes rows excluded from public posting so store performance remains accurate.

The master workbook includes these tabs:

- `Dashboard`: latest weekly location results, current server leaderboards, and a quick check average chart.
- `Run Notes`: source folder, date coverage, public snapshot dates, exclusions, and metric rules.
- `Weekly Server Metrics`: one weekly rollup row per server/location.
- `Weekly Server Rankings`: weekly server rankings for check average, wine percentage, rate of sale, and ticket time.
- `Server Week Trends`: week-over-week server changes, rank movement context, and a trend note.
- `Weekly Location Metrics`: one weekly rollup row per location.
- `Daily Server Detail`: daily source rows by server.
- `Daily Location Detail`: daily source rows by location.
- `Server Trend Summary`: all-time server summary plus latest/prior week comparisons.
- `Data Quality`: source file, date coverage, and location coverage checks.

Calculated metrics use these rollup rules:

- Check average = total gross sales / total guest count.
- Wine percentage = total wine sales / total gross sales.
- Rate of sale by guest count = guest-weighted average of the daily rate, displayed as a decimal to match the raw report. Lower values are better.
- Average ticket time = guest-weighted average of the daily ticket time.

Ranking logic:

- Check average and wine percentage rank highest value as best.
- Rate of sale by guest count and ticket time rank lowest value as best.
- Rankings use `master_min_guest_count_for_rankings` from `red_onion_config.json` if present. If that setting is not present, the program uses `public_min_guest_count`.

## Config Updates

Edit `red_onion_config.json` to adjust:

- Location short codes used in public file names.
- Minimum guest count for public posting.
- Public display aliases such as `Bar 1 Bar 1` to `Bar`.
- Names or name fragments to keep out of the public competition snapshot.

After editing the config, rerun `Run-WeeklySnapshot.ps1`.

## Troubleshooting

If Python is missing, install Python 3.9 or newer and check `Add python.exe to PATH`.

If PowerShell blocks the script, use:

```powershell
powershell -ExecutionPolicy Bypass -File ".\Run-WeeklySnapshot.ps1"
```

If an output file cannot be replaced, close the workbook in Excel and rerun the script.

If a source file is missing from the master workbook, confirm it is in `source_daily_reports`, starts with `Daily Report`, ends with `.xls`, and has synced locally in Dropbox.
