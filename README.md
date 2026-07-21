# Red Onion Weekly Metrics

This repository contains the Red Onion weekly metrics automation. In Dropbox it lives inside the operator workspace as `Red Onion Weekly Metrics Automation`; operators use the numbered folders and launcher one level above it.

## Folder Layout

```text
Red Onion Metrics\
  00 START HERE - Red Onion Weekly Metrics.txt
  01 Daily Reports - Drop Here\
  02 Finished Reports\
  03 Archive\
  Red Onion Weekly Metrics Automation\
  Run Weekly Snapshot.cmd
```

- `01 Daily Reports - Drop Here`: drop the current Toast daily `.xls` or `.xlsx` files here.
- `02 Finished Reports`: generated weekly workbooks appear here.
- `03 Archive`: processed reports, prior layouts, and historical workbooks are preserved here.
- `Red Onion Weekly Metrics Automation`: Git repository, code, config, tests, and technical documentation. Operators should not edit this folder.
- `Run Weekly Snapshot.cmd`: the only root file operators need to run.

## Weekly Run

1. Save the current Toast/Marketing Vitals files into `01 Daily Reports - Drop Here`.
2. Double-click `Run Weekly Snapshot.cmd`.
3. Open the generated workbooks in `02 Finished Reports`.

In the named Dropbox deployment, the launcher first verifies that `Red Onion Weekly Metrics Automation` is a clean Git checkout on `main` and that `HEAD` matches the checkout's local `origin/main` reference. A failed release check stops before the run creates folders, installs packages, builds workbooks, or moves source files. The check is intentionally local and does not fetch from GitHub; a technical maintainer must update and verify the deployed checkout as a separate release step.

The report files should be named like:

```text
Daily Report - TM (Auto-Run) - Marketing Vitals - 06-10-2026.xlsx
```

The program reads the business date inside the workbook. The filename date is only a fallback.
Excel temporary files whose names start with `~$` are ignored.

## What The Run Does

- Uses Red Onion's existing Tuesday-Sunday operating week. Mondays are closed and excluded from weekly public snapshots.
- Stops if `01 Daily Reports - Drop Here` is empty.
- Stops if the drop folder contains files from more than one Tuesday-Sunday operating week and lists the files to fix.
- Creates or replaces:
  - `02 Finished Reports\Check_Wine_RVA<week_end>.xlsx`
  - `02 Finished Reports\Check_Wine_VB<week_end>.xlsx`
  - `02 Finished Reports\Red_Onion_Server_Master.xlsx`
- Preserves a hashed copy of each generated workbook under:

```text
03 Archive\generated-workbooks\week-ending-YYYY-MM-DD\<run-id>\
```

- Writes the corresponding JSON integrity record under:

```text
03 Archive\run-manifests\<timestamp>-<kind>-<run-id>.json
```

- The master workbook opens to a management dashboard with current KPIs, store and group trends, selective rising/falling stars, and prioritized follow-up actions.
- Server actions require a credible sample: at least 25 guests, 3 active days, 2 prior full weeks, and 50 prior-period guests.
- Partial weeks stay visible in Data Quality but are excluded from management baselines and prominent server actions.
- Management can assign action owners, due dates, status, and notes in the master workbook. Those fields carry forward on the next successful run.
- Moves successfully processed source files to:

```text
03 Archive\processed-daily-reports\week-ending-YYYY-MM-DD\
```

The master workbook reads historical daily reports only from the canonical `processed-daily-reports` folder plus active files in the drop folder. Semantically identical same-date reports are counted once even when their file formats or bytes differ. Conflicting same-date reports stop the run before workbooks are created or active files are moved.

Source workbooks are read without rewriting their contents; after a successful run, the original files are moved into the canonical processed archive. The manifest records source and generated-file hashes for that run and links to the preceding manifest by its root-relative path and hash. The launcher also pins the exact latest manifest path and hash outside Dropbox under `%LOCALAPPDATA%\RedOnionMetrics\integrity-anchors`. Each legitimate run advances that machine-local trusted head only after the new Dropbox manifest and managed data verify successfully. Rewriting both raw data and its manifest, or deleting the Dropbox chain and requesting another baseline, therefore fails against the independently stored head.

The trusted-head file inherits the runner account's local profile permissions and must be backed up with the maintainer's independent recovery material. It is a tamper-evidence control, not encryption or immutable storage, and it does not protect against someone who already controls the automation runner's Windows account. There is no command-line reset operation. If the anchor is missing, the archive moves to another path or machine, or the pinned manifest cannot be restored, stop and have a technical maintainer reconcile the chain and raw data before explicitly adopting any existing verified history.

Before the first protected weekly run, a maintainer must explicitly confirm the current archive and finished reports as the starting state. This creates or verifies a clearly labeled integrity baseline without generating reports or moving intake files:

```powershell
& '.\Run Weekly Snapshot.cmd' -InitializeIntegrityBaseline
```

An ordinary run fails closed if that baseline, its manifest history, or the machine-local trusted head is missing; it never silently establishes a replacement baseline. On the first run after upgrading an older deployment that already has manifests but no local anchor, the same explicit initialization command verifies the complete current state and adopts that existing head once. Each later run verifies the trusted head, full manifest chain, canonical raw archive, generated-workbook archive, published public workbooks, and the master workbook's generated-content digest before reading prior management state. Approved scalar values in blue cells and the Action Board remain editable, while their styles, validation, hyperlinks, and protection metadata stay covered; formulas or external links in editable cells and changes to generated content stop the run.

## One-Time History Migration

Legacy backup folders can be copied into the canonical history without moving or deleting the originals. The command validates every candidate first, copies one representative per business date into the correct week-ending folder, and is safe to repeat:

```powershell
python _program\red_onion_weekly_metrics.py `
  --migrate-history-only `
  --migrate-history-from "..\03 Archive\pre-codex-restructure-20260706" `
  --migrate-history-from "..\03 Archive\pre-dropbox-sync-20260706"
```

If same-date files contain different metric data, migration stops before copying anything and lists the files to reconcile.

## If Something Fails

If parsing or workbook creation fails, the source files stay in `01 Daily Reports - Drop Here` so they can be fixed and rerun.

If the window identifies a file containing `No Data Available`, replace that file with a complete Toast Daily Report export. The program does not create a partial workbook or archive any current files when an input report is invalid.

Close any open output workbook before rerunning. Excel can block replacement while a workbook is open.

If the master workbook cannot be read or replaced, the run stops before moving the daily reports. Existing targets and action notes are never intentionally discarded.

If the launcher reports `Release preflight failed`, do not alter files to make the warning disappear. Ask the technical maintainer to restore the deployed repository to a clean `main` checkout whose `HEAD` matches local `origin/main`. The preflight does not download updates.

If Python is missing, install Python 3.9 or newer and select `Add python.exe to PATH`.

## Configuration

Configuration lives in `_program\red_onion_config.json`. It controls location short codes, confidence rules, scoring thresholds, materiality thresholds, display aliases, and exclusions. Only the technical maintainer should edit this file or anything else in `Red Onion Weekly Metrics Automation`.

The master workbook also contains a `Management Setup` sheet. Blue target cells in columns `B:G` hold optional store/group targets; blank targets use the rolling baseline of up to four prior full Tuesday-Sunday weeks. The visible Owner Roster begins at `A20` with `Owner Name` and `Active` columns. Add a manager once and mark departed managers inactive instead of deleting history. The `Action Board` Owner dropdown reflects active roster names immediately in Excel; the next successful run carries the roster forward and flags assignments that are inactive or no longer listed.

## Protected Edit Surface

Workbook protection is designed to prevent accidental changes while preserving the management workflow:

- `Management Setup`: configured-entity target cells in `B:G` and the Owner Roster input table are editable.
- `Action Board`: only Status (`D`), Owner (`E`), Due Date (`F`), and Manager Notes (`N`) data cells are editable.
- Other workbook cells are locked, technical sheets are `veryHidden`, and workbook structure is protected.

This protection is an operational guardrail, not encryption or a security boundary. Anyone who can download a workbook can keep a separate copy, and an authorized Dropbox editor can replace files. Use the manifests and restricted Dropbox roles to detect or reduce inappropriate changes.

## Dropbox Access And Recovery

For a personal Dropbox plan, split access by responsibility rather than giving every manager edit access to the whole operator root:

- The stable folder owner and designated technical maintainer control membership and retain edit access to the automation, archive, and integrity records.
- The weekly runner may edit the intake folder and run the launcher. Other submitters should use a Dropbox File Request where practical so they can upload without browsing or changing prior raw reports.
- Report consumers receive view-only access to `02 Finished Reports`; grant archive access only when their role requires it.
- Do not share the full parent folder with edit rights merely to make intake convenient. Dropbox editors can add, edit, delete, share, and download content; view-only users can still download or share copies.

Enable Dropbox two-factor authentication for every account with access, store recovery codes securely, set folder membership management to the owner, and review access whenever managers change. Dropbox version history provides a plan-dependent recovery window, but sync propagates changes and deletions. A separately administered, independently retained backup outside the live synced Dropbox tree is still required; choosing its owner, schedule, retention, and restore-test cadence is an operational follow-up and is not performed by this repository.

Official references: [Dropbox sharing permissions](https://help.dropbox.com/share/set-file-folder-permissions), [file requests](https://help.dropbox.com/share/create-file-request), [version history](https://help.dropbox.com/delete-restore/version-history-overview), and [two-factor authentication](https://help.dropbox.com/account-access/enable-2-factor-authentication).

## Standalone Use

The tracked `Run Weekly Snapshot.cmd` also works when this repository is copied outside the named `Red Onion Weekly Metrics Automation` deployment folder. In that standalone layout, the numbered runtime folders are created in the repository root and a `.git` directory is not required. The deployed-release Git preflight is deliberately limited to the canonical named deployment; standalone mode does not claim that its code is a verified release.

The management tabs are shown first. Raw daily, weekly, ranking, and calculation tabs remain in the workbook as `veryHidden` technical sheets for auditability and chart sources.

## What Is Not Redundant

- `Run Weekly Snapshot.cmd` is the operator launcher to double-click.
- `Red Onion Weekly Metrics Automation\_program\Run-WeeklySnapshot.ps1` is the internal runner used by the root launcher.
- `Red Onion Weekly Metrics Automation\_program\red_onion_weekly_metrics.py` is the workbook-building program.
