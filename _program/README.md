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

Runtime dependencies are exactly pinned in both `requirements.txt` and `pyproject.toml` so repeated launcher installs use the same compatible direct versions. Change the two files together and validate all supported Python versions when deliberately upgrading a pin.

## Validation

Run these checks before committing code changes:

```powershell
python -m pytest -q
python -m py_compile red_onion_weekly_metrics.py
```

Launcher integrity behavior can be checked separately with:

```powershell
python -m pytest -q tests\test_launcher_integrity.py
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

Each successful run also creates tamper-detection evidence, not immutable storage:

- `03 Archive\generated-workbooks\week-ending-YYYY-MM-DD\<run-id>\` contains hashed copies of that run's generated workbooks.
- `03 Archive\run-manifests\<timestamp>-<kind>-<run-id>.json` records source and generated-file hashes and links to the prior manifest by root-relative path and hash.

Hash verification detects a later change. It cannot prevent an editor from changing or deleting Dropbox content, so these artifacts do not replace restricted folder permissions, version history, or an independently retained backup.

A maintainer must explicitly initialize or verify the starting state before the first protected run, using `.\Run-WeeklySnapshot.ps1 -InitializeIntegrityBaseline` (or `python red_onion_weekly_metrics.py --initialize-integrity-baseline`). Ordinary runs fail closed when the baseline or manifest history is missing and never silently replace it. Subsequent runs fail before workbook generation if the chain, archived raw inputs, archived generated workbooks, published public reports, or protected generated portions of the master no longer match. The weekly publish is staged and run-locked; exact captured source bytes are used for calculation and archiving, and active source files are quarantined and deleted only after verified archive copies, final outputs, the generated-workbook snapshot, and the new manifest are committed.

## Deployed Release Preflight

When the repository folder is named exactly `Red Onion Weekly Metrics Automation`, `Run-WeeklySnapshot.ps1` fails closed unless all of these local checks pass before any runtime mutation:

- Git is available and the folder is a work tree.
- `git status --porcelain=v1 --untracked-files=all` is empty.
- The checkout is attached to branch `main`.
- `HEAD` equals the existing local `refs/remotes/origin/main` commit.

The launcher does not fetch, pull, reset, discard files, or contact an external service. Release deployment remains a maintainer operation. A non-Git standalone copy remains supported when its repository folder does not use the canonical deployment name; its numbered runtime folders live inside that standalone root.

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

The supported management edit surface is intentionally narrow:

- `Management Setup` configured-entity target cells in columns `B:G`.
- The visible Owner Roster table beginning at `A20:B20`, with `Owner Name` and `Active` fields. Add new managers and mark departing managers inactive so historical assignments remain readable.
- `Action Board` data cells for Status (`D`), Owner (`E`), Due Date (`F`), and Manager Notes (`N`).

All other cells are locked, technical sheets are `veryHidden`, and workbook structure is protected. This guards against accidental manipulation but is not encryption; Dropbox access and the manifest/backup controls remain necessary.

Technical calculation and raw-detail sheets remain in the workbook as `veryHidden` sheets. Do not delete them from the generator; they provide auditability and chart sources.

## Maintenance Notes

- Keep dependencies small and listed in `pyproject.toml` and `requirements.txt`.
- Do not add external services or credentials for the weekly run.
- Do not commit customer-facing workbooks, Toast source files, or archive contents unless explicitly requested.
- Keep the deployed Dropbox checkout clean, on `main`, and aligned with its local `origin/main`; never add a launcher bypass for the release preflight.
- On personal Dropbox plans, reserve edit access to the automation, archive, and manifests for the stable owner/technical maintainer. Give weekly submitters only the intake access they require and report consumers view-only access to finished reports.
- Require two-factor authentication and preserve Dropbox version history, while maintaining a separate independently retained backup and a documented restore test outside this repository.
