# Red Onion program notes

Use the parent [README](../README.md) for the operator workflow and
[SETUP](../SETUP.md) for Windows installation. This folder contains the
project's reporting rules, workbook generator, and local runtime helpers.

## Responsibilities

| Module | Responsibility |
| --- | --- |
| `red_onion_config.py` | Validate configuration without third-party dependencies. |
| `red_onion_weekly_metrics.py` | Parse reports, calculate business results, render and validate workbooks, expose CLI commands, and retain explicit legacy maintenance paths. |
| `red_onion_workflow.py` | Ordinary intake, complete-week gating, duplicate/correction handling, accepted history, management carry-forward, and result recording. |
| `red_onion_management.py` | Identify management tables and import allowed business inputs independently of workbook filename and formatting. |
| `red_onion_transactions.py` | Portable receipts, verified copies, publication journals, rollback, and post-commit cleanup. |
| `red_onion_runtime.py` | Atomic attempt/status output and run stages. |
| `red_onion_integrity.py` | Hash/canonical JSON helpers and retained legacy manifest primitives. |
| `Run-WeeklySnapshot.ps1` | Workspace resolution, dependency validation, source-only isolated Python startup, and option forwarding. |

The refactor retains the existing calculation and workbook code. The new
modules separate ordinary orchestration and recovery from the historical
manifest/anchor workflow; they do not add a mail scanner or external service.

## Environment and validation

The launcher uses Python 3.10-3.12 and a disposable environment at
`%LOCALAPPDATA%\RedOnionMetrics\.venv`. It records Python identity and
requirements-lock identity locally and validates installed versions and
`pip check`. A valid environment is reused; a missing or stale one is rebuilt
for an ordinary invocation. A plain health check never installs dependencies.

For a local operational workspace, the setup-only command is:

```powershell
.\Run-WeeklySnapshot.ps1 -RebuildEnvironment -HealthCheck
```

For source development and Linux/Cloud testing, use the isolated environment
and commands in [MAINTAINER.md](../MAINTAINER.md#development-and-tests).
The operational health check requires the local reporting folders and is not
part of a fresh source checkout's test setup.

The launcher loads local Python source directly under isolated startup.
Adjacent or cached bytecode is not an execution source. Harmless cache files
and Git working-tree state do not decide whether a reporting week can run.

## Workspace and portable state

Workspace discovery prefers an explicit `-OperationsRoot`, then numbered
folders inside the program root, then its parent. A fresh standalone copy
defaults to its own root. The outer launcher supports a renamed program folder
and stops if it finds multiple possible program subfolders.

Current ordinary-run technical data lives under:

```text
03 Archive\Reporting State\
  current.json
  attempts\<run-id>.json
  corrections\<review-id>.json
  runs\<run-id>\
    commit.json
    journal.json
    sources\
    outputs\
    previous\
```

`current.json` points to the accepted run record and its checksum.
`commit.json` contains accepted daily facts, source references, configuration,
methodology, correction receipts, management state, and publication identity.
Source and output copies support audit and recovery. `previous` contains
replaced output copies; it may be absent when nothing was replaced.
Transient staging and quarantine files are cleaned when safe.

The first portable run validates and imports canonical
`processed-daily-reports` history and readable management tables. Existing
legacy manifests and generated archives remain untouched. Later runs read the
accepted current state instead of requiring permanent byte identity for every
historical workbook. Do not hand-edit receipts to correct business facts.

## Intake and publication invariants

- A normal public report covers one full Tuesday-Sunday week at every
  configured restaurant. The two-store configuration has 12 restaurant-days;
  an explicit `closed_dates` entry can account for a verified closure.
- The embedded report date, restaurant totals, numeric values, daily
  reconciliation, duplicate resolution, and complete-week gate are business
  controls. Filenames and byte equality are not business identities.
- The input files used for calculation are captured before parsing. Their
  unchanged bytes are verified before accepted inputs leave intake.
- Conflicting business alternatives stop before publication unless one
  fact-bound choice per date has been reviewed. `--review-corrections` shows
  changes; `--corrections-file` supplies previously reviewed choices.
- A local OS lock prevents overlapping writers on one computer. Dropbox is
  used sequentially across computers, never as a distributed lock.
- Staged outputs are validated, previous outputs are saved, and the journal
  precedes publication. The receipt advances once the output set is installed.
- Recovery rolls back uncommitted publication or resumes committed cleanup.
  New external edits are preserved and require reconciliation.
- A post-commit cleanup/status error does not relabel publication as failed.
  A duplicate business submission does not regenerate the same reports.

The workbook's saved management inputs are read by table and column names.
After initial import, only targets, owners, and seven editable action fields
merge into saved state. Stable Action IDs bind edits to accepted actions;
saved action history and evidence are not replaced by edited generated cells.
Renaming or formatting the master is allowed. Ambiguous different copies
require `--management-workbook` within Finished Reports.

## Calculation scope

The detailed formulas and thresholds are in [MODEL_CARD.md](../MODEL_CARD.md).
Maintain these boundaries:

- People-review methodology `2026.07-v3` uses only Sales/Guest and Wine
  Percentage. Operational context and descriptive scorecards do not feed its
  direction, persistence, stability, or action classification.
- Descriptive methodology `2026.10-v2` includes all resolved named servers in
  the latest eight globally complete weeks. Historical peers are selected
  from their actual store/week. Main actual/peer summaries use the same
  qualified capped weights; wine gaps are percentage points. Low-data people
  remain visible with no unsupported classification.
- Shared POS identities and area aggregates remain separate operations views.
- Rate of Sale combines positive available ratios harmonically using Guests
  as opportunities. Ticket Time requires complete Check Count coverage and
  is check-weighted. Derived per-check context uses that same coverage.

Regression checks should assert business values, coverage, saved management
fields, relevant workbook layout, and recoverable transaction outcomes.
Exact hashes are appropriate for copy verification and a particular journal,
not for permanent presentation identity across Excel saves.

## Maintainer commands

Run ordinary history rebuilds and imports with explicit path options when
operating on a copy; see [MAINTAINER.md](../MAINTAINER.md). A rebuild ignores
active intake, uses accepted history, and preserves current management input.
An import-only command accepts history without generating reports and retains
the prior publication identity. The next ordinary run refreshes the master and
any complete affected weeks; incomplete history remains supporting detail.

`--health-check-json` returns local health schema version 2 for
`portable-v1`. It checks configuration, folders, and receipt readability.
It is not a complete-input readiness API or proof of cloud sync.

The old `--legacy-integrity`, baseline, and anchor-rebind commands are
maintainer-only operations on preserved legacy workspaces. They are not
required to set up, rename, restore, or run a portable workspace.

The existing evidence export remains a two-stage local candidate and exact
approval operation. It never uploads or sends. The future coordinator contract
is [design only](../COORDINATOR_CONTRACT.md).
