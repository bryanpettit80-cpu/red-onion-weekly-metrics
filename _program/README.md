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

Runtime dependencies are exactly pinned in both `requirements.txt` and
`pyproject.toml`. `requirements.lock` records the reviewed transitive graph and
artifact hashes; `requirements-constraints.txt` keeps that lock compatible with
Python 3.10-3.12.

The launcher records the Python identity and lock digest in
`%LOCALAPPDATA%\RedOnionMetrics\environment-state.json`. A normal weekly run
verifies installed versions and `pip check`; it does not invoke an installer.
A maintainer can deliberately rebuild with:

```powershell
.\Run-WeeklySnapshot.ps1 -RebuildEnvironment -HealthCheck
```

That maintenance-only combination rebuilds and validates the environment
without processing weekly reports. Omit `-HealthCheck` only when the rebuild
should be followed by the weekly run.

For a replacement-machine restore, follow the audited procedure in
`..\RECOVERY.md` and use
`-RebindRestoredIntegrityAnchor <SOURCE_ANCHOR_JSON>`. The command verifies the
restored head and managed outputs before creating the new path-bound anchor; it
does not process weekly reports.

## Validation

Run these checks before committing code changes:

```powershell
python -m pytest -q
python -m py_compile red_onion_weekly_metrics.py
python red_onion_config.py --config red_onion_config.json
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
- `03 Archive\run-attempts\<timestamp>-attempt-<run-id>.json` records each attempt, stage, safe error, and separate readiness dimensions.
- `02 Finished Reports\LAST RUN STATUS.txt` is the atomic operator summary.

Hash verification detects a later change. It cannot prevent an editor from changing or deleting Dropbox content, so these artifacts do not replace restricted folder permissions, version history, or an independently retained backup.

A maintainer must explicitly initialize or verify the starting state before the first protected run, using `.\Run-WeeklySnapshot.ps1 -InitializeIntegrityBaseline` (or `python red_onion_weekly_metrics.py --initialize-integrity-baseline`). Ordinary runs fail closed when the baseline or manifest history is missing and never silently replace it. A manifest-pinned owner-roster workbook from the immediately preceding protection contract may be adopted without rewriting it; history-only migration is blocked until the next ordinary weekly run regenerates that workbook with current strict controls. Subsequent runs fail before workbook generation if the chain, archived raw inputs, archived generated workbooks, published public reports, or protected generated portions of the master no longer match. The weekly publish is staged and run-locked; exact captured source bytes are used for calculation and archiving, and active source files are quarantined and deleted only after verified archive copies, final outputs, the generated-workbook snapshot, and the new manifest are committed.

The launcher stores the trusted manifest-head anchor outside Dropbox under `%LOCALAPPDATA%\RedOnionMetrics\integrity-anchors`. Keep access to that machine-local location restricted to the Windows account that runs the automation, and include it in the maintainer's independent recovery backup. The anchor pins the exact latest manifest path and hash and advances only after a verified manifest commit. Once present, neither an ordinary run nor `--initialize-integrity-baseline` will accept a deleted chain or a rewritten raw-data/manifest pair. There is intentionally no automated reset option; reconcile and restore trusted history before adopting an existing chain on a replacement runner.

## Deployed Release Preflight

When the repository folder is named exactly `Red Onion Weekly Metrics Automation`, `Run-WeeklySnapshot.ps1` fails closed unless all of these local checks pass before any runtime mutation:

- Git is available and the folder is a work tree.
- `git status --porcelain=v1 --untracked-files=all` is empty.
- The checkout is attached to branch `main`.
- `HEAD` equals the existing local `refs/remotes/origin/main` commit.

The launcher permits existing ignored Python bytecode but isolates each run by disabling bytecode writes and redirecting bytecode lookup to a fresh unused cache path. It does not delete those files, fetch, pull, reset, discard files, or contact an external service. Release deployment remains a maintainer operation. A non-Git standalone copy remains supported when its repository folder does not use the canonical deployment name; its numbered runtime folders live inside that standalone root.

## Operational Context And Person-Action Scope

Methodology `2026.07-v3` maintains two calculation scopes:

- `SERVER_PERSON_ACTION_FIELDS` contains only `check_average` (displayed as
  Sales/Guest) and `wine_pct`. Only these two fields may affect person-level
  Recent Movement, Peer Comparison, composite direction, persistence, or
  action classification.
- `SERVER_CONTEXT_FIELDS` contains `rate_of_sale_by_guest_count` and
  `average_ticket_time_seconds`. These fields, Check Count, Sales/Check, and
  Guests/Check remain visible descriptive context and must not affect a
  person-level action.

Red Onion defines Rate of Sale as
`opportunities / qualifying sales`; lower is better. For positive available
row-level values, the correct combined rate is the opportunity-weighted
harmonic result. The current `rate_of_sale_by_guest_count` field uses Guests as
the opportunity count:

```text
combined ROS = sum(opportunities) / sum(opportunities / row ROS)
```

Do not substitute an opportunity-weighted arithmetic mean. A nonpositive,
malformed, or missing ROS cannot be safely reconstructed from the ratio alone
and makes the combined context value unavailable.

Average Ticket Time is combined only when Check Count is valid for every
contributing row and total Check Count is positive:

```text
combined Ticket Time = sum(row Ticket Time * row Check Count)
                       / sum(row Check Count)
Sales/Check            = sum(sales) / sum(Check Count)
Guests/Check           = sum(guests) / sum(Check Count)
```

Incomplete Check Count coverage makes the combined Ticket Time and derived
per-check context unavailable; guest weighting is not a fallback.

## Coaching-Signal Interpretation

Methodology `2026.07-v3` is a deterministic, rule-based screening aid. It does
not estimate statistical confidence, predict future performance, establish
causality, or adjust for unobserved shift conditions. Its outputs are limited
to human-reviewed coaching and recognition prompts.

- `Recent Movement` compares the current complete week with up to four prior
  complete weeks for the same person.
- `Peer Comparison` compares the current week with a leave-one-person-out,
  same-store median from the prior four complete weeks.
- `Evidence Status` describes sample, peer-reference, persistence, and
  leave-one-active-day stability checks. It is not statistical confidence.
- `Context Review` is a first-week or context-sensitive signal. `Coaching
  Prompt` and `Recognition Prompt` require a second consecutive qualified
  signal with a recurring metric driver.

The peer reference includes only qualified, non-excluded server-weeks and
requires at least three usable prior weeks, five distinct peers in each usable
week, and 20 peer-week observations. If those requirements are not met, the
workbook displays `Reference Unavailable` and does not issue a coaching or
recognition prompt. Management targets remain visible as business context but
do not drive person-level prompts. Rank is descriptive only and cannot change
an action classification.

Because the people-review composite contains only Sales/Guest and Wine
Percentage, both metric families must agree and meet the configured materiality
rules before a person-level candidate exists. Operational context may suggest
questions for the manager, but it cannot supply the second agreeing metric or
become a recurring action driver.

Incomplete or low-volume weeks, unavailable peer references, gaps, changed
signal direction, common store-wide movement, and day-sensitive results prevent
escalation. The management signal sheets visibly state:

> Rule-based observational coaching signal—not a statistical, causal, or
> employment decision. Verify comparable work context and source accuracy.

These outputs must never be the sole or determinative basis for pay,
scheduling, discipline, promotion, or termination. See
[MODEL_CARD.md](../MODEL_CARD.md) and
[DATA_GOVERNANCE.md](../DATA_GOVERNANCE.md).

The management `Team Trends` sheet shows every current non-excluded server with
current sample, exact consecutive-week Sales/Guest and Wine % changes,
descriptive four-week movement, peer comparison, eight-week direction, evidence
status, gated action outcome, and exact history used. Descriptive movement is
visible even when action gates are not met; it does not create or strengthen a
person-level prompt.

`Action Board` is the single execution view and retains the seven editable
management fields. `Evidence Detail` is protected/read-only and records stable
codes, exact evidence weeks, source SHA-256/format/parser/date-source,
methodology version, and metric evidence.

`Data Quality` begins with the latest-week detail and a latest-first 16-week
location completeness matrix. The matrix labels every location-week as
Complete, Partial, or Missing before the historical exception, owner review,
and source-provenance sections.

Before recording a disposition, the manager should ask whether the source and
identity are correct, whether the work was reasonably comparable, whether
check volume or a common store condition explains the movement, and what
independent evidence supports the coaching or recognition conclusion.

`Management Setup` targets, owner names, and manual `Action Board` fields are read from the existing master before regeneration. The new workbook is written to a temporary file and atomically replaces the prior master only after validation succeeds.

The supported management edit surface is intentionally narrow:

- `Management Setup` configured-entity target cells in columns `B:G`.
- The visible Owner Roster table beginning at `A20:B20`, with `Owner Name` and `Active` fields. Add new managers and mark departing managers inactive so historical assignments remain readable.
- `Action Board` data cells for Status (`D`), Owner (`E`), Due Date (`F`),
  Context Notes (`N`), Review Disposition (`U`), Reviewed By (`V`), and Review
  Date (`W`). A generated item starts at `Review Needed`; it cannot advance
  without a non-pending disposition, reviewer, and review date.

All other cells are locked, technical sheets are `veryHidden`, and workbook structure is protected. This guards against accidental manipulation but is not encryption; Dropbox access and the manifest/backup controls remain necessary.

Technical calculation and raw-detail sheets remain in the workbook as `veryHidden` sheets. Do not delete them from the generator; they provide auditability and chart sources.

## History Migration And Rebuild

A technical maintainer can rebuild managed outputs entirely from
manifest-pinned canonical history:

```powershell
python red_onion_weekly_metrics.py --rebuild-from-history
```

The rebuild verifies the manifest chain, trusted head, master workbook, and
canonical history; ignores the active drop folder; preserves management review
fields by header; and publishes the rebuilt workbooks, generated-workbook
archive, successor manifest, and trusted head as one protected transaction.
Migration from a separately approved staging folder can be combined with the
rebuild:

```powershell
python red_onion_weekly_metrics.py `
  --rebuild-from-history `
  --migrate-history-from "C:\approved\red-onion-history"
```

Calibration and replay diagnostics are read-only and emit anonymized aggregate
JSON:

```powershell
python red_onion_model_validation.py `
  "C:\approved\canonical-history" `
  "C:\approved\temporary-backfill" `
  --start 2026-03-24 `
  --end 2026-07-19
```

## Read-Only Health Check

```powershell
.\Run-WeeklySnapshot.ps1 -HealthCheck
python red_onion_weekly_metrics.py --health-check-json
```

The command does not create folders, acquire the workflow lock, install
packages, generate reports, or access Google Drive. It verifies exact local
publication through the integrity manifest: exact managed per-location
workbook bytes plus the protected substantive content of the master workbook.
Generated values, formulas, schema, display-affecting formats, chart content
and bindings, hidden dimensions, links, drawings, editability, and protection
remain fail-closed. Proven equivalent Excel serialization defaults are reported
without blocking only when substantive content matches a manifest-inventoried
archived master. Approved editable master values, `LAST RUN STATUS.txt`,
Dropbox sync, and recipient access remain
outside that claim. Independent recovery is reported as
`ExternalCheckRequired` until the current private backup and restore-test
evidence are verified separately.

## Maintenance Notes

- Keep dependencies small and listed in `pyproject.toml` and `requirements.txt`.
- Do not add external services or credentials for the weekly run.
- Do not commit customer-facing workbooks, Red Onion source files, or archive contents unless explicitly requested.
- Keep the deployed Dropbox checkout clean, on `main`, and aligned with its local `origin/main`; never add a launcher bypass for the release preflight.
- On personal Dropbox plans, reserve edit access to the automation, archive, and manifests for the stable owner/technical maintainer. Give weekly submitters only the intake access they require and report consumers view-only access to finished reports.
- Require two-factor authentication and preserve Dropbox version history, while maintaining a separate independently retained backup and a documented restore test outside this repository.
