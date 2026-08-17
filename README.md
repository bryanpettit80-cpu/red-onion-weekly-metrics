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

- `01 Daily Reports - Drop Here`: drop the current Red Onion daily `.xls` or `.xlsx` files here.
- `02 Finished Reports`: generated weekly workbooks appear here.
- `03 Archive`: processed reports, prior layouts, and historical workbooks are preserved here.
- `Red Onion Weekly Metrics Automation`: Git repository, code, config, tests, and technical documentation. Operators should not edit this folder.
- `Run Weekly Snapshot.cmd`: the only root file operators need to run.

## Weekly Run

1. Save the current Red Onion Marketing Vitals files into `01 Daily Reports - Drop Here`.
2. Double-click `Run Weekly Snapshot.cmd`.
3. Open the generated workbooks in `02 Finished Reports`.

In the named Dropbox deployment, the launcher first verifies that `Red Onion Weekly Metrics Automation` is a clean Git checkout on `main`, that `HEAD` matches the checkout's local `origin/main` reference, and that `_program` contains no ignored `.pyc` or `.pyo` files that could bypass Git review. Python then runs in isolated mode with a source-program import guard that rejects sourceless bytecode even if it appears after the preflight. A failed release check stops before the run creates folders, installs packages, builds workbooks, or moves source files. The check is intentionally local and does not fetch from GitHub; a technical maintainer must update and verify the deployed checkout as a separate release step.

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

- Writes atomic attempt/status records for successful and failed runs:

```text
03 Archive\run-attempts\<timestamp>-attempt-<run-id>.json
02 Finished Reports\LAST RUN STATUS.txt
```

- The master workbook opens to a protected `How to Use` guide with the weekly workflow, signal meanings, required evidence review, editable-field rules, prohibited uses, and links to every visible sheet.
- Every visible sheet has one consistent menu link back to the protected `How to Use` workbook map.
- `Performance Dashboard` opens with an all-stores card sourced from location totals for the latest complete week, including Gross Sales, Guests, Check Average, Wine Mix, and comparisons with the preceding four complete weeks. It then summarizes the latest eight globally complete weeks for the current non-excluded roster. `Server Scorecards` separates peer-adjusted selling outcomes, consistency, and data sufficiency; `Weekly Performance` shows the dated pattern; and `Methodology` defines every calculation and limitation. These read-only views cannot create, strengthen, persist, or escalate a people-review prompt.
- `Shared & Area Trends` reports leading four-digit shared POS identities such as 5050 and 7070 separately from people. It also shows guest-weighted weekly Check Average (`Gross Sales / Guests`) for Bar, Patio, Dining Room, Banquets, and Wine Dinners at each store and for All Stores.
- `Management Center` consolidates the operator workflow that previously used separate `Action Board`, `Action History`, `Data Quality`, and `Management Setup` tabs. It shows a data-readiness summary, editable targets and owner roster, the current execution queue, and locked action history on one visible sheet.
- Detailed `Data Quality`, `Evidence Detail`, and `Run Notes` records remain in protected `veryHidden` audit/support sheets. Legacy presentation sheets such as `Team Trends`, `Store & Group Scorecards`, and `Dashboard` are also `veryHidden`; the focused analytics tabs replace them in the operator-facing layout.
- Person-level signals require at least 25 guests, 3 active days, 2 prior full self-weeks, and 50 prior-period guests.
- Partial weeks are surfaced in the `Management Center` readiness summary and retained in detailed `Data Quality`; they are excluded from management baselines and prominent server actions.
- Management can record Status, Owner, Due Date, Context Notes, Review Disposition, Reviewed By, and Review Date in the master workbook. Those fields carry forward on the next successful run.
- Moves successfully processed source files to:

```text
03 Archive\processed-daily-reports\week-ending-YYYY-MM-DD\
```

`LAST RUN STATUS.txt` separates run-scoped verification from external
assurance. `Local publication: Ready` means the exact managed per-location
workbook bytes and the protected generated content of the master workbook in
the configured `02 Finished Reports` folder match the committed manifest. It
does not cover approved editable master values, `LAST RUN STATUS.txt`, Dropbox
cloud synchronization, or recipient access. Independent recovery remains
`ExternalCheckRequired` until the current private backup and restore-test
evidence are verified outside the weekly runner.

The master workbook reads historical daily reports only from the canonical `processed-daily-reports` folder plus active files in the drop folder. Semantically identical same-date reports are counted once even when their file formats or bytes differ. Conflicting same-date reports stop the run before workbooks are created or active files are moved.

Source workbooks are read without rewriting their contents; after a successful run, the original files are moved into the canonical processed archive. The manifest records source and generated-file hashes for that run and links to the preceding manifest by its root-relative path and hash. The launcher also pins the exact latest manifest path and hash outside Dropbox under `%LOCALAPPDATA%\RedOnionMetrics\integrity-anchors`. Each legitimate run advances that machine-local trusted head only after the new Dropbox manifest and managed data verify successfully. Rewriting both raw data and its manifest, or deleting the Dropbox chain and requesting another baseline, therefore fails against the independently stored head.

The trusted-head file inherits the runner account's local profile permissions and must be backed up with the maintainer's independent recovery material. It is a tamper-evidence control, not encryption or immutable storage, and it does not protect against someone who already controls the automation runner's Windows account. There is no command-line reset operation. If the anchor is missing, the archive moves to another path or machine, or the pinned manifest cannot be restored, stop and have a technical maintainer reconcile the chain and raw data before explicitly adopting any existing verified history.

Before the first protected weekly run, a maintainer must explicitly confirm the current archive and finished reports as the starting state. This creates or verifies a clearly labeled integrity baseline without generating reports or moving intake files:

```powershell
& '.\Run Weekly Snapshot.cmd' -InitializeIntegrityBaseline
```

An ordinary run fails closed if that baseline, its manifest history, or the machine-local trusted head is missing; it never silently establishes a replacement baseline. On the first run after upgrading an older deployment that already has manifests but no local anchor, the same explicit initialization command verifies the complete current state and adopts that existing head once. A manifest-pinned owner-roster workbook from the immediately preceding supported layout, protection contract, or digest contract is accepted only through the explicit compatibility path; it is not silently reinterpreted, history-only migration remains blocked, and the next successful run regenerates it with the current layout and strict controls. Each later run verifies the trusted head, full manifest chain, canonical raw archive, generated-workbook archive, published public workbooks, and the master workbook's substantive-content digest before reading prior management state. Approved scalar values in the blue `Management Center` input cells remain editable. The v4 substantive digest fails closed for generated values and formulas, sheet and table schema, validation and editability boundaries, material cell styles, chart definitions and source bindings, meaningful worksheet-view settings, row and column sizes, internal-link destinations, drawings, external links, and protection controls. To remain stable across a no-edit Excel save, the generator materializes Excel's known persisted row heights and the digest intentionally normalizes only enumerated serializer defaults such as an explicit-versus-implicit `A1` scroll origin and chart caches derived from protected source formulas and cells. Row heights and column widths otherwise remain exact because even a small or cumulative change can alter rendered layout. Those serializer-sensitive details remain covered by the companion metadata-rich digest recorded in the manifest; a difference is surfaced as `metadata_drift=true` with a Ready warning rather than a substantive-integrity failure. Treat that warning as a request to review or regenerate the workbook, not as proof that every metadata change was harmless. Legacy v2 metadata-rich digests and the immediately preceding v3 substantive digest are never reinterpreted from the live workbook alone; their one-way compatibility bridge requires the exact manifest-inventoried archived master before the next successful run records the v4 contract.

## Workbook Layers And Metric Use

The workbook keeps three purposes separate:

- **Operations layer:** shows volume, trends, store and person comparisons, and
  descriptive context. Check Count supports total checks, Sales/Check, and
  Guests/Check. Rate of Sale and Ticket Time remain visible here.
- **Descriptive performance/consistency layer (`2026.08-v1`):** summarizes
  Sales/Guest and Wine Percentage across the latest eight globally complete
  weeks using current-roster, same-store, same-week peer medians. Its
  `Confidence` label means data sufficiency only; its labels are not statistical
  confidence, total-performance ratings, or employment decisions.
- **People-review layer (`2026.07-v3`):** may create a Context Review, Coaching Prompt, or
  Recognition Prompt using only Sales/Guest (the existing internal
  `check_average` field) and Wine Percentage. Rate of Sale, Ticket Time, Check
  Count, and the derived per-check measures cannot create, strengthen, persist,
  or escalate a person-level action.

Rate of Sale is Red Onion's inverse conversion measure:
`opportunities / qualifying sales`, so a lower positive value is better. When
the current field is Rate of Sale by Guest Count, Guests are the opportunity
count. When positive row-level rates are combined, the workbook uses the
opportunity-weighted harmonic calculation, equivalent to total opportunities
divided by reconstructed total qualifying sales. A zero, negative, malformed,
or missing rate cannot be safely reconstructed from the ratio alone and is
unavailable for the combined context value.

Ticket Time is check-weighted only when every contributing row has a valid
Check Count and the total Check Count is positive:
`sum(Ticket Time × Check Count) / sum(Check Count)`. If Check Count coverage is
incomplete, the workbook does not substitute guest weighting; the combined
Ticket Time is unavailable. With complete coverage and a positive total Check
Count, it also reports `Sales/Check = total sales / total checks` and
`Guests/Check = total guests / total checks`.

The operational layer is intended to help a manager ask:

- What changed, and is it isolated or store-wide?
- Did check volume, Guests/Check, or another observable operating condition
  move at the same time?
- Is the source complete and the work reasonably comparable?
- What should be verified before a coaching or recognition conversation?

### Descriptive Performance And Consistency

A server-week qualifies for `2026.08-v1` with at least 25 guests, at least
three active days, and at least five other qualified current-roster peers in
the same store and week. Weekly peer references are medians. Eight-week peer
gaps are weighted by `min(guests, 50)` so one large week cannot dominate.

`High` data sufficiency requires at least six qualified weeks and 200 qualified
guests; `Provisional` requires at least four weeks and 150 guests; otherwise it
is `Insufficient`. Consistency is the sample standard deviation of qualified
weekly peer gaps. High consistency requires Sales/Guest SD no greater than
$11.50 and Wine-gap SD no greater than 4.1 percentage points; Moderate requires
no greater than $17.50 and 5.7 points. Recent movement is the capped-weighted
recent four-week peer gap minus the preceding four-week gap and remains
descriptive. Missing and unqualified weeks are never treated as zero.

The runner recomputes this protected snapshot from verified source data during
each successful generation. It is not a live Excel recalculation model, and it
does not feed the people-review action path.

Shared POS and area trends use the same latest eight globally complete weeks.
Shared-number rows never enter person peer comparisons or people-review actions.
Bar, Patio, Banquets, and Wine Dinners use configured source-name patterns;
Dining Room is the residual set of otherwise eligible named servers. Wine
Dinners remains visibly unavailable until an actual POS label or shared number
is configured in `weekly_area_name_patterns` or `weekly_shared_number_areas`.

## `2026.07-v3` Coaching-Signal Interpretation

Methodology `2026.07-v3` is a deterministic, rule-based screening aid. It does
not estimate statistical confidence, predict future performance, establish
causality, or determine whether a person is performing fairly relative to
unobserved shift conditions. Its output is limited to coaching and recognition
review prompts.

The workbook uses the following visible terms:

- `Recent Movement`: the current complete week compared with up to four prior
  complete weeks for the same person.
- `Peer Comparison`: the current week compared with a leave-one-person-out,
  same-store median from the prior four complete weeks.
- `Evidence Status`: whether the sample, peer reference, persistence, and
  stability requirements are satisfied. It is not statistical confidence.
- `Context Review`: a first-week or context-sensitive signal that requires
  investigation before any coaching or recognition conclusion.
- `Coaching Prompt` and `Recognition Prompt`: a second consecutive qualified
  signal with a recurring metric driver that remains stable when any one active
  day is removed.

The peer reference includes only qualified, non-excluded server-weeks. It
requires at least three usable prior weeks, five distinct peers in each usable
week, and 20 peer-week observations. If those requirements are not met, the
workbook displays `Reference Unavailable` and does not issue a coaching or
recognition prompt. Management targets remain visible as business context, but
they do not drive person-level prompts. Rank, when shown, is descriptive only
and cannot change an action classification.

A gap, incomplete or low-volume week, unavailable peer reference, changed
signal direction, or day-sensitive result prevents escalation. The report also
checks for common store-wide movement before attributing a signal to one
person. Every management-facing signal sheet states:

> Rule-based observational coaching signal—not a statistical, causal, or
> employment decision. Verify comparable work context and source accuracy.

These outputs must never be the sole or determinative basis for pay,
scheduling, discipline, promotion, or termination. See
[MODEL_CARD.md](MODEL_CARD.md) and [DATA_GOVERNANCE.md](DATA_GOVERNANCE.md).

## History Migration And Rebuild

Legacy backup folders can be copied into the canonical history without moving or deleting the originals. The command validates every candidate first, copies one representative per business date into the correct week-ending folder, and is safe to repeat:

```powershell
python _program\red_onion_weekly_metrics.py `
  --migrate-history-only `
  --migrate-history-from "..\03 Archive\pre-codex-restructure-20260706" `
  --migrate-history-from "..\03 Archive\pre-dropbox-sync-20260706"
```

If same-date files contain different metric data, migration stops before copying anything and lists the files to reconcile.

After a validated migration, a technical maintainer can rebuild the managed
outputs entirely from canonical history:

```powershell
python _program\red_onion_weekly_metrics.py `
  --rebuild-from-history
```

`--rebuild-from-history` is a maintainer operation, not part of the normal
weekly launcher. It verifies the current manifest, trusted head, master
workbook, and canonical history before publication. It does not process the
drop folder or move active weekly inputs. The rebuild preserves management
review fields by header and publishes the rebuilt outputs, generated-workbook
archive, manifest, and trusted head as one protected transaction. A validation,
conflict, or write failure leaves the preexisting managed state unchanged.

Migration and rebuilding may be combined when importing a separately staged
history folder:

```powershell
python _program\red_onion_weekly_metrics.py `
  --rebuild-from-history `
  --migrate-history-from "C:\approved\red-onion-history"
```

### One-Time Gmail Backfill

The 16-week candidate calibration uses a one-time, read-only retrieval of 24
original `Daily Report - TM` attachments from the Marketing Vitals sender.
They add four complete Tuesday-Sunday weeks before the canonical archive. This
is a controlled backfill, not an ongoing Gmail integration:

- Select reports by their embedded business date; the attachment filename and
  email date may be one day later.
- Accept only the original TM report family. Exclude derived `Check_Wine`
  workbooks, Store reports, forwarded duplicates, Monday reports, unrelated
  messages, `No Data Available` workbooks, legacy incompatible schemas, and
  conflicting same-date files.
- Stage attachments outside the Git repository and live Dropbox operator
  folders. Do not retain email bodies or message identifiers.
- Validate the expected workbook schema, six Tuesday-Sunday dates per week,
  configured locations, metric values, and daily guest/sales reconciliation
  before migration.
- Use the protected history migration/rebuild transaction above. After the
  canonical archive and outputs verify, remove the temporary staging copies.

The repository contains no Gmail credentials, background mail reader, or
recurring email connector. A normal weekly run never accesses Gmail.

Maintainers can reproduce the anonymized calibration and replay diagnostics
without writing files:

```powershell
python _program\red_onion_model_validation.py `
  "C:\approved\canonical-history" `
  "C:\approved\temporary-backfill" `
  --start 2026-03-24 `
  --end 2026-07-19
```

The command prints JSON only. It reports aggregate observation counts,
thresholds, review rates, prompt stability, and reversal rates; it does not
emit person-level rows.

## If Something Fails

If parsing or workbook creation fails, the source files stay in `01 Daily Reports - Drop Here` so they can be fixed and rerun.

If the window identifies a file containing `No Data Available`, replace that file with a complete Red Onion Daily Report export. The program does not create a partial workbook or archive any current files when an input report is invalid.

Close any open output workbook before rerunning. Excel can block replacement while a workbook is open.

If the master workbook cannot be read or replaced, the run stops before moving the daily reports. Existing targets and action notes are never intentionally discarded.

If the launcher reports `Release preflight failed`, do not alter files to make the warning disappear. Ask the technical maintainer to restore the deployed repository to a clean `main` checkout whose `HEAD` matches local `origin/main`. The preflight does not download updates.

If Python is missing, install Python 3.10-3.12 and select `Add python.exe to PATH`.

A normal weekly run verifies the existing local environment and does not
reinstall packages. Technical maintainers can explicitly repair it with
`-RebuildEnvironment -HealthCheck`, or run a read-only local check with
`-HealthCheck`. The combined maintenance command rebuilds and validates without
processing weekly reports.

## Configuration

Configuration lives in `_program\red_onion_config.json`. It controls location
short codes, sample and peer-reference rules, calibrated scoring thresholds,
persistence and stability rules, materiality thresholds, display aliases, and
exclusions. Only the technical maintainer should edit this file or anything
else in `Red Onion Weekly Metrics Automation`.

The visible `Management Center` contains the optional store/group targets and
owner roster. The `ManagementTargets` table spans `C:I`: Entity in column `C`
is locked, while the blue target cells in `D:I` are editable. Targets and
rolling store baselines remain store-level business context; person-level
prompts use the qualified same-store peer reference described above. The Owner
Roster occupies `K:L`, with editable `Owner Name` and `Active` fields. Add a
manager once and mark departed managers inactive instead of deleting history.
The Current Actions Owner dropdown reflects active roster names immediately in
Excel; the next successful run carries the roster forward and flags assignments
that are inactive or no longer listed.

The frozen people-review thresholds identify their calibration dataset, date,
quantile method, rounding rules, and methodology version. `2026.07-v3` uses
only Sales/Guest and Wine Percentage. It uses the larger of the business
minimum and the R-7 75th percentile of absolute qualified deviations for a
neutral band, and the larger of the business minimum and the R-7 90th
percentile for a strong band. Values are rounded half-up to $0.50 for
Sales/Guest and 0.001 for Wine Percentage. Thresholds are frozen for a
methodology release and are not recomputed during an ordinary weekly run.
The descriptive `2026.08-v1` layer reuses the configured Sales/Guest and Wine
movement neutral/strong bands as its performance and High/Moderate consistency
boundaries. Its weekly run is deterministic and does not recalibrate them.

Rate of Sale, Ticket Time, Check Count, Sales/Check, and Guests/Check have no
people-review scoring thresholds. They remain descriptive operating context
under the aggregation and completeness rules above.

## Protected Edit Surface

Workbook protection is designed to prevent accidental changes while preserving the management workflow:

- `Management Center`: only configured-entity target cells in `D:I` (Entity in
  `C` stays locked), Owner Roster cells in `K:L`, and Current Actions cells for
  Status (`D`), Owner (`E`), Due Date (`F`), Context Notes (`N`), Review
  Disposition (`U`), Reviewed By (`V`), and Review Date (`W`) are editable.
  Its readiness summary and Action History section are locked.
- `Performance Dashboard`, `Server Scorecards`, `Weekly Performance`, `Shared &
  Area Trends`, and `Methodology` are protected and read-only by default.
  `_Consistency Calc` is protected and `veryHidden`.
- Other workbook cells are locked. Detailed `Data Quality`, `Evidence Detail`,
  and `Run Notes`, legacy presentation sheets, and technical calculation/raw
  sheets are `veryHidden`; workbook structure is protected.

Generated rows begin with `Pending Review`. A completed disposition requires a
reviewer and review date. Valid dispositions are `Coaching Accepted`,
`Recognition Accepted`, `Context Explains`, `Data Issue`, and `Monitor`.
Invalid or incomplete combinations remain visibly pending and are rejected by
the next protected run.

This protection is an operational guardrail, not encryption or a security boundary. Authorized users can choose **Review > Unprotect Sheet** (or unprotect workbook structure) with the documented lowercase password `redonion`. Use a separate copy for exploratory edits: saving generated values or formulas, sheet structure, visibility, material styles, chart definitions or bindings, meaningful worksheet-view settings, row or column sizes, editability, or protection into the managed master changes its substantive digest and can stop the next protected run. Enumerated Excel serializer defaults and derived chart-cache changes can instead surface as a non-blocking metadata-drift warning; review or regenerate the workbook whenever that warning appears. Anyone who can download a workbook can keep a separate copy, and an authorized Dropbox editor can replace files. Use the manifests and restricted Dropbox roles to detect or reduce inappropriate changes.

## Dropbox Access And Recovery

For a personal Dropbox plan, split access by responsibility rather than giving every manager edit access to the whole operator root:

- The stable folder owner and designated technical maintainer control membership and retain edit access to the automation, archive, and integrity records.
- The weekly runner may edit the intake folder and run the launcher. Other submitters should use a Dropbox File Request where practical so they can upload without browsing or changing prior raw reports.
- Report consumers receive view-only access to `02 Finished Reports`; grant archive access only when their role requires it.
- Do not share the full parent folder with edit rights merely to make intake convenient. Dropbox editors can add, edit, delete, share, and download content; view-only users can still download or share copies.

Enable Dropbox two-factor authentication for every account with access, store recovery codes securely, set folder membership management to the owner, and review access whenever managers change. Dropbox version history provides a plan-dependent recovery window, but sync propagates changes and deletions. A separately administered, independently retained backup outside the live synced Dropbox tree is still required. `Build-RecoveryBundle.ps1` creates a verified local release/recovery bundle for that private destination; it never uploads automatically. Retain 13 weekly and 12 monthly bundles and perform a documented restore test quarterly.

Official references: [Dropbox sharing permissions](https://help.dropbox.com/share/set-file-folder-permissions), [file requests](https://help.dropbox.com/share/create-file-request), [version history](https://help.dropbox.com/delete-restore/version-history-overview), and [two-factor authentication](https://help.dropbox.com/account-access/enable-2-factor-authentication).

Maintainers should also follow [MAINTAINER.md](MAINTAINER.md),
[RECOVERY.md](RECOVERY.md), [INCIDENT_RESPONSE.md](INCIDENT_RESPONSE.md), and
[DATA_GOVERNANCE.md](DATA_GOVERNANCE.md). The analytical-use contract and
validation limits are documented in [MODEL_CARD.md](MODEL_CARD.md).

## Standalone Use

The tracked `Run Weekly Snapshot.cmd` also works when this repository is copied outside the named `Red Onion Weekly Metrics Automation` deployment folder. In that standalone layout, the numbered runtime folders are created in the repository root and a `.git` directory is not required. The deployed-release Git preflight is deliberately limited to the canonical named deployment; standalone mode does not claim that its code is a verified release.

The operator-facing workbook has exactly seven visible tabs: `How to Use`,
`Performance Dashboard`, `Server Scorecards`, `Weekly Performance`, `Shared &
Area Trends`, `Methodology`, and `Management Center`. Detailed audit/support,
legacy presentation, raw daily/weekly/ranking, and calculation tabs—including
`Data Quality`, `Evidence Detail`, `Run Notes`, and `_Consistency Calc`—remain
`veryHidden` for auditability and chart sources.

## What Is Not Redundant

- `Run Weekly Snapshot.cmd` is the operator launcher to double-click.
- `Red Onion Weekly Metrics Automation\_program\Run-WeeklySnapshot.ps1` is the internal runner used by the root launcher.
- `Red Onion Weekly Metrics Automation\_program\red_onion_weekly_metrics.py` is the workbook-building program.
