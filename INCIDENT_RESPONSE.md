# Incident response

Preserve the intake, console error, `LAST RUN STATUS.txt`, and relevant
`03 Archive\Reporting State` attempt, receipt, and journal before changing
files. Read the outcome first: `published_with_warnings` means the report
committed even if cleanup or a status refresh needs attention.

## Missing or incorrect daily reports

Leave intake available for correction. Confirm the embedded business date,
both restaurant locations, and the listed missing restaurant-days. Replace
invalid exports from the authoritative reporting source. An actual closure
needs an explicit verified `closed_dates` entry; do not assume a missing day
was closed.

Identical business facts are safe duplicates. For changed facts, run
`-ReviewCorrections`, inspect the displayed differences, and choose once per
date. Do not delete accepted history or rename a conflicting report to hide it.

## Publication or recovery conflict

Close Excel and wait for cloud files to be available. Retry only after
resolving the stated issue. The runner reconciles pending journals before
starting new reporting work.

If recovery detects a newer external edit, preserve both that file and the
previous-good copy identified by the journal. Reconcile the business inputs
before choosing the intended version. Do not remove Reporting State, edit
receipt hashes, or reset a baseline to silence an error.

If state itself is damaged, restore its matching backup as described in
[RECOVERY.md](RECOVERY.md). Historical legacy-manifest problems should be
investigated in a preserved legacy copy; they are not a reason to reconstruct
machine-bound trust for portable reporting.
Do not initialize a replacement baseline to conceal a legacy discrepancy.

## Management workbook conflict

A filename change or formatting save is allowed. Required tables, columns,
Action IDs, and valid input values must remain readable. If multiple copies
contain different edits, preserve them and select the intended copy with
`-ManagementWorkbook`. Missing/duplicate Action IDs or invalid dispositions
must be corrected explicitly.

The committed snapshot can recover management state saved at the last run.
It cannot recover later workbook edits without that newer workbook or backup.

## Source or environment failure

Review a source change through normal repository maintenance. Ordinary runs
do not require a clean branch or `HEAD == origin/main`.

For a missing/stale Python environment, use
`-RebuildEnvironment -HealthCheck`. This repairs local dependencies without
processing reports. Verify the corrected source/environment on a copy when
the failure involves implementation changes.

## Employee evidence or distribution incident

Restrict access and preserve the candidate, fingerprint, approval receipt, and
known distribution facts. An approval for one candidate/purpose cannot be
reused for altered bytes or broader distribution. Document the business owner,
recipients, dates, scope, correction/deletion actions, and retention obligations.
The reporting application does not send evidence automatically.

Record the incident owner, timeline, cause, preserved evidence, correction,
validation, and decision to resume. Report what was verified and what remains
unknown.
