# Employee Data Governance

Server scorecards, named signals, coaching/recognition prompts, review
dispositions, owner assignments, evidence details, and approved evidence
exports are Restricted Employee Performance Information.

The current `2026.07-v3` methodology uses deterministic observational
screening rules. They are not a statistical, predictive, causal, or
employment-decision model. Data integrity and reproducibility do not establish
that a signal fairly attributes performance to one person.

## Allowed Use

- Weekly context review by authorized Red Onion leaders.
- Identifiable coaching and recognition conversations after source accuracy,
  comparable-work context, peer-reference sufficiency, and stability have been
  reviewed.
- Accountable tracking of the generated signal, manager disposition, reviewer,
  review date, and resulting follow-up.
- AI-assisted analysis only after a named authorized manager reviews the exact
  `ManagementEvidencePackageV2` candidate and approves its candidate and
  fingerprint hashes for a specific purpose.

The approval does not authorize a different file, a regenerated package, a
different purpose, or broader distribution. Existing V1 packages remain
readable for audit and retention, but all newly generated packages use V2.

## Prohibited Use

The signal or evidence package must never be the sole or determinative basis
for:

- discipline, termination, compensation, promotion, demotion, or scheduling;
- reducing hours, shifts, sections, opportunities, or training;
- a formal performance rating or any other adverse employment action; or
- asserting causality, statistical significance, predictive validity, or
  demographic fairness.

An authorized manager must independently corroborate any coaching or
recognition decision. The report does not observe shift, daypart, section,
party mix, staffing, event, tenure, training, or menu-availability context.
Names must not be used to infer protected characteristics.

## Required Human Review

Every generated row begins with `Pending Review`. The reviewer must check:

1. the source date, identity, location, completeness, and reconciliation;
2. the sample and peer-reference requirements;
3. whether assignment or operating context plausibly explains the result;
4. the recurring metric drivers and leave-one-active-day stability result; and
5. whether independent observations support the proposed follow-up.

The reviewer records one of `Coaching Accepted`, `Recognition Accepted`,
`Context Explains`, `Data Issue`, or `Monitor`, together with Reviewed By and
Review Date. Context Notes should contain only the minimum information needed
to explain the disposition. A prompt remains pending if those fields are
incomplete.

## Minimum Necessary Data

`ManagementEvidencePackageV2` includes the action, person/location, status,
owner, due date, recommended next step, reason/action codes, exact evidence
weeks, source hash/parser provenance, metric evidence, comparator type, peer
cohort size and weeks, threshold version, Evidence Status, recurring drivers,
leave-one-day stability result, review disposition, reviewer, review date, and
methodology version.

The approved package excludes free-form Context Notes and raw source
workbooks. Raw workbooks are not automatically attached, uploaded, or sent.
V1 packages remain readable but are never silently upgraded or treated as if
they contain V2 review evidence.

## Retention

- Identifiable action/coaching workbook and approved-AI evidence: 365 days.
- Each evidence package records its `delete_after` date.
- Deletion is manual and accountable; the automation does not silently delete,
  upload, email, or transmit evidence.
- Recovery copies follow 13-weekly and 12-monthly retention, with access
  limited to the business owner and technical recovery custodian.

## Sharing and Storage

- Finished reports: view-only access for authorized consumers.
- Intake/archive/automation: edit access only for the stable owner, technical
  maintainer, and the minimum required weekly runner role.
- Independent Google Drive recovery folder: private/restricted, independently
  administered, with two-factor authentication and access review.
- Do not forward, place in personal collaboration tools, or make offline copies
  outside an approved business purpose.

The one-time Gmail history backfill is restricted to original Marketing Vitals
TM report attachments needed for the approved historical window. Retrieve them
read-only, stage them outside the repository and live Dropbox folders, and
retain no email body or message identifier. Exclude forwarded duplicates,
derived `Check_Wine` workbooks, `No Data Available` workbooks, legacy
incompatible schemas, conflicting same-date files, and other report families.
Only complete Tuesday-Sunday weeks from the original TM reports may enter the
calibration cohort. Remove temporary staging copies after the canonical history
migration and rebuild verify. No recurring Gmail connector, credentials, or
mailbox metadata belong in this repository or normal weekly workflow.

## Corrections, Disputes, and Access Changes

Employees or managers may dispute an identity, source, date, metric,
peer-reference, missing context, threshold, or model/signal result through the
business owner. Preserve the original evidence and disposition for audit. If
authoritative source data are wrong, correct the source and regenerate through
the protected transaction; if context or methodology explains the dispute,
record that outcome without rewriting the original signal.

The business owner resolves use and context disputes. The technical maintainer
resolves source, integrity, and implementation defects. Material methodology
changes require a new version, updated model card, regression/backtest review,
and owner approval before deployment.

Remove access promptly when roles change and review membership quarterly.

See [MODEL_CARD.md](MODEL_CARD.md) for the formulas, cohort rules, limitations,
calibration contract, validation history, and change cadence.
