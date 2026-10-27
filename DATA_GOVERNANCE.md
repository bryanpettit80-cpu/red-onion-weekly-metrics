# Employee Data Governance

Server scorecards, named signals, coaching/recognition prompts, review
dispositions, owner assignments, evidence details, and approved evidence
exports are Restricted Employee Performance Information.

The tracked-action methodology `2026.07-v3` uses deterministic observational
screening rules. They are not a statistical, predictive, causal, or
employment-decision model. Data integrity and reproducibility do not establish
that a signal fairly attributes performance to one person.

The separate `2026.10-v3` staff discussion views include all resolved named
servers in the latest eight complete weeks. Usable positive sales and guests
can support feedback after one active day when at least five other qualified
same-store/week peers are available. Established, Building, and Early
indication describe the sample, not statistical confidence. Small samples
retain conversation topics with context; missing denominators or peer
references remain unavailable. These suggestions do not create or escalate
tracked actions. Exact, verified per-location identity aliases may join name variants for descriptive reports;
they do not alter management Action IDs or transfer assignments. Shared POS
identities must never be converted into person identities.

## Allowed Use

- Weekly context review by authorized Red Onion leaders.
- Identifiable coaching and recognition conversations after checking source
  accuracy, the comparison sample, and comparable-work context. Tracked-action
  dispositions also require their documented persistence and stability review.
- Accountable tracking of the generated signal, manager disposition, reviewer,
  review date, and resulting follow-up.
- AI-assisted analysis only after a named authorized manager reviews the exact
  `ManagementEvidencePackageV2` candidate and approves its candidate and
  fingerprint hashes for a specific purpose.

The approval does not authorize a different file, a regenerated package, a
different purpose, or broader distribution. Existing V1 packages remain
readable for audit and retention, but all newly generated packages use V2.

## Aggregate Operational Views

Leading four-digit shared POS identities are shared operating records, not
person identities. They must remain outside server scorecards, peer cohorts,
people-review signals, evidence packages, and employment decisions.

Bar, Patio, Dining Room, Banquets, and Wine Dinners trends are operational
aggregates across complete weeks. Spend per guest is guest-weighted: total Gross
Sales divided by total Guests. Dining Room is the fallback for eligible named
rows not assigned to another area. Wine Dinners must show as unavailable until
the source supplies an explicitly configured name or
`weekly_shared_number_areas` maps a shared POS number to that area. These views
may support high-level operating observation, but they cannot create, change,
or escalate a person-level review action.

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

Every generated tracked-action row in Management Center begins with
`Pending Review`. Before completing its disposition, the reviewer must check:

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

## Workbook Presentation And Edit Boundaries

The operator-facing workbook has seven visible tabs: `How to Use`,
`Performance Dashboard`, `Server Scorecards`, `Weekly Performance`, `Shared &
Area Trends`, `Methodology`, and `Management Center`. `Management Center`
consolidates the data-readiness summary, editable targets and owner roster,
Current Actions, and locked Action History that previously occupied separate
operator tabs.

Staff discussion suggestions remain separate from this action queue. History
distinguishes completed, dismissed, and cleared signals; clearing a signal
preserves its recorded review status and does not certify completed follow-up.

The permitted workbook inputs are target values in `D:I` (Entity in `C` stays
locked), Owner Roster values in `K:L`, and Current Actions Status (`D`), Owner
(`E`), Due Date (`F`), Context Notes (`N`), Review Disposition (`U`), Reviewed
By (`V`), and Review Date (`W`). The readiness summary and Action History are
read-only. Detailed `Data Quality`, `Evidence Detail`, and `Run Notes`, legacy
presentation sheets, and technical calculation/raw layers remain protected
`veryHidden` audit/support sheets. Their visibility does not change their
classification, retention, review, or access requirements.

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

The lowercase workbook password `redonion` is only a convenience against
accidental edits for authorized operators. It is not encryption, does not grant
Dropbox access, and does not replace role restrictions or recovery copies.
Ordinary reporting accepts a harmless Excel save, formatting change, or master
filename change. Required management tables, column names, valid inputs, and
stable Action IDs must remain readable. Generated metrics are recomputed from
accepted daily facts; editing a displayed number does not change those facts.
Correct the authoritative source through the correction workflow and regenerate.

Portable Reporting State contains accepted daily facts, source copies, saved
management inputs, generated action evidence, and run receipts. Treat it and
the latest editable master as restricted operational data. After the first
import, only the documented editable fields merge from the workbook; saved
generated actions, history, and evidence retain their accepted identities.
Write access to Reporting State grants authority over accepted history and
recovery actions in the intake and finished-report folders. Give it only to
trusted reporting operators; report consumers and upload-only users must not
have it. Receipts and checksums detect damage or incomplete copies, not
deliberate edits by someone with that write access.
Back up the whole operations root, not only the finished reports. The local
Python environment and old machine anchors are not needed for portable recovery.

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
authoritative source data are wrong, preserve the original and review the
changed business facts through the correction workflow before regeneration;
if context or methodology explains the dispute,
record that outcome without rewriting the original signal.

The business owner resolves use and context disputes. The technical maintainer
resolves source, integrity, and implementation defects. Material methodology
changes require a new version, updated model card, regression/backtest review,
and owner approval before deployment.

Remove access promptly when roles change and review membership quarterly.

See [MODEL_CARD.md](MODEL_CARD.md) for the formulas, cohort rules, limitations,
calibration contract, validation history, and change cadence.

The future shared-coordinator interface is [design only](COORDINATOR_CONTRACT.md).
It must preserve AOL as the read/unread authority, keep Gmail as its mirror,
leave existing gift-card Micros intake separate, and create management emails
only as drafts for review. This repository currently performs no recurring mail
scan, notification send, draft creation, or schedule activation. The existing
two-stage local evidence approval remains separate from ordinary reporting.
