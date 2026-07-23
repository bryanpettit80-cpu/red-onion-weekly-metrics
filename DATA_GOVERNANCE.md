# Employee Data Governance

Server scorecards, named actions, coaching/recognition recommendations, owner
assignments, evidence details, and approved evidence exports are Restricted
Employee Performance Information.

## Allowed Use

- Weekly management review by authorized Red Onion leaders.
- Identifiable coaching, recognition, and action tracking.
- AI-assisted analysis only after a named authorized manager reviews the exact
  `ManagementEvidencePackageV1` candidate and approves its candidate and
  fingerprint hashes for a specific purpose.

The approval does not authorize a different file, a regenerated package, a
different purpose, or broader distribution.

## Minimum Necessary Data

The approved evidence package includes the action, person/location, status,
owner, due date, recommended next step, reason/action codes, evidence weeks,
source hash/parser provenance, and metric evidence. It excludes free-form
manager notes. Raw workbooks are not automatically sent with the package.

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

## Corrections and Access Changes

Employees or managers may report an identity, source, date, or metric error to
the business owner. Preserve the original evidence, correct the authoritative
source, regenerate through the normal transaction, and document the correction.
Remove access promptly when roles change and review membership quarterly.
