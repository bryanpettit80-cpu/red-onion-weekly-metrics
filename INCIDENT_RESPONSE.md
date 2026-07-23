# Incident Response

Stop weekly processing when integrity, release, input, workbook, environment,
or evidence approval checks fail. Preserve the intake files, console error,
`LAST RUN STATUS.txt`, the matching `03 Archive\run-attempts` JSON, current
manifest/anchor, and affected artifact hashes.

## Release or Environment Failure

- Do not edit the deployed checkout to silence preflight.
- Restore a clean `main` at the approved release commit.
- Use `-RebuildEnvironment` only after configuration and lock review.
- Resume after release parity, dependency verification, health check, and full
  tests pass.

## Manifest, Raw Data, or Workbook Integrity Failure

- Do not initialize a replacement baseline.
- Prevent further Dropbox edits where practical.
- Compare the trusted anchor, manifest chain, canonical raw inventory,
  generated-workbook snapshot, published workbook, and independent recovery
  copy.
- Restore only an exact verified version. Record hashes and who approved it.
- Resume only after health and integrity verification succeed.

## Incorrect or Missing Daily Report

- Leave active files in intake.
- Replace incomplete exports from the authoritative Toast source.
- Re-run only after the business date, locations, and duplicate/conflict
  messages are reconciled.

## Employee-Performance or Evidence Incident

- Restrict access immediately and preserve the approved candidate,
  fingerprint, approval receipt, and distribution facts.
- Do not create a replacement approval for altered bytes.
- Notify the business owner; document recipients, scope, dates, and requested
  deletion/correction.
- Resume evidence use only after the exact approved purpose and 365-day
  retention obligations are restored.

For every incident, record owner, timeline, evidence preserved, root cause,
corrective action, validation, approval to resume, and any required employee
correction.
