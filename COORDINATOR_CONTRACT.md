# Future shared coordinator contract — DESIGN ONLY

This is a proposed interface, not an implemented or enabled integration.
There is no new mail scanner, readiness endpoint, coordinator, scheduler,
notification sender, or management-email draft flow in this change.

Current `--health-check-json` checks local configuration/storage/receipt
health. It does **not** answer whether an incoming reporting set is complete.
Current run attempts and receipts are internal records. A future adapter must
implement and test the contract below before a coordinator relies on it.

## Ownership and boundaries

Red Onion owns business-date parsing, restaurant coverage, explicit closure
rules, reconciliation, duplicate/correction decisions, report calculations,
management carry-forward, local publication, and its project result.

A shared coordinator would own attachment routing, scheduling, safe retry
policy, notification deduplication, and management draft preparation. It
should deliver inputs to the configured intake folder and call the project's
interface rather than reimplement these reporting rules.

AOL remains authoritative for read/unread state; Gmail mirrors AOL, never the
reverse. This project must not alter mailbox state or store mailbox credentials.
Gift-card Micros intake remains with its existing separate automation.
Management messages are drafts for review. Publishing a report, returning a
successful result, or preparing a draft does not authorize sending it.

## Proposed invocation

A future implementation should provide two operations:

1. **Inspect reporting set:** read-only, bounded to an explicit workspace and
   Tuesday-Sunday period. Return readiness, coverage, and a business-fact
   fingerprint without generating workbooks, moving files, acquiring an
   approval, or changing mail state.
2. **Run reporting set:** accept the explicit period and the fingerprint from
   inspection. Revalidate the captured facts, current accepted state, config,
   and methodology under the local workflow lock, then publish or return a
   clear nonpublication outcome. A stale fingerprint requires fresh inspection.

These names describe operations; no corresponding CLI flags are available yet.
Only one machine may write at a time. No remote coordinator may treat the local
OS lock as a cross-computer lock.

## Proposed readiness schema

Schema name: `RedOnionReportingReadinessV1`. All listed fields are required
unless marked nullable or optional. Paths are relative to the operations root,
never credentials or machine-specific defaults. Timestamps are UTC ISO 8601;
business dates use `America/New_York`.

| Field | Type / contract |
| --- | --- |
| `schema_version` | Integer `1`. |
| `contract` | String `RedOnionReportingReadinessV1`. |
| `project` | String `red-onion`. |
| `checked_at_utc` | Timestamp. |
| `period` | Object with `start`, `end`, and `timezone`. |
| `status` | `ready`, `waiting_for_inputs`, `needs_review`, `already_processed`, or `unavailable`. |
| `coverage` | Counts for expected, observed, and explicitly closed restaurant-days; arrays `missing` and `conflicts`. |
| `input_set_id` | Nullable opaque business-fact fingerprint; null if facts cannot be validated. |
| `base_publication_id` | Nullable ID of the publication inspected. |
| `configuration_id` | Fingerprint of the effective configuration. |
| `methodology` | Management and descriptive version strings. |
| `issues` | Array of stable code, plain message, relative file reference when relevant, and required next action. |

Coverage is by restaurant and business date, not by file count. A missing
location/day is never an implicit closure. `unavailable` means that required
local/cloud files could not be read; it must not be reported as an empty intake.

`input_set_id` should cover canonical parsed facts, accepted revision base,
period, configuration, and methodology. A renamed report or equivalent Excel
save must not create a new business input set. Source-byte hashes remain
internal copy/provenance evidence. If accepted history or reviewed correction
alternatives change between inspection and execution, execution must not use
the stale inspection.

Illustrative incomplete result using the methodology versions recorded for
October 12, 2026. A current run records its actual versions; release v0.7.0
uses descriptive methodology `2026.10-v3` and management methodology `2026.07-v3`.

```json
{
  "schema_version": 1,
  "contract": "RedOnionReportingReadinessV1",
  "project": "red-onion",
  "checked_at_utc": "2026-10-12T12:00:00Z",
  "period": {
    "start": "2026-10-06",
    "end": "2026-10-11",
    "timezone": "America/New_York"
  },
  "status": "waiting_for_inputs",
  "coverage": {
    "expected_restaurant_days": 12,
    "observed_restaurant_days": 11,
    "closed_restaurant_days": 0,
    "missing": [
      {"location": "RC Virginia Beach", "business_date": "2026-10-11"}
    ],
    "conflicts": []
  },
  "input_set_id": null,
  "base_publication_id": "example-previous-publication",
  "configuration_id": "example-config-fingerprint",
  "methodology": {
    "management": "2026.07-v3",
    "descriptive": "2026.10-v2"
  },
  "issues": [
    {
      "code": "missing_restaurant_day",
      "message": "Virginia Beach Sunday data is missing.",
      "file": null,
      "next_action": "Supply the original complete daily report."
    }
  ]
}
```

With all 12 days accounted for, no conflict, validated reconciliation, and a
new input set, `status` becomes `ready` and `input_set_id` is non-null.
An unresolved corrected report returns `needs_review`; the coordinator may
surface the review request but may not choose an alternative.

## Proposed run-result schema

Schema name: `RedOnionReportingResultV1`.

| Field | Type / contract |
| --- | --- |
| `schema_version` | Integer `1`. |
| `contract` | String `RedOnionReportingResultV1`. |
| `project` | String `red-onion`. |
| `attempt_id` | Unique ID for this invocation. |
| `completed_at_utc` | Timestamp. |
| `period` | Same business-period object as readiness. |
| `input_set_id` | Nullable inspected/accepted business input identity. |
| `outcome` | `waiting_for_inputs`, `needs_review`, `already_processed`, `published`, `published_with_warnings`, or `failed_before_publication`. |
| `publication_id` | Nullable durable publication identity; repeat submissions retain the existing identity. |
| `published_this_attempt` | Boolean. |
| `changed_weeks` | Array of affected Sunday dates. |
| `outputs` | Array of role and operations-root-relative path. |
| `issues` | Structured issue array as above, including cleanup warnings. |
| `local_publication` | `committed`, `unchanged`, or `not_committed`. |
| `cloud_sync` | `not_checked` unless a separately implemented cloud check supplies evidence. |
| `management_evidence` | Object with local approval status and nullable approved artifact reference; no raw workbook or free-form notes. |

Illustrative post-commit cleanup warning:

```json
{
  "schema_version": 1,
  "contract": "RedOnionReportingResultV1",
  "project": "red-onion",
  "attempt_id": "example-attempt",
  "completed_at_utc": "2026-10-12T12:05:00Z",
  "period": {
    "start": "2026-10-06",
    "end": "2026-10-11",
    "timezone": "America/New_York"
  },
  "input_set_id": "example-business-fingerprint",
  "outcome": "published_with_warnings",
  "publication_id": "example-durable-publication",
  "published_this_attempt": true,
  "changed_weeks": ["2026-10-11"],
  "outputs": [
    {
      "role": "master",
      "path": "02 Finished Reports/Red_Onion_Server_Master.xlsx"
    }
  ],
  "issues": [
    {
      "code": "input_cleanup_pending",
      "message": "Reports were published; one saved input remains in intake.",
      "file": "01 Daily Reports - Drop Here/example-daily.xlsx",
      "next_action": "Close the input workbook; safe cleanup resumes next run."
    }
  ],
  "local_publication": "committed",
  "cloud_sync": "not_checked",
  "management_evidence": {
    "approval_status": "not_requested",
    "approved_artifact": null
  }
}
```

Examples use placeholder identities and an abbreviated output list; production
results must list the full published set. Current local `history_imported`
is a maintainer operation and is outside the proposed ordinary coordinator
run interface.

## Notification, retry, and draft rules

The coordinator must use the durable publication ID, not just a process exit
code or attempt ID, to deduplicate success notifications and drafts.
`already_processed` may acknowledge the original publication but cannot
create another success event. A corrected week produces a new publication
identity and lists changed weeks so the coordinator can clearly describe it.

A post-commit warning never triggers an automatic fresh publication. A
`needs_review` response waits for a real correction decision. A transient
`unavailable` readiness result may be inspected again after files become
available; elapsed time is not proof of completeness. A process error with an
uncertain commit outcome requires receipt reconciliation before retry.

The minimal result envelope contains no employee-level metrics, management
notes, raw source workbooks, credentials, or mail identifiers. Preparing a
management draft that uses named evidence must respect the existing exact
local candidate/fingerprint approval and its allowed purpose. Draft creation
and review belong to the shared coordinator's separately authorized workflow;
nothing here bypasses that review or sends a message.

## Required validation before implementation is enabled

Test incomplete and closed days, conflicting corrections, equivalent renamed
inputs, stale inspection fingerprints, a retry after commit, a cleanup warning,
unavailable storage, preserved management edits, and a moved workspace. Prove
that repeated accepted facts produce one publication notification and that
mail read/unread state is untouched. Record a real two-PC sequential Dropbox
exercise separately from local copied-folder tests.

Implement and review this adapter separately. Adding this document does not
authorize a connector, schedule, external write, or message send.
