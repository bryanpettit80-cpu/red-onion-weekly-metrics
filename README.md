# Red Onion Weekly Metrics

Red Onion turns daily Marketing Vitals reports into weekly restaurant reports
and a master workbook. Operators use the numbered folders and one launcher.
See [SETUP.md](SETUP.md) to prepare another Windows computer.

The current release is [v0.6.0](https://github.com/bryanpettit80-cpu/red-onion-weekly-metrics/releases/tag/v0.6.0).
The reporting workflow runs locally on Windows with Python 3.10-3.12;
Python 3.12 is recommended. Excel is optional for generation and useful for
reviewing and editing the finished workbooks. Linux CI exercises the Python
calculation and recovery tests; a separate Windows job tests the launchers.
Cloud development can use the source and synthetic tests. Live Dropbox data,
Windows launchers, and cross-computer sync require the local workflow.

- Operators: [quick guide](00%20START%20HERE%20-%20Red%20Onion%20Weekly%20Metrics.txt),
  [Windows setup](SETUP.md), and the weekly steps below.
- Developers: [development and Cloud setup](MAINTAINER.md#development-and-tests),
  [program map](_program/README.md), and [repository instructions](AGENTS.md).
- Maintainers: [releases and corrections](MAINTAINER.md),
  [recovery](RECOVERY.md), and [incident response](INCIDENT_RESPONSE.md).

GitHub's reviewed `main` branch is the shared source history. A maintainer
checkout is for source changes; the program folder in Dropbox is the deployed
copy used by the launcher. The numbered Dropbox folders hold the operational
inputs, finished workbooks, and recovery history. Keep the latest master with
its management edits and the archive together. A source checkout or release
archive is not a backup of that operational data.

## Folder layout

```text
Red Onion Metrics\
  01 Daily Reports - Drop Here\
  02 Finished Reports\
  03 Archive\
  Red Onion Weekly Metrics Automation\
  Run Weekly Snapshot.cmd
```

The program folder contains source, configuration, tests, and technical
documentation. It may be renamed. Git is useful for maintaining the program;
an ordinary run needs neither Git nor a particular branch or machine-local
integrity anchor.

## Weekly run

1. Let Dropbox finish downloading the workspace. Use one computer at a time.
2. Save the week's original Marketing Vitals `.xls` or `.xlsx` reports in
   `01 Daily Reports - Drop Here`. Close the workbooks in Excel.
3. Double-click `Run Weekly Snapshot.cmd`.
4. Read the outcome, then open the reports in `02 Finished Reports`.

Each workbook must contain an explicit single business date in its Date(s)
field. Filenames do not supply or override that date.
Excel temporary files beginning with `~$` are ignored.

A complete week is Tuesday through Sunday for both configured restaurants:
**12 restaurant-days**, normally supplied by six daily workbooks containing
both locations. A file count alone does not establish completeness. The run
checks dates, restaurant coverage, numeric values, duplicate business records,
and reconciliation with restaurant totals. Missing days remain missing; they
are not inferred to be closed or treated as zero. An actual closure must be
recorded explicitly in the location's `closed_dates` configuration.

Intake must contain only one operating week. Previously accepted reports may
supply other dates in that same week. An incomplete set leaves previous
finished reports unchanged and reports the missing restaurant-days.

## Outcomes and corrections

| Outcome | Meaning and next step |
| --- | --- |
| `waiting_for_inputs` | No new inputs or no complete reporting set. Supply the listed reports; previous reports remain available. |
| `needs_review` | A business conflict, invalid management input, or recovery conflict needs attention. Read the specific message. |
| `already_processed` | The business facts were already accepted and the configuration/methodology are unchanged. Reports were not regenerated. |
| `published` | The complete report set was published locally. |
| `published_with_warnings` | Publication succeeded; review the stated cleanup or status warning. |
| `failed_before_publication` | The new reporting set did not commit. Resolve the stated problem before retrying. |
| `history_imported` | Existing history was saved without generating reports. The next ordinary run refreshes complete accepted history. |

`LAST RUN STATUS.txt` is the convenient summary. The portable run receipt in
`03 Archive\Reporting State` records the committed result. Neither a success
message nor a health check verifies Dropbox cloud sync or recipient access.

Semantically identical reports count once even if their filenames, formats, or
file bytes differ. A changed business value requires one explicit choice for
each conflicting date. To see the values and choose:

```powershell
& '.\Run Weekly Snapshot.cmd' -ReviewCorrections
```

The review offers the accepted report, where one exists, and the incoming
alternatives. It records the choice against those exact business facts.
Changing the alternatives requires a new review. An ordinary complete,
unambiguous week needs no approval prompt. See [MAINTAINER.md](MAINTAINER.md)
for reviewed correction files and historical imports.

## Reports and saved work

The runner generates the applicable `Check_Wine_RVA<week_end>.xlsx` and
`Check_Wine_VB<week_end>.xlsx` reports and
`Red_Onion_Server_Master.xlsx`. Correcting an older accepted week updates
that week's restaurant reports and rebuilds the master through the latest
complete week.

The master has seven visible tabs:

- `How to Use`: workflow, navigation, editable fields, and interpretation.
- `Performance Dashboard`: store totals and all named-server summaries.
- `Server Scorecards`: actual results, peer references, differences,
  consistency, coverage, and low-data explanations.
- `Weekly Performance`: the dated pattern behind those summaries.
- `Shared & Area Trends`: shared POS identities and operational areas.
- `Methodology`: definitions and limits.
- `Management Center`: data readiness, targets, owners, current actions,
  and action history.

Detailed evidence, data quality, run notes, legacy views, and calculation
sheets remain protected and `veryHidden` support layers.

Management targets, the owner roster, and Current Actions Status, Owner,
Due Date, Context Notes, Review Disposition, Reviewed By, and Review Date
carry forward. The workbook is recognized by its management tables; a
harmless filename change, Excel save, or formatting change does not block
ordinary reporting. Keep the table names, required columns, and Action IDs.
If several master copies contain different management inputs, the run asks
which one to use instead of discarding edits.

After the first import, saved action history and generated evidence come
from the portable reporting state. Only the documented editable fields are
merged from the workbook. Generated metrics are recomputed from accepted
daily facts. Changing a displayed generated number does not change those
facts and the next generation replaces it. Use the source correction workflow
to correct reporting data.

Protection helps avoid accidental edits; the lowercase password `redonion`
is not encryption. The supported edit cells remain blue in Management Center.
Keep exploratory work in a separate copy.

## Calculation and interpretation

Sales/Guest is total Gross Sales divided by total Guests. Wine Percentage is
total Wine Sales divided by total Gross Sales. Totals and combined rates use
their underlying amounts, rather than averages of displayed percentages.

Rate of Sale is Red Onion's inverse conversion measure,
`opportunities / qualifying sales`; lower positive values are better. The
current field uses Guests as opportunities. Missing or nonpositive rates
make the combined context unavailable.

Ticket Time is check-weighted only when every contributing row has valid
Check Count and the total is positive. Guest weighting is not a substitute.
See the [metric formulas](MODEL_CARD.md#descriptive-context-metrics) for
aggregation details and unavailable-data rules.

The `2026.10-v2` descriptive scorecards cover every resolved named server
seen in the latest eight complete weeks, including someone absent from the
latest week. Low-data people remain visible with explicit coverage limits.
Shared POS identities and operational areas remain separate. The
[scorecard methodology](MODEL_CARD.md#descriptive-performance-and-consistency-layer)
defines qualified weeks, historical peers, weights, and classification rules.

The separate `2026.07-v3` people-review prompts use only Sales/Guest and Wine
Percentage with sample, history, peer, persistence, and stability checks.
Operational context and descriptive scorecard labels cannot create or escalate
those prompts. Missing evidence prevents escalation; missing weeks are not zero.

These are rule-based observational review prompts. They do not establish
causation, statistical confidence, total job performance, or employment
decisions. A manager must review source accuracy and comparable work context.
See [MODEL_CARD.md](MODEL_CARD.md) and
[DATA_GOVERNANCE.md](DATA_GOVERNANCE.md) for the complete analytical contract.

## Dependability and recovery

The runner holds a local workflow lock, captures input bytes, validates and
stages workbooks, saves previous-good copies, and journals publication.
Inputs leave intake only after a durable commit and verification of their
saved copies. A changed input or management workbook is preserved for review.

An interruption before commit rolls back the incomplete publication on
recovery. After commit, a cleanup failure remains a successful publication
with a warning. The next ordinary invocation reconciles pending journals
before accepting more work. A conflicting external edit stops recovery and
preserves the conflicting files.

Copy the whole operations root for backup or a move, including the latest
master and `03 Archive\Reporting State`. The local Python environment is
disposable. Historical generated-workbook hashes, an old manifest chain,
and computer-specific anchors are not ordinary-run prerequisites. See
[RECOVERY.md](RECOVERY.md) and [INCIDENT_RESPONSE.md](INCIDENT_RESPONSE.md).

Sequential use across Windows computers requires fully downloaded files,
one active writer, and completed Dropbox sync before switching computers.
The local lock does not coordinate simultaneous writers on different PCs.
Automated copied-folder tests do not constitute a physical two-computer
Dropbox restore exercise.

## Configuration and maintenance

`_program\red_onion_config.json` defines locations, explicit closures,
thresholds, exclusions, aliases, and area mappings. A closure example is:

```json
{
  "closed_dates": {
    "RC Richmond": ["2026-10-06"]
  }
}
```

This is an illustration, not an assertion that Richmond was closed that day.
Record only verified closures. A closure that conflicts with observed business
data requires correction.

Use `-HealthCheck` for a read-only configuration/storage/receipt check and
`-RebuildEnvironment -HealthCheck` to repair the local dependencies without
processing reports. [SETUP.md](SETUP.md) covers setup and path selection;
[MAINTAINER.md](MAINTAINER.md) covers releases, imports, evidence export,
and the separate legacy integrity commands.

The repository does not scan mail, send notifications, create management
email drafts, or activate schedules. The future shared-coordinator interface
is documented, **design only**, in [COORDINATOR_CONTRACT.md](COORDINATOR_CONTRACT.md).
