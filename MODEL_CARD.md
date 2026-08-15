# Red Onion Weekly Metrics Model Card

## Identity And Status

- **System:** Red Onion Weekly Metrics
- **Release contract:** `v0.5.1` (descriptive performance/consistency views;
  people-review methodology remains `2026.07-v3`)
- **Operational status:** protected backfill migration, exact post-migration
  replay, and managed local publication are complete; authorized people-review
  use remains blocked until a separate business-owner approval record is
  documented and verified
- **People-review methodology:** `2026.07-v3`
- **Descriptive performance/consistency methodology:** `2026.08-v1`
- **System type:** deterministic, rule-based observational coaching signal
- **Owners:** Red Onion business owner and designated technical maintainer
- **Decision authority:** accountable human manager; never the automation

This model card describes three intentionally separate layers. The operations
layer presents trends, volume, and general context. The descriptive
performance/consistency layer summarizes observed Sales/Guest and Wine outcomes
and data sufficiency. The people-review layer covers person-level Recent
Movement, Peer Comparison, Context Review, Coaching Prompt, and Recognition
Prompt behavior.

## Intended Use

The system identifies changes that an authorized manager may investigate
during weekly coaching or recognition review. It organizes evidence and makes
the calculation repeatable; it does not decide whether an employee has
performed well or poorly.

Appropriate uses are:

- identifying an operational change that warrants a source or context check;
- comparing check, guest, and sales mix without treating the comparison as an
  employment conclusion;
- asking context-aware coaching or recognition questions;
- checking source accuracy and comparable-work conditions;
- recording a manager's independent disposition and follow-up; and
- monitoring aggregate signal quality during the validation pilot; and
- reviewing a descriptive eight-week selling-outcome pattern while separately
  checking data sufficiency, consistency, and operating context.

## Prohibited Use

No signal may be the sole or determinative basis for discipline, termination,
pay, promotion, demotion, scheduling, reduced opportunity, formal performance
ratings, or another adverse employment decision. The system must not be
described as:

- statistically confident or statistically significant;
- predictive of future performance;
- causal evidence of employee behavior;
- adjusted for assignments or equal work opportunity; or
- validated for demographic fairness.

Independent corroboration and an accountable human review are mandatory.
`Overall Read`, Performance, Consistency, Confidence, weekly bands, ranks, and
dashboard categories are not formal performance ratings or standalone
employment evidence. They measure neither total job performance nor equal
sales opportunity.

## Inputs, Population, And Missing Context

The system reads Tuesday-Sunday Red Onion Marketing Vitals daily reports.
Available person-level inputs are location, displayed identity, sales, guests,
wine, Rate of Sale, Average Ticket Time, and Check Count when supplied.
Configured aliases and exclusions define the eligible server population.

The source does not provide a stable employee identifier, shift, daypart,
section, party mix, scheduled hours, staffing, events, tenure, training, menu
availability, or protected attributes. Check Count measures volume but does not
provide those missing comparable-work conditions. Those omissions limit
performance attribution and fairness assessment. Protected characteristics
must not be inferred from names.

A positive guest count paired with a blank, malformed, NaN, or infinite Rate of
Sale or Ticket Time is unavailable data, not a favorable zero. Negative values,
invalid time components, and malformed or non-whole Check Counts are also
unavailable. Context-metric unavailability is reported but cannot create or
improve a person-level signal.

## Eligibility And Comparisons

A person-week is eligible only when the latest week is complete and reconciled
and the person has:

- at least 25 guests;
- at least three active days;
- at least two prior complete self-weeks; and
- at least 50 guests across the prior comparison period.

`Recent Movement` compares the current complete week with up to four prior
complete weeks for the same person.

`Peer Comparison` uses a leave-one-person-out, same-store median from the prior
four complete weeks. The cohort contains only non-excluded peer-weeks meeting
the same guest and active-day gates. A peer reference requires:

- at least three usable prior weeks;
- at least five distinct eligible peers in each usable week; and
- at least 20 peer-week observations.

If these requirements are not met, the result is `Reference Unavailable` and a
coaching or recognition prompt is suppressed. Management targets may be shown
as business context, but they do not drive person-level prompts. Any retained
eligible-cohort percentile or rank is descriptive only and cannot change a
classification.

## Descriptive Performance And Consistency Layer

Methodology `2026.08-v1` is a separate deterministic review aid. It uses the
latest eight globally complete Tuesday-Sunday weeks and the current
non-excluded roster as of the latest complete shared week. It never supplies an
input, score, driver, persistence result, or action to the `2026.07-v3`
people-review path.

A current-roster person-week qualifies only with at least 25 guests, at least
three active days, and at least five other qualified current-roster servers in
the same store and week. The peer reference is the leave-one-person-out median.
Sales/Guest and Wine gaps are the person's result minus that weekly peer
median. Summary peer gaps use `min(guests, 50)` as the weekly weight; eight-week
Sales/Guest and Wine % use aggregated qualified sales, guests, and wine sales.

Weekly bands use these rules:

- `Strong`: Sales/Guest gap is at least +$11.50 with nonnegative Wine gap, or
  Wine gap is at least +4.1 percentage points with nonnegative Sales/Guest gap.
- `Below`: the symmetric negative boundary with the other gap nonpositive.
- `Near Peer`: neither boundary is met and the gaps share a sign.
- `Mixed`: neither boundary is met and the gaps have opposite signs.
- `Not Qualified`: the sample or peer gate fails. A missing row remains visibly
  unavailable; neither case is converted to zero.

Performance applies the same boundary logic to capped guest-weighted average
peer gaps. Data-sufficiency `Confidence` is `High` at six qualified weeks and
200 qualified guests, `Provisional` at four weeks and 150 guests, and
`Insufficient` otherwise. This is not statistical confidence, significance,
predictive validity, or causal certainty.

Consistency uses sample standard deviation of the qualified weekly peer gaps.
It is `High` only when Sales/Guest SD is no greater than $11.50 and Wine-gap SD
is no greater than 4.1 percentage points; `Moderate` requires no greater than
$17.50 and 5.7 points; otherwise it is `Low`. Insufficient-confidence rows are
always `Insufficient` consistency. `Overall Read` combines the performance and
consistency axes only after the data-sufficiency gate; an insufficient row is
always `Insufficient Data`.

Recent movement is the capped-weighted recent-four peer gap minus the
prior-four peer gap. The workbook highlights magnitudes of at least $5 or 1.5
percentage points, but this remains descriptive and cannot create an action.
The runner regenerates static protected values from verified inputs; these are
not live Excel formulas.

The `Performance Dashboard` also presents an all-stores operating card. Its
sales, guests, Sales/Guest, and wine measures come only from reconciled location
totals; changes compare the latest complete week with the combined prior four
complete weeks. No person-level score or shared POS identity contributes to
that card.

The separate `Shared & Area Trends` view treats any leading four-digit POS
identity as a shared operating identity, never as a person or people-review
input. Bar, Patio, Dining Room, Banquets, and Wine Dinners are aggregated by
complete week, with Sales/Guest calculated as total Gross Sales divided by
total Guests. Dining Room is the fallback for eligible named rows not mapped
to another area. Wine Dinners remains unavailable until the source supplies a
configured name or `weekly_shared_number_areas` explicitly maps a shared POS
number to it. These aggregates are operational context only and cannot create,
change, or escalate a person-level action.

## Metrics, Scoring, And Calibration

Methodology `2026.07-v3` limits the people-review composite to two
action-driving metric families:

- **Sales/Guest:** the existing internal `check_average` field, calculated as
  total sales divided by total guests.
- **Wine Percentage:** the existing wine percentage measure.

Each qualified action-metric deviation receives a deterministic score from
`-2` to `+2`. A candidate direction requires an absolute composite of at least
three and at least two agreeing metric families. Because only two metric
families are eligible, both must agree. Rate of Sale, Average Ticket Time,
Check Count, Sales/Check, and Guests/Check cannot supply an agreeing metric,
change a composite, become a recurring driver, or affect persistence or
escalation.

The action-metric thresholds are frozen in configuration for each methodology
release. For `2026.07-v3`, the Sales/Guest and Wine Percentage bands use the
verified 16-week calibration:

- **Neutral:** the larger of the documented business minimum and the R-7 75th
  percentile of absolute qualified deviations.
- **Strong:** the larger of the documented business minimum and the R-7 90th
  percentile of absolute qualified deviations.
- **Rounding:** half-up to $0.50 for Sales/Guest and 0.001 for Wine
  Percentage. If rounding makes Neutral and Strong equal, Strong increases by
  one rounding increment.

The frozen `2026.07-v3` people-review bands are:

| Metric | Movement neutral / strong | Peer neutral / strong |
|---|---:|---:|
| Sales/Guest | $11.50 / $17.50 | $11.00 / $16.50 |
| Wine Percentage | 4.1 / 5.7 percentage points | 3.9 / 5.4 percentage points |

Movement calibration used 227 self-history-qualified person-weeks, or 454
action-driving metric deviations; peer calibration used 202
peer-reference-qualified person-weeks, or 404 action-driving metric deviations.
The 16 complete weeks span March 24 through July 19, 2026, and the two
comparator families were calibrated independently.

Calibration runs only as a maintainer-controlled release activity. An ordinary
weekly run never adapts thresholds. Calibration metadata records the source
window, method, date, threshold version, and resulting frozen values. Review
the bands when the verified history first reaches 12 weeks and again at 26
weeks; any change requires a new methodology version and regression review.

### Descriptive Context Metrics

Red Onion defines Rate of Sale as inverse conversion:

```text
Rate of Sale = opportunities / qualifying sales
```

Lower positive values are better. When positive available row-level rates are
combined, the workbook uses the opportunity-weighted harmonic calculation. The
current Rate of Sale by Guest Count field uses Guests as the opportunity count:

```text
combined ROS = sum(opportunities) / sum(opportunities / row ROS)
```

This is equivalent to total opportunities divided by the reconstructed total
qualifying sales. An arithmetic average is not valid. A nonpositive, malformed,
or missing rate cannot be safely reconstructed from the ratio alone and makes
the combined context value unavailable.

Average Ticket Time is check-weighted only when every contributing row has a
valid Check Count and total Check Count is positive:

```text
combined Ticket Time = sum(row Ticket Time * row Check Count)
                       / sum(row Check Count)
```

Incomplete Check Count coverage makes the combined Ticket Time unavailable;
guest weighting is not a fallback. Complete coverage with a positive total
Check Count also supports:

```text
Checks       = sum(Check Count)
Sales/Check  = sum(sales) / sum(Check Count)
Guests/Check = sum(guests) / sum(Check Count)
```

These measures help a manager evaluate workload and mix. They remain
descriptive operational context because the available reports do not establish
comparable assignments, causal responsibility, or equal sales opportunity.

## `2026.07-v3` People-Review Signal And Persistence Rules

Visible values are:

- Peer Comparison: `Above Peer Reference`, `Within Peer Range`,
  `Below Peer Reference`, or `Reference Unavailable`.
- Recent Movement: `Upward`, `Downward`, `Stable`, or `Not Evaluated`.
- Actions: `Context Review`, `Coaching Prompt`, `Recognition Prompt`, or
  `Monitor`.

A positive candidate requires Upward movement and Above Peer Reference. A
negative candidate requires Downward movement and Below Peer Reference.
Movement that is not materially different from the current qualified peer
median is treated as a possible common store shock and cannot by itself be
attributed to one person.

Only Sales/Guest and Wine Percentage participate in those directions. A change
in Rate of Sale, Ticket Time, Check Count, Sales/Check, or Guests/Check may
suggest a manager question, but it cannot create or reverse a direction.

The first qualified positive or negative candidate creates a `Context Review`.
A `Coaching Prompt` or `Recognition Prompt` requires a second consecutive
qualified week with:

- the same signal polarity;
- at least one recurring metric driver; and
- the same result after separately removing each active day from both weeks.

An incomplete or low-volume week, unavailable peer reference, missing calendar
week, changed direction, or failed stability check breaks escalation. A
day-sensitive signal cannot exceed `Context Review`.

## Human Review And Evidence

Every generated row in the `Management Center` Current Actions section begins
with Review Disposition `Pending Review`. An authorized manager checks the
source, identity, peer cohort, comparable-work context, recurring drivers, and
stability evidence, then records one of:

- `Coaching Accepted`
- `Recognition Accepted`
- `Context Explains`
- `Data Issue`
- `Monitor`

Reviewed By and Review Date are required for a completed disposition. Context
Notes are limited to information necessary to explain that decision.

The review should answer, at minimum:

- Are the source values and displayed identity correct?
- Was the work reasonably comparable by role, shift, section, party mix, and
  operating conditions?
- Did check volume, Guests/Check, or a common store condition move at the same
  time?
- Do both Sales/Guest and Wine Percentage support the person-level direction?
- What independent evidence supports the final coaching, recognition, context,
  data-issue, or monitor disposition?

New approved evidence exports use `ManagementEvidencePackageV2`. In addition
to V1 lineage fields, V2 records the comparator type, peer cohort size and
weeks, threshold version, Evidence Status, recurring drivers, leave-one-day
stability result, review disposition, reviewer, review date, and methodology
version. Existing V1 packages remain readable for audit, but they are not
treated as containing V2 review evidence.

Evidence export remains a manual, exact-fingerprint approval workflow. No
package is uploaded, emailed, or transmitted automatically.

## Workbook Presentation And Controls

The operator-facing workbook has seven visible tabs: `How to Use`,
`Performance Dashboard`, `Server Scorecards`, `Weekly Performance`, `Shared &
Area Trends`, `Methodology`, and `Management Center`. `Management Center`
consolidates the data-readiness summary, targets and owner roster, Current
Actions, and locked Action History that were previously presented on separate
tabs.

The supported inputs are limited to target values in `D:I` (Entity in `C`
remains locked), Owner Roster values in `K:L`, and Current Actions Status (`D`),
Owner (`E`), Due Date (`F`), Context Notes (`N`), Review Disposition (`U`),
Reviewed By (`V`), and Review Date (`W`). Detailed `Data Quality`, `Evidence
Detail`, and `Run Notes`, legacy presentation sheets, and technical
calculation/raw layers remain protected `veryHidden` audit/support sheets.
Their hidden status does not remove them from validation, lineage, retention,
or substantive-digest controls.

## Validation And Acceptance

Release validation includes unit, workbook-contract, integrity, migration, and
historical backtests. The candidate historical replay must demonstrate:

- 100% person-action invariance when only Rate of Sale, Ticket Time, Check
  Count, Sales/Check, or Guests/Check changes;
- 100% action invariance when excluded rows or descriptive ranks change;
- 100% leave-one-active-day stability for Coaching and Recognition Prompts;
- no more than 30% of qualified person-weeks requiring review overall;
- no more than 40% requiring review in any store-week unless a documented
  target breach explains the exception;
- less than 25% category reversal between consecutive qualified weeks; and
- exact guest reconciliation and sales/wine reconciliation within $0.01.

The `2026.08-v1` regression contract additionally verifies current-roster peer
exclusion, even-count medians, the 50-guest cap, sample rather than population
SD, exact boundary behavior, missing versus unqualified weeks, confidence and
overall-read mappings, rolling-window behavior, and 100% invariance of
`2026.07-v3` action outputs. Workbook tests require locked visible sheets, a
protected `veryHidden` calculation layer, no comments or unapproved drawings,
valid chart bindings, and substantive-digest coverage. These tests establish
determinism and controls, not fairness or causal validity.

The lowercase workbook password `redonion` is documented for authorized
operators as an accidental-edit convenience. It is not encryption or an access
control. The v4 substantive digest remains fail-closed for generated values and
formulas, sheet and table schema, validation and editability boundaries,
material cell styles, chart definitions and source bindings, meaningful
worksheet-view settings, hidden dimensions, internal links, drawings, external
links, and protection. To stay stable across a no-edit Excel save, it normalizes
only enumerated serializer defaults such as an explicit-versus-implicit `A1`
scroll origin and chart caches derived from protected formulas and cells. The
generator writes Excel's known persisted row heights up front; row heights and
column widths otherwise remain exact because a small or cumulative change can
alter rendered layout. The companion metadata-rich digest
still reports those serializer-sensitive differences as `metadata_drift=true`
with a Ready warning. That warning requires review or regeneration and does not
prove that every metadata change was harmless; corrections to substantive
content must be made in source/configuration and regenerated.
The immediately preceding v3 digest is accepted only through a one-way bridge
to the exact manifest-inventoried archived master; the next successful run
records the v4 contract.

The read-only `2026.07-v3` candidate replay covered 16 complete weeks, 96
business-date reports, 202 qualified person-weeks, and 25 store-week groups.
Nine qualified person-weeks generated a Context Review and none escalated to a
Coaching or Recognition Prompt. The overall review/action rate was 4.46%, and
the highest store-week rate was 25%. There were no consecutive candidate
pairs, so a reversal rate was not estimable. Prompt stability was also not
estimable because there were zero prompts; the methodology intentionally has
no minimum prompt quota. Automated regressions separately prove that
context-only metrics cannot change a person action.

Those results meet the observable numerical acceptance limits for the
candidate. The 24 validated backfill files were migrated through the protected
history transaction, and the exact post-migration replay matched these
aggregate results before managed workbooks were published. No separate
business-owner approval record has been verified. Until one is documented and
verified, authorized operational use of the people-review signals remains
blocked; technical migration, replay, and publication do not by themselves
establish approval.

These are operational stability and alert-quality tests, not proof of
statistical, causal, or demographic fairness. No minimum prompt quota is
allowed. Failure to meet an acceptance limit keeps the methodology under
review and must not be hidden by changing the evaluated cohort.

At 26 verified weeks, review threshold stability, seasonality, location and
sample-volume differences, manager dispositions, reversals, and disputed
signals. More advanced modeling requires a separately approved design with an
independently defined outcome and appropriate validation data.

## Historical Backfill

The 16-week candidate calibration incorporates four older complete
Tuesday-Sunday weeks retrieved once from 24 original Marketing Vitals TM
attachments: March 24-29, April 7-12, April 14-19, and April 21-26, 2026.
Selection uses the embedded business date, not the filename or email date.
Forwarded duplicates, derived `Check_Wine` workbooks, Store reports, Monday
reports, `No Data Available` workbooks, legacy incompatible schemas, and
conflicting same-date files are excluded.

The retrieval is read-only and temporary. Attachments are staged outside the
repository and live Dropbox folders, validated, migrated through the protected
history transaction, and removed from staging after verification. Email bodies
and message identifiers are not retained. The system has no ongoing Gmail
connector and a weekly run never accesses Gmail.

The maintainer-only `--rebuild-from-history` operation rebuilds managed outputs
from verified canonical history without processing active weekly inputs. It
preserves management fields by header and advances the generated archive,
manifest chain, published files, and trusted head only as one successful
transaction. Failure leaves the prior managed state unchanged.

## Monitoring, Disputes, And Change Control

Managers and employees may dispute source data, identity, peer comparability,
data-sufficiency labels, consistency calculations, missing operating context,
thresholds, or the signal itself. Preserve the
original signal and final disposition. Correct source defects through the
protected transaction; document context or methodology disagreements rather
than rewriting prior evidence.

The business owner governs permitted use and resolves context disputes. The
technical maintainer owns source, integrity, implementation, and release
controls. Material changes to either methodology's inputs, cohorts, scoring,
thresholds, persistence, terminology, or permitted use require:

1. a new methodology version;
2. updated model card and governance documentation;
3. regression and historical backtesting;
4. business-owner approval; and
5. a protected release and deployment.

A `2026.08-v1` change must also prove that the people-review action path is
unchanged unless a separately versioned and approved `2026.07` successor is
explicitly in scope.
