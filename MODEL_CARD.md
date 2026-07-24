# Red Onion Weekly Metrics Model Card

## Identity And Status

- **System:** Red Onion Weekly Metrics
- **Release contract:** `v0.4.0` (methodology changed from `v0.3.2`);
  publication requires the v3 validation and business-owner approval below
- **Methodology:** `2026.07-v3`
- **System type:** deterministic, rule-based observational coaching signal
- **Owners:** Red Onion business owner and designated technical maintainer
- **Decision authority:** accountable human manager; never the automation

This model card describes two intentionally separate layers. The operations
layer presents trends, volume, and descriptive context. The people-review layer
covers person-level Recent Movement, Peer Comparison, Context Review, Coaching
Prompt, and Recognition Prompt behavior.

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
- monitoring aggregate signal quality during the validation pilot.

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
release. For `2026.07-v3`, the Sales/Guest and Wine Percentage bands retain the
verified 12-week calibration:

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
| Sales/Guest | $11.50 / $18.50 | $11.00 / $16.50 |
| Wine Percentage | 4.1 / 5.7 percentage points | 4.1 / 5.8 percentage points |

Movement calibration used 169 self-history-qualified person-weeks, or 338
action-driving metric deviations; peer calibration used 153
peer-reference-qualified person-weeks, or 306 action-driving metric deviations.
The 12 complete weeks span April 28 through July 19, 2026, and the two
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

## Signal And Persistence Rules

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

Every generated row begins with Review Disposition `Pending Review`. An
authorized manager checks the source, identity, peer cohort, comparable-work
context, recurring drivers, and stability evidence, then records one of:

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

## Validation And Acceptance

Release validation includes unit, workbook-contract, integrity, migration, and
historical backtests. The initial 12-week pilot must demonstrate:

- 100% person-action invariance when only Rate of Sale, Ticket Time, Check
  Count, Sales/Check, or Guests/Check changes;
- 100% action invariance when excluded rows or descriptive ranks change;
- 100% leave-one-active-day stability for Coaching and Recognition Prompts;
- no more than 30% of qualified person-weeks requiring review overall;
- no more than 40% requiring review in any store-week unless a documented
  target breach explains the exception;
- less than 25% category reversal between consecutive qualified weeks; and
- exact guest reconciliation and sales/wine reconciliation within $0.01.

The `2026.07-v2` historical baseline covered 153 qualified person-weeks. Eleven
generated a review item (7.19% overall); the highest store-week rate was 25%.
There were no reversals in the one consecutive candidate transition, and no
escalated prompt failed the leave-one-active-day requirement. Zero Coaching or
Recognition Prompts were escalated. Those figures describe the prior four-
metric methodology and do not validate `2026.07-v3`. A separate v3 historical
replay is required before the next protected release; neither methodology
imposes a minimum prompt quota.

These are operational stability and alert-quality tests, not proof of
statistical, causal, or demographic fairness. No minimum prompt quota is
allowed. Failure to meet an acceptance limit keeps the methodology under
review and must not be hidden by changing the evaluated cohort.

At 26 verified weeks, review threshold stability, seasonality, location and
sample-volume differences, manager dispositions, reversals, and disputed
signals. More advanced modeling requires a separately approved design with an
independently defined outcome and appropriate validation data.

## Historical Backfill

The initial 12-week calibration may incorporate five older Tuesday-Sunday
weeks retrieved once from original Marketing Vitals TM attachments. Selection
uses the embedded business date, not the filename or email date. Forwarded
duplicates, Store reports, Monday reports, and unrelated attachments are
excluded.

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
missing operating context, thresholds, or the signal itself. Preserve the
original signal and final disposition. Correct source defects through the
protected transaction; document context or methodology disagreements rather
than rewriting prior evidence.

The business owner governs permitted use and resolves context disputes. The
technical maintainer owns source, integrity, implementation, and release
controls. Material changes to inputs, cohorts, scoring, thresholds,
persistence, terminology, or permitted use require:

1. a new methodology version;
2. updated model card and governance documentation;
3. regression and historical backtesting;
4. business-owner approval; and
5. a protected release and deployment.
