# Red Onion maintainer guide

Operators use `Run Weekly Snapshot.cmd`. Maintainers own configuration,
program releases, unusual corrections, and recovery. The ordinary workflow
does not require Git, a release tag, a clean checkout, historical derivative
hashes, or a machine-local trust anchor.

## Architecture and releases

[Program notes](./_program/README.md) map module responsibilities.
Calculation and workbook behavior remain in the existing generator; ordinary
orchestration, management input, and recoverable publication have small,
separate modules.

## Development and tests

Use a source checkout outside the live Dropbox workspace. Do not point tests,
experiments, or Cloud tasks at the numbered operational folders. The tests
create synthetic workbooks and isolated temporary workspaces; private source
reports are not required. Supported package runtimes are Python 3.10-3.12;
use Python 3.12 for a new development environment.

From the repository root, create an isolated environment on Windows:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

On Linux, including a Cloud setup workspace:

```bash
python3.12 -m venv .venv
source .venv/bin/activate
```

Then use the same commands on either platform:

```text
cd _program
python -m pip install --require-hashes -r requirements.lock
python -m pip install -e ".[dev]"
python -m pip check
python red_onion_config.py --config red_onion_config.json
python -m pytest -q
python -m compileall -q .
```

Run `git diff --check` from the repository root before committing. Production
dependencies are pinned and hash-checked; the development extra currently
allows `pytest>=8` and the build backend is not locked. Record the actual Python
and installed package versions when diagnosing a difference between runs.
There is no separate configured linter. Linux skips tests requiring Windows
PowerShell; CI separately runs launcher behavior tests in the
`windows-launcher-test` job.

### Codex Cloud development

Current [Cloud environments](https://learn.chatgpt.com/docs/environments/cloud-environments)
provide published reusable development workspaces. Legacy environments remain
relevant to Code Review and GitHub integrations. Inspect both inventories and
reuse the environment associated with this repository before creating another.
Keep a legacy review environment if those integrations still use it.

Configure the development environment for this repository and its reviewed
`main` branch, Python 3.12, and the installation commands above. Use the actual
checkout directory reported by setup; no fixed `/workspace` path is required.
There is no web server, background service, startup command, application secret,
or required application environment variable. Each task should activate the
virtual environment, or call its Python executable explicitly. Install packages
from `pypi.org` and `files.pythonhosted.org`; repository access uses GitHub.
Review the existing network and privacy settings before any change. A personal
environment should be limited to its owner unless sharing is approved.

After setup tests pass, save and publish or republish. Start a **fresh task**
from that published environment, record its environment identifier and
`git rev-parse HEAD`, verify `python --version` and `python -m pip check`, and
rerun configuration validation, the test suite, and compilation. Record the
commands, counts, skips, and failures. Saved settings or local/CI success alone
do not establish that the published Cloud environment works.

Cloud tasks use synthetic data. They do not inherit this computer's Dropbox
files, sign-ins, scheduled tasks, local Python runtime, or Excel session.
Do not run the operator `.cmd`/PowerShell launcher, process live intake, import
private history, or upload operating workbooks for routine Cloud development.
Windows launcher verification remains a Windows CI/local responsibility.

### Source releases

For a material source release:

1. Make a focused branch and update the package version.
2. From `_program`, install `.[dev]`, run `python -m pytest -q`, and run
   `python -m py_compile red_onion_weekly_metrics.py`.
3. Validate configuration and run `git diff --check`.
4. Review the patch and CI, including Python 3.10-3.12 and Windows launcher
   checks. Record what was actually validated.
5. Publish the reviewed source through the repository's normal release process.
6. Deploy source deliberately while no reporting run is active. Use a
   fast-forward update for a maintained checkout, or replace the program source
   from the reviewed release. Preserve the operational folders and config.
7. The tracked `00 START HERE - Red Onion Weekly Metrics.txt` is the operator
   guide template. If its shared outer copy needs updating, review that exact
   destination and obtain approval for the live-folder change. Back up the old
   guide, copy the reviewed template beside the outer launcher, and compare
   hashes. Updating the repository alone does not refresh that outer copy.

Git/release review controls code deployment. They are not daily prerequisites.
Do not commit source workbooks, finished reports, reporting state, credentials,
or machine-local environments. A local source/release bundle can be built with
`Build-RecoveryBundle.ps1`; portable operational backup is covered separately
in [RECOVERY.md](RECOVERY.md).

## Configuration and environment

```powershell
python _program\red_onion_config.py --config _program\red_onion_config.json
python _program\red_onion_weekly_metrics.py --validate-config
.\_program\Run-WeeklySnapshot.ps1 -RebuildEnvironment -HealthCheck
```

The rebuild/check combination installs locked dependencies and checks the
workspace without processing intake. Ordinary runs reuse a valid environment
and rebuild it if missing or stale. Plain `-HealthCheck` is read-only and
does not install. The environment under
`%LOCALAPPDATA%\RedOnionMetrics` is disposable; do not copy it between PCs.

Keep `requirements.txt`, `pyproject.toml`, the hashed `requirements.lock`,
and `requirements-constraints.txt` aligned. Regenerate a lock deliberately
using the documented Python range and review the dependency diff:

```powershell
python -m pip install pip-tools==7.5.2
python -m piptools compile requirements.txt --constraint requirements-constraints.txt --generate-hashes --allow-unsafe --output-file requirements.lock
```

Run those dependency commands from `_program`. A dependency update is a
source-maintenance task; ordinary reporting does not download a new release.

`closed_dates` is an explicit per-location list of verified Tuesday-Sunday
closures, not a way to suppress missing-report messages. `person_identity_aliases`
maps exact raw names to verified canonical names per location for descriptive
scorecards only. It cannot map shared/numbered IDs to people or chain aliases.
Same-week merges require disjoint daily evidence; ambiguous overlaps stop.
Neither setting should be inferred from a person's apparent absence.

## Health, result, and readiness are different

```powershell
.\_program\Run-WeeklySnapshot.ps1 -HealthCheck
python _program\red_onion_weekly_metrics.py --health-check-json
```

Current health output checks configuration, local folders, accepted reporting
state, and pending recovery. It does not prove a new week is complete, that
cloud sync finished, or that a recipient can open a workbook. A journal may
need reconciliation by the next ordinary invocation.

The current JSON attempts and committed receipts support local diagnostics.
Do not treat them as a promised external API. Complete-input readiness and a
coordinator result envelope are proposed in
[COORDINATOR_CONTRACT.md](COORDINATOR_CONTRACT.md), **design only**.

## Corrections and management inputs

For interactive correction review:

```powershell
& '.\Run Weekly Snapshot.cmd' -ReviewCorrections
```

The review shows business-value changes and records one choice per conflicting
date. It includes keeping accepted facts when available. Reviewed selections
are bound to the current alternatives. A changed selection set requires review
again. A noninteractive caller can provide `-CorrectionsFile <review.json>`;
it must never invent a choice or auto-select the newest filename.

Management is read by table/column names and stable Action IDs. An Excel save,
formatting change, or master rename is allowed. If different master copies
exist, select the intended one with `-ManagementWorkbook <path>`; it must
be in Finished Reports. Preserve the other copy until its edits are reconciled.

After initial import, saved generated actions/history/evidence remain
authoritative. The importer merges targets, the owner roster, and only the
seven editable Current Actions fields. A missing master can be rebuilt from
the last committed management snapshot, but workbook edits made after that
snapshot require the latest workbook copy. Never claim those unsaved-to-state
edits were recovered from an older receipt.

## History import and rebuild

First test these operations on a complete copy of the workspace. Explicit path
arguments prevent accidentally choosing a live default. For example, from the
repository root:

```powershell
python _program\red_onion_weekly_metrics.py --input-dir "D:\Review\01 Daily Reports - Drop Here" --output-dir "D:\Review\02 Finished Reports" --archive-dir "D:\Review\03 Archive" --migrate-history-only --migrate-history-from "D:\Approved History"
```

This copies validated source bytes into portable state and accepts history
without generating reports or deleting originals. Duplicate facts count once.
Conflicting corrections use the same review path. Partial history can be
imported, but only complete reporting weeks can be published.
An import retains the existing publication identity and records a pending
refresh. The next ordinary run refreshes the master and complete affected weeks
even when no new files are in the drop folder.

Rebuild from accepted history with the same three explicit folder arguments
and `--rebuild-from-history`. Rebuild ignores active intake, preserves current
management inputs, and regenerates the latest complete restaurant reports and
master. `--migrate-history-from` may be combined with rebuild when needed.
Correcting older facts through incoming reports updates affected historical
restaurant reports as well as the current master.

The first portable invocation reads existing canonical
`processed-daily-reports` history and imports it. It does not rewrite old
manifests, regenerate a trust chain, or require historical derivative files.
With empty intake, it saves history and the current management inputs without
changing reports. A subsequent ordinary run refreshes complete accepted history.

## Local evidence export

The existing two-stage employee-evidence workflow remains separate from
weekly report generation:

```powershell
python _program\red_onion_weekly_metrics.py --export-management-evidence C:\SecureReview\candidate.json
python _program\red_onion_weekly_metrics.py --export-management-evidence C:\SecureReview\candidate.json --approval-file C:\SecureReview\approval.json
```

First review the exact candidate and fingerprint, then complete the generated
approval template for that candidate and purpose. The second command promotes
those exact reviewed bytes locally. A changed candidate needs a new review.
No upload, email, draft, or deletion is performed. See
[DATA_GOVERNANCE.md](DATA_GOVERNANCE.md).

## Legacy forensic maintenance

The retained `--legacy-integrity` path verifies the old manifest chain,
canonical raw inventory, generated archive, workbook digest, and machine
anchor. `--initialize-integrity-baseline` and
`--rebind-restored-integrity-anchor` are also legacy operations. They are
not portable setup, repair, or routine reporting steps.

Use a preserved pre-migration workspace for legacy maintenance. Portable
writers leave old evidence untouched; legacy writers are blocked from altering
a workspace with committed portable state. Do not create a replacement
baseline to bypass a discrepancy. If a historical investigation requires
anchor rebind, restore the exact old chain and its independently saved anchor
into an isolated copy, then use the legacy recovery command and verify the
result there. See [RECOVERY.md](RECOVERY.md).
