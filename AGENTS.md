# Repository Instructions

This project builds the Red Onion weekly metrics snapshot. Keep the repository focused on repeatable weekly reporting and avoid committing generated customer-facing outputs unless the user explicitly asks for them to be versioned.

## Workflow

- Treat raw source files and exported reports as local operating data unless they are already intentionally tracked.
- Keep the deployed Dropbox parent operator-friendly: `01 Daily Reports - Drop Here`, `02 Finished Reports`, `03 Archive`, one root cmd launcher, and `Red Onion Weekly Metrics Automation` for the Git repository.
- Preserve `Run Weekly Snapshot.cmd` as the outer operator launcher. The tracked launcher must also work when the repository is used standalone.
- Keep dependencies minimal and documented in `_program\requirements.txt` and `_program\pyproject.toml`.
- Do not introduce external services or credentials for a simple weekly run.
- Do not add Darden fiscal calendar logic, gift-card workflows, Gmail import, or monthly close behavior unless explicitly requested.

## Development and Cloud boundaries

- Use Python 3.12 for new development environments; the package currently supports 3.10-3.12.
- Work in a source checkout, with an isolated `.venv`. From `_program`, install the hashed `requirements.lock` before the editable development package and run `python -m pip check`.
- Tests use synthetic data and temporary folders. Do not use live Dropbox intake, finished reports, reporting state, or private employee data for routine development or Cloud validation.
- Linux/Cloud supports Python development and tests. Windows `.cmd`/PowerShell launchers are checked separately on Windows; a Linux skip is not proof of launcher behavior.
- No application secrets, external service, server startup, or scheduled task is needed for development. See [MAINTAINER.md](MAINTAINER.md#development-and-tests) for setup and fresh-task Cloud verification.
- Do not claim a Cloud environment is ready from local tests or saved settings. Record the published environment, tested revision, commands, and results from a fresh Cloud task.

## Validation

Activate the isolated development environment described in
[MAINTAINER.md](MAINTAINER.md#development-and-tests) before running these
commands. Do not install test dependencies into the workstation's base Python.
Run this before committing changes:

```powershell
cd _program
python -m pip install --require-hashes -r requirements.lock
python -m pip install -e ".[dev]"
python -m pip check
python red_onion_config.py --config red_onion_config.json
python -m pytest -q
python -m py_compile red_onion_weekly_metrics.py
```

Run `git diff --check` from the repository root. For Windows launcher changes, run `python -m pytest -q tests/test_launcher_integrity.py` from `_program` on Windows and record platform-specific skips honestly. CI remains the source for its complete compile list and platform matrix.

Windows launcher tests also require the launcher's interpreter discovery to resolve Python 3.10-3.12; running pytest in a supported virtual environment alone does not establish that.

## Project contracts

Read the relevant contract when changing its area: `MODEL_CARD.md` for calculations and evidence, `DATA_GOVERNANCE.md` for people-related outputs and allowed workbook edits, `RECOVERY.md` for operational recovery, and `MAINTAINER.md` for releases. `COORDINATOR_CONTRACT.md` describes a future design, not an available integration or implemented CLI.

## Code Review Rules

- Preserve complete Tuesday-Sunday coverage and explicit closure/correction decisions. An unavailable workspace or missing restaurant-day is not empty intake or an implicit closure; conflicting corrected business facts require the established review path.
- Preserve editable management inputs, stable Action IDs, accepted history, and publication/recovery behavior. Recompute generated metrics from accepted daily facts; merge only documented editable workbook fields. Do not discard manager edits or confuse a post-publication cleanup warning with a failed publication that should be repeated.
- Keep shared POS identities and operational-area aggregates out of person-level cohorts and actions. Preserve `Pending Review`, evidence eligibility, and documented human-review requirements; do not turn descriptive reporting into an automatic employment decision or transmit raw/private evidence.

## Pull requests

Use a task branch and open a focused draft PR against `main`, following [the PR template](.github/pull_request_template.md). Record checks actually run and material limitations. Merge only when requested and after required checks and reviews are satisfied; these instructions grant no standing merge authorization.
