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
python -m pip install -e ".[dev]"
python -m pytest -q
python -m py_compile red_onion_weekly_metrics.py
```
