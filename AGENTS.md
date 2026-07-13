# Repository Instructions

This project builds the Red Onion weekly metrics snapshot. Keep the repository focused on repeatable weekly reporting and avoid committing generated customer-facing outputs unless the user explicitly asks for them to be versioned.

## Workflow

- Treat raw source files and exported reports as local operating data unless they are already intentionally tracked.
- Keep the deployed Dropbox parent operator-friendly: `01 Daily Reports - Drop Here`, `02 Finished Reports`, `03 Archive`, one root cmd launcher, and `Red Onion Weekly Metrics Automation` for the Git repository.
- Preserve `Run Weekly Snapshot.cmd` as the outer operator launcher. The tracked launcher must also work when the repository is used standalone.
- Keep dependencies minimal and documented in `_program\requirements.txt` and `_program\pyproject.toml`.
- Do not introduce external services or credentials for a simple weekly run.
- Do not add Darden fiscal calendar logic, gift-card workflows, Gmail import, or monthly close behavior unless explicitly requested.

## Validation

Run this before committing changes:

```powershell
cd _program
python -m pip install -e ".[dev]"
python -m pytest -q
python -m py_compile red_onion_weekly_metrics.py
```
