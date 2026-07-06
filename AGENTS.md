# Repository Instructions

This project builds the Red Onion weekly metrics snapshot. Keep the repository focused on repeatable weekly reporting and avoid committing generated customer-facing outputs unless the user explicitly asks for them to be versioned.

## Workflow

- Treat raw source files and exported reports as local operating data unless they are already intentionally tracked.
- Keep the root folder operator-friendly: `Daily Reports`, `Output`, `Archive - Old Files`, one root cmd launcher, and `_program` for technical files.
- Preserve `Run Weekly Snapshot.cmd` as the root operator launcher unless a requested change updates the operator flow.
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
