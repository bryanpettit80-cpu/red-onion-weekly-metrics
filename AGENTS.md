# Repository Instructions

This project builds the Red Onion weekly metrics snapshot. Keep the repository focused on repeatable weekly reporting and avoid committing generated customer-facing outputs unless the user explicitly asks for them to be versioned.

## Workflow

- Treat raw source files and exported reports as local operating data unless they are already intentionally tracked.
- Preserve the existing PowerShell launcher and single-script workflow unless a requested change updates the operator flow.
- Keep dependencies minimal and documented in `requirements.txt`.
- Do not introduce external services or credentials for a simple weekly run.

## Validation

Run this before committing changes:

```powershell
python -m pip install -r requirements.txt
python -m py_compile red_onion_weekly_metrics.py
```

