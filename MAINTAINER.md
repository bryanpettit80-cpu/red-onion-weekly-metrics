# Red Onion Maintainer Guide

This repository is the released generator inside the operator-friendly Dropbox
workspace. Operators use the root `Run Weekly Snapshot.cmd`; maintainers own
release, configuration, recovery, and incident work.

## Architecture

- `_program/red_onion_config.py`: standard-library schema validation.
- `_program/red_onion_runtime.py`: atomic attempt logs and last-run status.
- `_program/red_onion_integrity.py`: inventories, manifests, hashes, and atomic
  chained-manifest primitives.
- `_program/red_onion_weekly_metrics.py`: parsing, deterministic business rules,
  workbook generation, evidence export, publication, and transactional rollback.
- `_program/Run-WeeklySnapshot.ps1`: release preflight, environment validation,
  explicit environment rebuild, read-only health-check routing, and Python launch.

Generated workbooks, source reports, archives, machine-local environments, and
recovery bundles are operating data and are not committed.

## Release Procedure

1. Start from a clean branch based on `origin/main`.
2. Update the package version for a material release.
3. From `_program`, run:

   ```powershell
   python -m pip install -e ".[dev]"
   python -m pytest -q
   python -m py_compile red_onion_weekly_metrics.py
   ```

4. Run `git diff --check`.
5. Use a pull request and require all Python 3.10-3.12 and Windows launcher
   checks to pass.
6. Merge, update the deployed Dropbox checkout by fast-forward only, and prove
   `HEAD == origin/main`. The launcher isolates Python bytecode lookup and does
   not require maintainer cleanup of ignored `__pycache__` files.
7. Create an annotated `vX.Y.Z` tag and GitHub Release.
8. Build and independently retain a recovery bundle:

   ```powershell
   .\Build-RecoveryBundle.ps1 -DestinationDirectory C:\Recovery\RedOnion
   ```

No generated customer workbook belongs in a GitHub Release.

## Configuration

Validate before review or deployment:

```powershell
python _program\red_onion_config.py --config _program\red_onion_config.json
python _program\red_onion_weekly_metrics.py --validate-config
```

Unknown keys, duplicate JSON keys, empty or duplicate location codes, invalid
types, non-finite/negative thresholds, and inconsistent trend windows fail.
Configuration is validated before runtime folders, locks, or environments are
created.

## Dependencies and Runtime

`requirements.txt` lists direct dependencies. `requirements.lock` contains the
reviewed transitive graph and hashes. `requirements-constraints.txt` keeps the
single lock compatible with Python 3.10-3.12.

Regenerate deliberately:

```powershell
python -m pip install pip-tools==7.5.2
python -m piptools compile requirements.txt `
  --constraint requirements-constraints.txt `
  --generate-hashes --allow-unsafe `
  --output-file requirements.lock
```

The weekly launcher does not reinstall packages. It verifies the Python
identity, lock hash, installed direct versions, and `pip check`. It rebuilds
only when state is missing/stale or a maintainer explicitly runs:

```powershell
.\_program\Run-WeeklySnapshot.ps1 -RebuildEnvironment -HealthCheck
```

That combination rebuilds and validates the environment without processing
weekly reports. Using `-RebuildEnvironment` by itself rebuilds first and then
continues with the requested weekly run.

## Health and Evidence

The local health check is read-only:

```powershell
.\_program\Run-WeeklySnapshot.ps1 -HealthCheck
python _program\red_onion_weekly_metrics.py --health-check-json
```

It verifies exact managed per-location workbook bytes plus the protected
substantive content of the master workbook through the integrity manifest.
Generated values and formulas, workbook schema, editable-cell boundaries,
display-affecting formats, chart content and bindings, hidden dimensions,
internal-link destinations, drawings, external links, and protection controls
remain fail-closed. Proven equivalent Excel serialization defaults are reported
without blocking when substantive content still matches the manifest-pinned
archived master. Approved editable master values, `LAST RUN STATUS.txt`,
Dropbox sync, and recipient access are outside
that claim. It never claims that independent recovery is healthy; recovery
remains `ExternalCheckRequired` until the current private backup and
restore-test evidence are verified separately.

Replacement-machine recovery uses
`-RebindRestoredIntegrityAnchor <SOURCE_ANCHOR_JSON>`. The recovery-only command
verifies the backed-up manifest hash against the fully restored chain, raw
inventory, generated archive, published outputs, and workbook digest before it
writes a new path-bound anchor and audit receipt. It never runs the weekly
workflow. See `RECOVERY.md` for the required sequence.

Approved management evidence is two-stage. First stage a candidate. Review its
JSON and fingerprint, complete a copy of the approval template, then promote
the same candidate bytes:

```powershell
python _program\red_onion_weekly_metrics.py `
  --export-management-evidence C:\SecureReview\candidate.json

python _program\red_onion_weekly_metrics.py `
  --export-management-evidence C:\SecureReview\candidate.json `
  --approval-file C:\SecureReview\approval.json
```

Promotion is local only. There is no automatic upload, send, or deletion.
