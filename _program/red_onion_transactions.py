"""Recoverable local publication with a portable current-run receipt.

Checksums here protect a copy or an interrupted transaction. Old presentation
files and historical receipts are not prerequisites for the next weekly run.
Only one machine writes at a time; callers hold the existing OS workflow lock.
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import shutil
import tempfile
from typing import Any, Callable

from red_onion_integrity import sha256_file
from red_onion_runtime import utc_now, write_json_atomic


class RecoveryRequired(ValueError):
    """An interrupted operation conflicts with a subsequent external edit."""


def child(root: Path, relative: str) -> Path:
    parsed = PurePosixPath(relative)
    if not relative or "\\" in relative or ":" in relative or parsed.is_absolute() or any(
        part in {".", ".."} for part in parsed.parts
    ):
        raise RecoveryRequired("Invalid relative path in reporting state.")
    result = root.joinpath(*parsed.parts)
    if not result.resolve().is_relative_to(root.resolve()):
        raise RecoveryRequired("Reporting state points outside its workspace.")
    for path in (result, *result.parents):
        if path == root.parent:
            break
        if path.is_symlink():
            raise RecoveryRequired(f"Reporting path is a symbolic link: {path.name}")
    return result


def json_read(path: Path) -> dict[str, Any]:
    def unique(items: list[tuple[str, Any]]) -> dict[str, Any]:
        value: dict[str, Any] = {}
        for key, item in items:
            if key in value:
                raise RecoveryRequired(f"Duplicate field {key} in {path.name}.")
            value[key] = item
        return value
    with path.open(encoding="utf-8") as handle:
        payload = json.load(handle, object_pairs_hook=unique)
    if not isinstance(payload, dict):
        raise RecoveryRequired(f"Expected an object in {path.name}.")
    return payload


def atomic_bytes(path: Path, data: bytes, *, exclusive: bool = False) -> str:
    """Write and verify in the destination filesystem before an atomic rename."""
    path.parent.mkdir(parents=True, exist_ok=True)
    expected = hashlib.sha256(data).hexdigest()
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix=".report-copy-", dir=path.parent, delete=False) as handle:
            temporary = Path(handle.name)
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        if sha256_file(temporary) != expected:
            raise OSError(f"Could not verify the staged copy of {path.name}.")
        if exclusive:
            # Windows rename is no-clobber. POSIX needs an exclusive hard-link
            # installation to preserve a file created after our existence check.
            if os.name == "nt":
                os.rename(temporary, path)
            else:
                os.link(temporary, path)
                temporary.unlink()
        else:
            os.replace(temporary, path)
        temporary = None
        return expected
    finally:
        if temporary is not None:
            try:
                temporary.unlink()
            except OSError:
                pass


class ReportStore:
    def __init__(self, archive_dir: Path, output_dir: Path, input_dir: Path):
        self.archive_dir = archive_dir.resolve()
        self.output_dir = output_dir.resolve()
        self.input_dir = input_dir.resolve()
        self.root = child(self.archive_dir, "Reporting State")
        self.runs = child(self.root, "runs")
        self.pointer = child(self.root, "current.json")

    def current(self) -> dict[str, Any] | None:
        if not self.pointer.exists():
            if self.runs.exists():
                committed = []
                for path in self.runs.glob("*/journal.json"):
                    journal = json_read(path)
                    if journal.get("phase") in {"committed", "cleanup_pending", "complete"}:
                        committed.append(journal)
                if committed:
                    raise RecoveryRequired(
                        "The current reporting receipt is missing. Run recovery to restore "
                        "it from the most recent completed journal."
                    )
            return None
        pointer = json_read(self.pointer)
        if pointer.get("schema_version") != 1 or not isinstance(pointer.get("run_id"), str):
            raise RecoveryRequired("The current reporting receipt has an unsupported schema.")
        commit_path = child(self.runs, pointer["run_id"] + "/commit.json")
        if sha256_file(commit_path) != pointer.get("commit_sha256"):
            raise RecoveryRequired("The current reporting receipt is damaged. Restore its saved run record.")
        commit = json_read(commit_path)
        if commit.get("schema_version") != 1 or commit.get("run_id") != pointer["run_id"]:
            raise RecoveryRequired("The current run record does not match its receipt.")
        return commit

    def _write_pointer(self, run_id: str) -> None:
        commit = child(self.runs, run_id + "/commit.json")
        write_json_atomic(self.pointer, {
            "schema_version": 1, "run_id": run_id,
            "commit_sha256": sha256_file(commit),
        })

    def recover(self) -> list[str]:
        """Reconcile journals before accepting another run, including cleanup."""
        warnings: list[str] = []
        if not self.runs.exists():
            return warnings
        journals = [(p, json_read(p)) for p in self.runs.glob("*/journal.json")]
        journals.sort(key=lambda item: (str(item[1].get("started_at_utc", "")), item[0].as_posix()))
        if not self.pointer.exists():
            completed = [j for _, j in journals if j.get("phase") in {"committed", "cleanup_pending", "complete"}]
            if completed:
                # Sequential machines may have different clock settings. The
                # successor relationship determines the latest receipt.
                predecessors = {j.get("previous_run_id") for j in completed}
                heads = [j for j in completed if j["run_id"] not in predecessors]
                if len(heads) != 1:
                    raise RecoveryRequired("Several completed histories remain without a current receipt. Preserve the copies and select the intended workspace before recovery.")
                latest = heads[0]
                self._write_pointer(latest["run_id"])
                warnings.append("Restored the current reporting receipt from its completed journal.")
        pointer_id = json_read(self.pointer).get("run_id") if self.pointer.exists() else None
        for path, journal in journals:
            phase = journal.get("phase")
            if phase in {"complete", "rolled_back"}:
                continue
            if journal.get("schema_version") != 1 or journal.get("run_id") != path.parent.name:
                raise RecoveryRequired("Unrecognized publication journal; no report was overwritten.")
            if phase in {"committed", "cleanup_pending"} or journal["run_id"] == pointer_id:
                warnings.extend(self.finish_cleanup(path, journal))
            elif phase == "publishing":
                self.rollback(path, journal)
                warnings.append("Restored the previous complete report after an interrupted publication.")
            else:
                raise RecoveryRequired(f"Unrecognized publication phase: {phase}")
        return warnings

    def rollback(self, journal_path: Path, journal: dict[str, Any]) -> None:
        """Undo only this transaction's bytes; preserve any newer external edit."""
        for item in reversed(journal["outputs"]):
            destination = child(self.output_dir, item["name"])
            displaced = child(self.output_dir, item["displaced"])
            quarantine = child(self.output_dir, item["displaced"] + ".rollback")
            if quarantine.exists():
                if sha256_file(quarantine) != item["sha256"]:
                    raise RecoveryRequired(f"A newer edit in {quarantine.name} needs review; all copies are preserved.")
                quarantine.unlink()
            current_hash = sha256_file(destination) if destination.exists() else None
            old_hash = item["old_sha256"]
            backup = child(journal_path.parent, item["backup"]) if item["backup"] else None
            if old_hash is not None and (backup is None or not backup.is_file() or sha256_file(backup) != old_hash):
                raise RecoveryRequired(f"Previous-good copy unavailable for {destination.name}.")
            if current_hash is not None and current_hash not in {item["sha256"], old_hash}:
                raise RecoveryRequired(
                    f"{destination.name} changed during an interrupted run. Its newer edit "
                    f"and the previous copy in {journal_path.parent / 'previous'} are preserved. "
                    "Move the conflicting file aside before retrying recovery."
                )
            if current_hash == item["sha256"] and current_hash != old_hash:
                # Capture the actual file before deleting our publication. A
                # manager save between hash and rename must remain recoverable.
                os.rename(destination, quarantine)
                if sha256_file(quarantine) != item["sha256"]:
                    if not destination.exists():
                        atomic_bytes(destination, quarantine.read_bytes(), exclusive=True)
                    raise RecoveryRequired(f"{destination.name} has a newer edit; it and {quarantine.name} were preserved.")
                quarantine.unlink()
                current_hash = None
            if old_hash is not None and current_hash is None:
                atomic_bytes(destination, backup.read_bytes(), exclusive=True)
            if displaced.exists():
                if sha256_file(displaced) != old_hash:
                    raise RecoveryRequired(f"A changed displaced copy was preserved: {displaced.name}")
                displaced.unlink()
        journal["phase"] = "rolled_back"
        journal["finished_at_utc"] = utc_now()
        write_json_atomic(journal_path, journal)

    def finish_cleanup(self, journal_path: Path, journal: dict[str, Any]) -> list[str]:
        warnings: list[str] = []
        for item in journal["inputs"]:
            source = child(self.input_dir, item["name"])
            archive = child(self.archive_dir, item["archive"])
            if not source.exists():
                continue
            try:
                if not archive.is_file() or sha256_file(archive) != item["sha256"]:
                    raise OSError("The saved input copy is not available for verification.")
                if sha256_file(source) != item["sha256"]:
                    warnings.append(f"A changed input was preserved for the next run: {source.name}")
                    # This replacement is no longer ours to clean up.
                    continue
                # Rename then verify before deletion: a replacement between check
                # and rename must be preserved, not lost during cleanup.
                quarantine = child(self.input_dir, item["quarantine"])
                if quarantine.exists():
                    if sha256_file(quarantine) != item["sha256"]:
                        raise OSError(f"A different quarantined input remains at {quarantine.name}.")
                    quarantine.unlink()
                os.rename(source, quarantine)
                if sha256_file(quarantine) != item["sha256"]:
                    if not source.exists():
                        os.rename(quarantine, source)
                    raise OSError("Input changed during cleanup; preserved for review.")
                quarantine.unlink()
            except OSError as exc:
                warnings.append(f"Input cleanup pending for {source.name}: {exc}")
        # A power loss can occur after rename but before deletion.
        for item in journal["inputs"]:
            quarantine = child(self.input_dir, item["quarantine"])
            if quarantine.exists():
                try:
                    archive = child(self.archive_dir, item["archive"])
                    if sha256_file(quarantine) != item["sha256"] or sha256_file(archive) != item["sha256"]:
                        raise OSError("Quarantined input differs from its saved copy.")
                    quarantine.unlink()
                except OSError as exc:
                    warnings.append(f"Input cleanup pending for {quarantine.name}: {exc}")
        for item in journal["outputs"]:
            displaced = child(self.output_dir, item["displaced"])
            if displaced.exists():
                try:
                    if sha256_file(displaced) != item["old_sha256"]:
                        raise OSError("Displaced output changed; it has been preserved.")
                    displaced.unlink()
                except OSError as exc:
                    warnings.append(f"Temporary output cleanup pending: {displaced.name}: {exc}")
        staging = child(journal_path.parent, "staging")
        if staging.exists():
            try:
                shutil.rmtree(staging)
            except OSError as exc:
                warnings.append(f"Temporary staging cleanup pending: {exc}")
        # These are warnings after the durable commit, never failed publication.
        journal["phase"] = "cleanup_pending" if any("pending" in w for w in warnings) else "complete"
        journal["warnings"] = warnings
        journal["finished_at_utc"] = utc_now()
        try:
            write_json_atomic(journal_path, journal)
        except OSError as exc:
            warnings.append(f"Publication succeeded; cleanup status could not be saved: {exc}")
        return warnings

    def publish(
        self, run_id: str, staged: list[Path], commit: dict[str, Any],
        captured_inputs: list[dict[str, Any]], expected_outputs: dict[str, str | None],
        *, checkpoint: Callable[[str], None] | None = None,
    ) -> tuple[list[Path], list[str]]:
        """Publish a complete set and commit once; fault injection is test-only."""
        notify = checkpoint or (lambda stage: None)
        run_root = child(self.runs, run_id)
        run_root.mkdir(parents=True, exist_ok=True)
        journal_path = run_root / "journal.json"
        if journal_path.exists():
            raise RecoveryRequired("This run already has a publication journal.")
        outputs: list[dict[str, Any]] = []
        for path in staged:
            destination = child(self.output_dir, path.name)
            old_hash = sha256_file(destination) if destination.exists() else None
            if old_hash != expected_outputs.get(path.name):
                raise RecoveryRequired(f"{path.name} changed while the report was being built. Rerun to retain its latest edits.")
            backup_name = "previous/" + path.name if old_hash is not None else None
            if backup_name:
                original = destination.read_bytes()
                if atomic_bytes(child(run_root, backup_name), original, exclusive=True) != old_hash:
                    raise RecoveryRequired(f"{path.name} changed while its previous-good copy was saved.")
            saved = child(run_root, "outputs/" + path.name)
            new_hash = atomic_bytes(saved, path.read_bytes(), exclusive=True)
            outputs.append({
                "name": path.name, "sha256": new_hash, "old_sha256": old_hash,
                "backup": backup_name, "snapshot": "outputs/" + path.name,
                "displaced": f".report-{run_id}-{path.name}.previous",
            })
        commit.update({"schema_version": 1, "run_id": run_id, "outputs": outputs})
        write_json_atomic(run_root / "commit.json", commit)
        journal = {
            "schema_version": 1, "run_id": run_id, "phase": "publishing",
            "started_at_utc": utc_now(), "previous_run_id": commit.get("previous_run_id"),
            "outputs": outputs, "inputs": captured_inputs,
        }
        write_json_atomic(journal_path, journal)
        committed = False
        try:
            notify("journal_saved")
            for index, item in enumerate(outputs):
                target = child(self.output_dir, item["name"])
                displaced = child(self.output_dir, item["displaced"])
                if item["old_sha256"] is not None:
                    os.rename(target, displaced)
                    if sha256_file(displaced) != item["old_sha256"]:
                        # It contains the actual newer edit; restore it, not the older backup.
                        if not target.exists():
                            os.rename(displaced, target)
                        raise RecoveryRequired(f"{target.name} changed during publication; the newer edit was preserved.")
                notify(f"output_{index}_displaced")
                saved = child(run_root, item["snapshot"])
                atomic_bytes(target, saved.read_bytes(), exclusive=True)
                if sha256_file(target) != item["sha256"]:
                    raise OSError(f"Publication verification failed for {target.name}.")
                notify(f"output_{index}_published")
            # Never replace the receipt if another local process bypassed the lock.
            current = self.current()
            current_id = current["run_id"] if current else None
            if current_id != commit.get("previous_run_id"):
                raise RecoveryRequired("The current reporting receipt changed during publication.")
            notify("before_commit")
            self._write_pointer(run_id)
            committed = True
            notify("after_commit")
        except Exception:
            if not committed:
                # A filesystem wrapper may raise after installing the pointer.
                # The durable receipt, not the Python return, determines commit.
                installed = self.current()
                committed = installed is not None and installed["run_id"] == run_id
            if not committed:
                self.rollback(journal_path, journal)
            else:
                journal["phase"] = "cleanup_pending"
                try:
                    write_json_atomic(journal_path, journal)
                except OSError:
                    pass
                return [child(self.output_dir, item["name"]) for item in outputs], [
                    "Publication succeeded; post-commit cleanup will resume on the next run."
                ]
            raise
        journal["phase"] = "committed"
        try:
            write_json_atomic(journal_path, journal)
            warnings = self.finish_cleanup(journal_path, journal)
        except OSError as exc:
            warnings = [f"Publication succeeded; cleanup will resume on the next run: {exc}"]
        return [child(self.output_dir, item["name"]) for item in outputs], warnings
