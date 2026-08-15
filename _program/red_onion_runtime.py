from __future__ import annotations

import json
import os
import re
import stat
import tempfile
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Mapping


_OPERATIONAL_STATUS_REPLACE_RETRY_DELAYS = (0.05, 0.15)


class RunStage(str, Enum):
    CONFIG_VALIDATED = "ConfigValidated"
    WAITING_FOR_LOCK = "WaitingForLock"
    INTEGRITY_PREFLIGHT = "IntegrityPreflight"
    READING_INPUTS = "ReadingInputs"
    BUILDING_WORKBOOKS = "BuildingWorkbooks"
    PUBLISHING = "Publishing"
    COMMITTING_MANIFEST = "CommittingManifest"
    COMPLETE = "Complete"
    FAILED = "Failed"


class RunReadiness(str, Enum):
    NOT_EVALUATED = "NotEvaluated"
    RUNNING = "Running"
    READY = "Ready"
    ATTENTION = "Attention"
    NOT_CHECKED = "NotChecked"
    EXTERNAL_CHECK_REQUIRED = "ExternalCheckRequired"


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def safe_message(value: Any, *, limit: int = 1000) -> str:
    text = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]+", " ", str(value))
    return " ".join(text.split())[:limit]


def write_json_atomic(path: Path, payload: Mapping[str, Any]) -> Path:
    path = path.absolute()
    path.parent.mkdir(parents=True, exist_ok=True)
    serialized = (
        json.dumps(
            payload,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            indent=2,
        )
        + "\n"
    )
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            prefix=f".{path.name}.",
            suffix=".tmp",
            dir=str(path.parent),
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
            handle.write(serialized)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
        temporary = None
    finally:
        if temporary is not None:
            try:
                temporary.unlink()
            except FileNotFoundError:
                pass
    return path


def write_text_atomic(path: Path, text: str) -> Path:
    path = path.absolute()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            prefix=f".{path.name}.",
            suffix=".tmp",
            dir=str(path.parent),
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
            handle.write(text.rstrip() + "\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
        temporary = None
    finally:
        if temporary is not None:
            try:
                temporary.unlink()
            except FileNotFoundError:
                pass
    return path


def _verify_operational_status_bytes(path: Path, expected: bytes) -> None:
    persisted = path.read_bytes()
    if persisted != expected:
        raise OSError(
            f"Operational status verification failed after writing {path}."
        )


def _assert_opened_operational_status_identity(path: Path, handle: Any) -> None:
    path_stat = os.lstat(path)
    opened_stat = os.fstat(handle.fileno())
    if stat.S_ISLNK(path_stat.st_mode):
        raise OSError(
            f"Refusing in-place operational status rewrite through a link: {path}"
        )
    if not stat.S_ISREG(path_stat.st_mode) or not stat.S_ISREG(opened_stat.st_mode):
        raise OSError(
            f"Operational status fallback requires a regular file: {path}"
        )
    if path_stat.st_nlink != 1 or opened_stat.st_nlink != 1:
        raise OSError(
            "Operational status fallback requires exactly one filesystem link: "
            f"{path}"
        )
    path_identity = (path_stat.st_dev, path_stat.st_ino)
    opened_identity = (opened_stat.st_dev, opened_stat.st_ino)
    if path_identity != opened_identity:
        raise OSError(
            f"Operational status destination changed while it was opened: {path}"
        )


def _rewrite_existing_operational_status(path: Path, content: bytes) -> None:
    """Rewrite an existing status file when a cloud placeholder blocks replace."""

    # A Dropbox placeholder is a regular file reparse point. Reject actual links so
    # the narrowly scoped fallback cannot be redirected outside the managed path.
    if path.is_symlink():
        raise OSError(
            f"Refusing in-place operational status rewrite through a link: {path}"
        )
    with path.open("r+b") as handle:
        _assert_opened_operational_status_identity(path, handle)
        handle.seek(0)
        handle.write(content)
        handle.truncate()
        handle.flush()
        os.fsync(handle.fileno())
    _verify_operational_status_bytes(path, content)


def _hydrate_operational_status_for_retry(path: Path) -> None:
    if path.is_symlink():
        raise OSError(
            f"Refusing operational status replacement through a link: {path}"
        )
    try:
        with path.open("rb") as handle:
            handle.read(1)
    except OSError:
        # A bounded replace retry can still succeed if the placeholder disappeared
        # between checks. If it does not, the verified fallback reports its own error.
        pass


def _write_operational_status_bytes(path: Path, content: bytes) -> Path:
    """Persist a run-status artifact despite an existing cloud-placeholder race.

    Run-attempt JSON and LAST RUN STATUS are operational status artifacts that are
    rewritten at every stage. They normally retain the same atomic replace behavior
    as other writers. Some Dropbox Windows placeholders reject replacing an existing
    reparse-point destination with ``PermissionError``. Only for an already-existing,
    single-link regular operational status file, fall back to an fsynced in-place
    rewrite and verify the exact persisted bytes. Critical workbook, manifest,
    anchor, and other general atomic writers do not use this helper.
    """

    path = path.absolute()
    path.parent.mkdir(parents=True, exist_ok=True)
    destination_existed = path.is_file()
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            prefix=f".{path.name}.",
            suffix=".tmp",
            dir=str(path.parent),
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        try:
            os.replace(temporary, path)
        except PermissionError as replace_error:
            if not destination_existed:
                raise
            for delay in _OPERATIONAL_STATUS_REPLACE_RETRY_DELAYS:
                _hydrate_operational_status_for_retry(path)
                time.sleep(delay)
                try:
                    os.replace(temporary, path)
                except PermissionError as retry_error:
                    replace_error = retry_error
                else:
                    temporary = None
                    _verify_operational_status_bytes(path, content)
                    break
            else:
                try:
                    _rewrite_existing_operational_status(path, content)
                except OSError as fallback_error:
                    raise fallback_error from replace_error
        else:
            temporary = None
            _verify_operational_status_bytes(path, content)
    finally:
        if temporary is not None:
            try:
                temporary.unlink()
            except FileNotFoundError:
                pass
    return path


def write_operational_json(path: Path, payload: Mapping[str, Any]) -> Path:
    serialized = (
        json.dumps(
            payload,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            indent=2,
        )
        + "\n"
    ).encode("utf-8")
    return _write_operational_status_bytes(path, serialized)


def write_operational_text(path: Path, text: str) -> Path:
    serialized = (text.rstrip() + "\n").encode("utf-8")
    return _write_operational_status_bytes(path, serialized)


@dataclass
class RunAttemptRecorder:
    run_id: str
    operation: str
    attempt_path: Path
    status_path: Path | None
    started_at_utc: str = field(default_factory=utc_now)
    stage: RunStage = RunStage.CONFIG_VALIDATED
    readiness: dict[str, str] = field(
        default_factory=lambda: {
            "release": RunReadiness.NOT_EVALUATED.value,
            "integrity": RunReadiness.RUNNING.value,
            "workbook": RunReadiness.NOT_EVALUATED.value,
            "distribution": RunReadiness.NOT_EVALUATED.value,
            "recovery": RunReadiness.EXTERNAL_CHECK_REQUIRED.value,
        }
    )
    message: str = "Configuration validated."
    details: dict[str, Any] = field(default_factory=dict)
    completed_at_utc: str | None = None
    outcome: str = "Running"

    def payload(self) -> dict[str, Any]:
        return {
            "schema_version": 1,
            "run_id": self.run_id,
            "operation": self.operation,
            "outcome": self.outcome,
            "stage": self.stage.value,
            "started_at_utc": self.started_at_utc,
            "completed_at_utc": self.completed_at_utc,
            "readiness": dict(self.readiness),
            "message": safe_message(self.message),
            "details": dict(self.details),
        }

    def _status_text(self) -> str:
        payload = self.payload()
        lines = [
            "RED ONION WEEKLY METRICS - LAST RUN STATUS",
            f"Outcome: {payload['outcome']}",
            f"Stage: {payload['stage']}",
            f"Run ID: {payload['run_id']}",
            f"Started (UTC): {payload['started_at_utc']}",
            f"Finished (UTC): {payload['completed_at_utc'] or 'Still running'}",
            f"Message: {payload['message']}",
            "",
            "Run verification:",
        ]
        labels = {
            "release": "Release",
            "integrity": "Integrity",
            "workbook": "Workbook",
            "distribution": "Local publication",
        }
        for name in ("release", "integrity", "workbook", "distribution"):
            lines.append(
                f"  {labels[name]}: {payload['readiness'].get(name, RunReadiness.NOT_EVALUATED.value)}"
            )
        lines.extend(
            [
                "",
                "External assurance:",
                (
                    "  Independent recovery: "
                    f"{payload['readiness'].get('recovery', RunReadiness.EXTERNAL_CHECK_REQUIRED.value)}"
                ),
                (
                    "    Verify the current independent backup and restore-test evidence "
                    "separately; the local weekly run does not access Google Drive."
                ),
                "",
                (
                    "Scope: Local publication verifies exact bytes for managed per-location "
                    "workbooks plus the protected generated content of the master workbook "
                    "in the configured 02 Finished Reports folder. It does not verify approved "
                    "editable master values, LAST RUN STATUS.txt, Dropbox cloud sync, or "
                    "recipient access."
                ),
            ]
        )
        return "\n".join(lines)

    def write(self) -> None:
        write_operational_json(self.attempt_path, self.payload())
        if self.status_path is not None:
            write_operational_text(self.status_path, self._status_text())

    def update(
        self,
        stage: RunStage,
        message: str,
        *,
        readiness: Mapping[str, RunReadiness | str] | None = None,
        details: Mapping[str, Any] | None = None,
    ) -> None:
        self.stage = stage
        self.message = safe_message(message)
        if readiness:
            self.readiness.update(
                {
                    key: value.value if isinstance(value, RunReadiness) else str(value)
                    for key, value in readiness.items()
                }
            )
        if details:
            self.details.update(dict(details))
        self.write()

    def succeed(self, message: str, *, details: Mapping[str, Any] | None = None) -> None:
        self.outcome = "Success"
        self.completed_at_utc = utc_now()
        self.stage = RunStage.COMPLETE
        self.message = safe_message(message)
        self.readiness["integrity"] = RunReadiness.READY.value
        if self.operation in {"weekly-run", "history-rebuild"}:
            self.readiness["workbook"] = RunReadiness.READY.value
        if details:
            self.details.update(dict(details))
        # The attempt log is the authoritative post-commit result. A human-readable
        # status-file refresh is useful but must not turn an already committed
        # publication and manifest into a reported operational failure.
        write_operational_json(self.attempt_path, self.payload())
        if self.status_path is not None:
            try:
                write_operational_text(self.status_path, self._status_text())
            except OSError as exc:
                self.details["last_run_status_write_warning"] = safe_message(
                    f"{type(exc).__name__}: {exc}"
                )
                try:
                    write_operational_json(self.attempt_path, self.payload())
                except OSError:
                    # The successful attempt was already persisted before the
                    # optional status-file write was attempted.
                    pass

    def fail(self, exc: BaseException) -> None:
        self.outcome = "Failed"
        self.completed_at_utc = utc_now()
        self.stage = RunStage.FAILED
        self.message = safe_message(f"{type(exc).__name__}: {exc}")
        self.readiness.update(
            {
                "integrity": RunReadiness.ATTENTION.value,
                "workbook": RunReadiness.ATTENTION.value,
            }
        )
        if self.readiness.get("distribution") == RunReadiness.RUNNING.value:
            self.readiness["distribution"] = RunReadiness.ATTENTION.value
        self.write()
