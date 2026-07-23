from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Mapping


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
            "recovery": RunReadiness.NOT_CHECKED.value,
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
            "Readiness:",
        ]
        for name, value in payload["readiness"].items():
            lines.append(f"  {name.title()}: {value}")
        return "\n".join(lines)

    def write(self) -> None:
        write_json_atomic(self.attempt_path, self.payload())
        if self.status_path is not None:
            write_text_atomic(self.status_path, self._status_text())

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
        self.readiness["workbook"] = (
            RunReadiness.READY.value
            if self.operation == "weekly-run"
            else RunReadiness.NOT_EVALUATED.value
        )
        if details:
            self.details.update(dict(details))
        self.write()

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
        self.write()
