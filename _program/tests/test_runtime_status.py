from __future__ import annotations

import json
import os
from argparse import Namespace
from pathlib import Path

import pytest

import red_onion_weekly_metrics as metrics
import red_onion_runtime as runtime
from red_onion_runtime import RunAttemptRecorder, RunReadiness, RunStage


def test_attempt_log_and_human_status_are_atomic_and_versioned(tmp_path: Path) -> None:
    attempt_path = tmp_path / "archive" / "attempt.json"
    status_path = tmp_path / "output" / metrics.LAST_RUN_STATUS_FILE
    recorder = RunAttemptRecorder(
        run_id="run-123",
        operation="weekly-run",
        attempt_path=attempt_path,
        status_path=status_path,
    )

    recorder.update(
        RunStage.BUILDING_WORKBOOKS,
        "Building workbooks",
        readiness={"workbook": RunReadiness.RUNNING},
    )
    recorder.update(
        RunStage.COMMITTING_MANIFEST,
        "Verified exact local publication.",
        readiness={"distribution": RunReadiness.READY},
    )
    recorder.succeed("Published exact staged bytes.", details={"generated_files": ["one.xlsx"]})

    payload = json.loads(attempt_path.read_text(encoding="utf-8"))
    assert payload["schema_version"] == 1
    assert payload["outcome"] == "Success"
    assert payload["stage"] == "Complete"
    assert payload["readiness"]["integrity"] == "Ready"
    assert payload["readiness"]["workbook"] == "Ready"
    assert payload["readiness"]["distribution"] == "Ready"
    assert payload["readiness"]["recovery"] == "ExternalCheckRequired"
    assert payload["details"]["generated_files"] == ["one.xlsx"]
    status = status_path.read_text(encoding="utf-8")
    assert "Outcome: Success" in status
    assert "Run ID: run-123" in status
    assert "Local publication: Ready" in status
    assert "Independent recovery: ExternalCheckRequired" in status
    assert "managed per-location workbooks" in status
    assert "approved editable master values" in status
    assert "Dropbox cloud sync" in status
    assert not list(tmp_path.rglob("*.tmp"))


@pytest.mark.parametrize(
    ("operation", "expected"),
    [
        ("weekly-run", "Ready"),
        ("history-rebuild", "Ready"),
        ("history-migration", "NotEvaluated"),
        ("integrity-baseline", "NotEvaluated"),
    ],
)
def test_success_marks_workbook_ready_only_for_workbook_generating_operations(
    tmp_path: Path, operation: str, expected: str
) -> None:
    recorder = RunAttemptRecorder(
        run_id=f"run-{operation}",
        operation=operation,
        attempt_path=tmp_path / f"{operation}.json",
        status_path=None,
    )

    recorder.succeed("Complete.")

    payload = json.loads(
        (tmp_path / f"{operation}.json").read_text(encoding="utf-8")
    )
    assert payload["readiness"]["workbook"] == expected


def test_success_remains_success_when_final_human_status_write_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    attempt_path = tmp_path / "archive" / "attempt.json"
    status_path = tmp_path / "output" / metrics.LAST_RUN_STATUS_FILE
    recorder = RunAttemptRecorder(
        run_id="run-status-warning",
        operation="weekly-run",
        attempt_path=attempt_path,
        status_path=status_path,
    )

    def reject_status_write(path: Path, text: str) -> Path:
        raise PermissionError("status file is open in another application")

    monkeypatch.setattr(runtime, "write_operational_text", reject_status_write)

    recorder.succeed("Publication and manifest committed.")

    payload = json.loads(attempt_path.read_text(encoding="utf-8"))
    assert payload["outcome"] == "Success"
    assert payload["stage"] == "Complete"
    assert payload["readiness"]["integrity"] == "Ready"
    assert payload["readiness"]["workbook"] == "Ready"
    assert "PermissionError" in payload["details"][
        "last_run_status_write_warning"
    ]
    assert not status_path.exists()


def test_recorder_rewrites_existing_cloud_placeholders_when_replace_is_denied(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    attempt_path = tmp_path / "archive" / "attempt.json"
    status_path = tmp_path / "output" / metrics.LAST_RUN_STATUS_FILE
    recorder = RunAttemptRecorder(
        run_id="run-placeholder-race",
        operation="weekly-run",
        attempt_path=attempt_path,
        status_path=status_path,
    )
    recorder.write()

    original_replace = runtime.os.replace
    denied_destinations: list[Path] = []

    def reject_replacing_existing_status(source: object, destination: object) -> None:
        destination_path = Path(destination)
        if (
            destination_path in {attempt_path, status_path}
            and destination_path.exists()
        ):
            denied_destinations.append(destination_path)
            raise PermissionError(13, "cloud placeholder rejected replace", destination)
        original_replace(source, destination)

    monkeypatch.setattr(runtime.os, "replace", reject_replacing_existing_status)
    monkeypatch.setattr(runtime.time, "sleep", lambda delay: None)

    recorder.update(
        RunStage.PUBLISHING,
        "Publishing verified bytes.",
        readiness={"distribution": RunReadiness.RUNNING},
    )
    recorder.succeed("Publication and manifest committed.")

    payload = json.loads(attempt_path.read_text(encoding="utf-8"))
    assert payload["outcome"] == "Success"
    assert payload["stage"] == "Complete"
    assert payload["completed_at_utc"] is not None
    assert "Outcome: Success" in status_path.read_text(encoding="utf-8")
    retries_per_write = 1 + len(runtime._OPERATIONAL_STATUS_REPLACE_RETRY_DELAYS)
    assert denied_destinations.count(attempt_path) == 2 * retries_per_write
    assert denied_destinations.count(status_path) == 2 * retries_per_write
    assert not list(tmp_path.rglob("*.tmp"))


def test_operational_status_first_write_permission_error_fails_closed(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    status_path = tmp_path / "new-status.txt"

    def reject_replace(source: object, destination: object) -> None:
        raise PermissionError(13, "destination rejected replace", destination)

    monkeypatch.setattr(runtime.os, "replace", reject_replace)

    with pytest.raises(PermissionError, match="destination rejected replace"):
        runtime.write_operational_text(status_path, "Running")

    assert not status_path.exists()
    assert not list(tmp_path.rglob("*.tmp"))


def test_operational_status_retries_atomic_replace_before_fallback(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    status_path = tmp_path / "status.txt"
    status_path.write_text("Running\n", encoding="utf-8")
    original_replace = runtime.os.replace
    replace_attempts = 0
    delays: list[float] = []

    def transient_placeholder_race(source: object, destination: object) -> None:
        nonlocal replace_attempts
        replace_attempts += 1
        if replace_attempts < 3:
            raise PermissionError(13, "cloud placeholder rejected replace", destination)
        original_replace(source, destination)

    def reject_fallback(path: Path, content: bytes) -> None:
        raise AssertionError("in-place fallback should not run after a successful retry")

    monkeypatch.setattr(runtime.os, "replace", transient_placeholder_race)
    monkeypatch.setattr(runtime.time, "sleep", delays.append)
    monkeypatch.setattr(
        runtime,
        "_rewrite_existing_operational_status",
        reject_fallback,
    )

    runtime.write_operational_text(status_path, "Success")

    assert status_path.read_bytes() == b"Success\n"
    assert replace_attempts == 3
    assert delays == list(runtime._OPERATIONAL_STATUS_REPLACE_RETRY_DELAYS)
    assert not list(tmp_path.rglob("*.tmp"))


def test_operational_status_fallback_failure_propagates_and_cleans_temp(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    status_path = tmp_path / "status.txt"
    original = b"Running\n"
    status_path.write_bytes(original)
    original_open = runtime.Path.open

    def reject_replace(source: object, destination: object) -> None:
        raise PermissionError(13, "cloud placeholder rejected replace", destination)

    def reject_destination_open(
        path: Path,
        mode: str = "r",
        *args: object,
        **kwargs: object,
    ):
        if path == status_path and mode == "r+b":
            raise PermissionError(13, "cloud placeholder remained unavailable", path)
        return original_open(path, mode, *args, **kwargs)

    monkeypatch.setattr(runtime.os, "replace", reject_replace)
    monkeypatch.setattr(runtime.time, "sleep", lambda delay: None)
    monkeypatch.setattr(runtime.Path, "open", reject_destination_open)

    with pytest.raises(PermissionError, match="remained unavailable"):
        runtime.write_operational_text(status_path, "Success")

    with original_open(status_path, "rb") as handle:
        assert handle.read() == original
    assert not list(tmp_path.rglob("*.tmp"))


def test_operational_status_fallback_rejects_link_destination(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    status_path = tmp_path / "status.txt"
    status_path.write_text("Running\n", encoding="utf-8")
    original = status_path.read_bytes()
    original_is_symlink = runtime.Path.is_symlink

    def reject_replace(source: object, destination: object) -> None:
        raise PermissionError(13, "destination rejected replace", destination)

    def report_managed_status_as_link(path: Path) -> bool:
        return path == status_path or original_is_symlink(path)

    monkeypatch.setattr(runtime.os, "replace", reject_replace)
    monkeypatch.setattr(runtime.Path, "is_symlink", report_managed_status_as_link)

    with pytest.raises(OSError, match="through a link"):
        runtime.write_operational_text(status_path, "Success")

    assert status_path.read_bytes() == original
    assert not list(tmp_path.rglob("*.tmp"))


def test_operational_status_fallback_rejects_hard_link_without_mutating_victim(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    victim = tmp_path / "victim.txt"
    victim_bytes = b"do-not-change\n"
    victim.write_bytes(victim_bytes)
    status_path = tmp_path / "status.txt"
    os.link(victim, status_path)

    def reject_replace(source: object, destination: object) -> None:
        raise PermissionError(13, "destination rejected replace", destination)

    monkeypatch.setattr(runtime.os, "replace", reject_replace)
    monkeypatch.setattr(runtime.time, "sleep", lambda delay: None)

    with pytest.raises(OSError, match="exactly one filesystem link"):
        runtime.write_operational_text(status_path, "Success")

    assert victim.read_bytes() == victim_bytes
    assert status_path.read_bytes() == victim_bytes
    assert not list(tmp_path.rglob("*.tmp"))


def test_operational_status_fallback_rejects_opened_file_identity_change(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    status_path = tmp_path / "status.txt"
    original = b"Running\n"
    status_path.write_bytes(original)
    other_path = tmp_path / "other.txt"
    other_path.write_bytes(b"Other\n")
    original_lstat = runtime.os.lstat

    def reject_replace(source: object, destination: object) -> None:
        raise PermissionError(13, "destination rejected replace", destination)

    def report_different_identity(path: object):
        if Path(path) == status_path:
            return original_lstat(other_path)
        return original_lstat(path)

    monkeypatch.setattr(runtime.os, "replace", reject_replace)
    monkeypatch.setattr(runtime.os, "lstat", report_different_identity)
    monkeypatch.setattr(runtime.time, "sleep", lambda delay: None)

    with pytest.raises(OSError, match="changed while it was opened"):
        runtime.write_operational_text(status_path, "Success")

    assert status_path.read_bytes() == original
    assert other_path.read_bytes() == b"Other\n"
    assert not list(tmp_path.rglob("*.tmp"))


def test_general_atomic_writer_does_not_use_operational_status_fallback(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    destination = tmp_path / "critical.json"
    original = b'{"state": "old"}\n'
    destination.write_bytes(original)

    def reject_replace(source: object, target: object) -> None:
        raise PermissionError(13, "destination rejected replace", target)

    monkeypatch.setattr(runtime.os, "replace", reject_replace)

    with pytest.raises(PermissionError, match="destination rejected replace"):
        runtime.write_json_atomic(destination, {"state": "new"})

    assert destination.read_bytes() == original
    assert not list(tmp_path.rglob("*.tmp"))


def test_failed_attempt_preserves_safe_error_without_recovery_claim(tmp_path: Path) -> None:
    recorder = RunAttemptRecorder(
        run_id="run-failed",
        operation="weekly-run",
        attempt_path=tmp_path / "attempt.json",
        status_path=tmp_path / "status.txt",
    )

    recorder.update(
        RunStage.PUBLISHING,
        "Publishing",
        readiness={"distribution": RunReadiness.RUNNING},
    )
    recorder.fail(ValueError("bad\x00input"))

    payload = json.loads((tmp_path / "attempt.json").read_text(encoding="utf-8"))
    assert payload["outcome"] == "Failed"
    assert payload["readiness"]["integrity"] == "Attention"
    assert payload["readiness"]["distribution"] == "Attention"
    assert payload["readiness"]["recovery"] == "ExternalCheckRequired"
    assert "\x00" not in payload["message"]


def test_health_check_is_read_only_when_runtime_folders_are_missing(
    tmp_path: Path,
) -> None:
    args = Namespace(
        input_dir=str(tmp_path / "input"),
        output_dir=str(tmp_path / "output"),
        archive_dir=str(tmp_path / "archive"),
        integrity_anchor_dir=str(tmp_path / "anchor"),
        config=str(tmp_path / "missing-config.json"),
    )

    payload = metrics.build_health_check(args)

    assert payload["overall"] == "Attention"
    assert payload["readiness"]["distribution"] == "Attention"
    assert payload["readiness"]["recovery"] == "ExternalCheckRequired"
    assert payload["note"].startswith("Independent Google Drive backup freshness")
    assert any(
        check["name"] == "Local publication" and check["status"] == "Attention"
        for check in payload["checks"]
    )
    assert list(tmp_path.iterdir()) == []


def test_health_check_does_not_call_empty_integrity_baseline_a_publication(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    input_dir = tmp_path / "input"
    output_dir = tmp_path / "output"
    archive_dir = tmp_path / "archive"
    for path in (input_dir, output_dir, archive_dir):
        path.mkdir()
    manifest = archive_dir / "baseline.json"
    manifest_sha256 = "a" * 64
    monkeypatch.setattr(
        metrics, "latest_integrity_manifest_path", lambda archive: manifest
    )
    monkeypatch.setattr(
        metrics, "integrity_anchor_exists", lambda archive, anchor: True
    )
    monkeypatch.setattr(
        metrics,
        "verify_integrity_anchor",
        lambda archive, anchor: (manifest.resolve(), manifest_sha256),
    )
    def verify_pinned_health_state(archive, output, selected_manifest, **kwargs):
        assert kwargs == {"allow_legacy_master_upgrade": True}
        return (
            {
                "kind": "integrity-baseline",
                "master_generated_content_sha256": None,
                "published_output_inventory": [],
            },
            manifest_sha256,
        )

    monkeypatch.setattr(
        metrics,
        "verify_integrity_state",
        verify_pinned_health_state,
    )

    payload = metrics.build_health_check(
        Namespace(
            input_dir=str(input_dir),
            output_dir=str(output_dir),
            archive_dir=str(archive_dir),
            integrity_anchor_dir=str(tmp_path / "anchor"),
            config=str(tmp_path / "missing-config.json"),
        )
    )

    assert payload["readiness"]["integrity"] == "Ready"
    assert payload["readiness"]["workbook"] == "Attention"
    assert payload["readiness"]["distribution"] == "Attention"
    assert payload["overall"] == "Attention"
    assert any(
        check["name"] == "Workbook" and check["status"] == "Attention"
        for check in payload["checks"]
    )


def test_health_check_reports_metadata_drift_as_structured_ready_warning(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    input_dir = tmp_path / "input"
    output_dir = tmp_path / "output"
    archive_dir = tmp_path / "archive"
    for path in (input_dir, output_dir, archive_dir):
        path.mkdir()
    manifest = archive_dir / "weekly.json"
    manifest_sha256 = "a" * 64
    monkeypatch.setattr(
        metrics, "latest_integrity_manifest_path", lambda archive: manifest
    )
    monkeypatch.setattr(
        metrics, "integrity_anchor_exists", lambda archive, anchor: True
    )
    monkeypatch.setattr(
        metrics,
        "verify_integrity_anchor",
        lambda archive, anchor: (manifest.resolve(), manifest_sha256),
    )
    monkeypatch.setattr(
        metrics,
        "verify_integrity_state",
        lambda archive, output, selected_manifest, **kwargs: (
            {
                "kind": "weekly-run",
                "master_generated_content_sha256": "b" * 64,
                "published_output_inventory": [{"path": "report.xlsx"}],
                "_master_metadata_drift": True,
            },
            manifest_sha256,
        ),
    )

    payload = metrics.build_health_check(
        Namespace(
            input_dir=str(input_dir),
            output_dir=str(output_dir),
            archive_dir=str(archive_dir),
            integrity_anchor_dir=str(tmp_path / "anchor"),
            config=str(tmp_path / "missing-config.json"),
        )
    )

    workbook_check = next(
        check for check in payload["checks"] if check["name"] == "Workbook"
    )
    assert workbook_check["status"] == "Ready"
    assert workbook_check["metadata_drift"] is True
    assert "metadata differs" in workbook_check["detail"]
    assert payload["readiness"]["workbook"] == "Ready"
