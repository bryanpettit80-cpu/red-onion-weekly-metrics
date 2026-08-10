from __future__ import annotations

import json
from argparse import Namespace
from pathlib import Path

import pytest

import red_onion_weekly_metrics as metrics
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
    monkeypatch.setattr(
        metrics,
        "verify_integrity_state",
        lambda archive, output, selected_manifest: (
            {
                "kind": "integrity-baseline",
                "master_generated_content_sha256": None,
                "published_output_inventory": [],
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
        lambda archive, output, selected_manifest: (
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
