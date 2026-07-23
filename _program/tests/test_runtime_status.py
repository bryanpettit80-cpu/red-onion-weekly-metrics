from __future__ import annotations

import json
from argparse import Namespace
from pathlib import Path

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
    recorder.succeed("Published exact staged bytes.", details={"generated_files": ["one.xlsx"]})

    payload = json.loads(attempt_path.read_text(encoding="utf-8"))
    assert payload["schema_version"] == 1
    assert payload["outcome"] == "Success"
    assert payload["stage"] == "Complete"
    assert payload["readiness"]["integrity"] == "Ready"
    assert payload["readiness"]["workbook"] == "Ready"
    assert payload["details"]["generated_files"] == ["one.xlsx"]
    status = status_path.read_text(encoding="utf-8")
    assert "Outcome: Success" in status
    assert "Run ID: run-123" in status
    assert not list(tmp_path.rglob("*.tmp"))


def test_failed_attempt_preserves_safe_error_without_recovery_claim(tmp_path: Path) -> None:
    recorder = RunAttemptRecorder(
        run_id="run-failed",
        operation="weekly-run",
        attempt_path=tmp_path / "attempt.json",
        status_path=tmp_path / "status.txt",
    )

    recorder.fail(ValueError("bad\x00input"))

    payload = json.loads((tmp_path / "attempt.json").read_text(encoding="utf-8"))
    assert payload["outcome"] == "Failed"
    assert payload["readiness"]["integrity"] == "Attention"
    assert payload["readiness"]["recovery"] == "NotChecked"
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
    assert payload["readiness"]["recovery"] == "NotChecked"
    assert payload["note"].startswith("Independent Google Drive backup freshness")
    assert list(tmp_path.iterdir()) == []
