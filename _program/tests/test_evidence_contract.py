from __future__ import annotations

import json
from argparse import Namespace
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook

import red_onion_weekly_metrics as metrics
from red_onion_runtime import write_json_atomic


def sample_signal() -> dict:
    return {
        "Entity Key": "server|rc richmond|alex|coaching",
        "Priority": "High",
        "Status": "Open",
        "Owner": "Pat Manager",
        "Due Date": date(2026, 7, 30),
        "Location": "RC Richmond",
        "Person / Area": "Alex",
        "Action": "Coach Now",
        "Signal": "Falling / Below Benchmark / 8-week Declining",
        "Why It Matters": "Watch: check average",
        "Recommended Next Step": "Review recent shifts.",
        "Performance Level": "Below Benchmark",
        "Momentum": "Falling",
        "Confidence": "High",
        "Last Seen": date(2026, 7, 19),
        "_evidence_week_ends": ["2026-07-12", "2026-07-19"],
        "_source_evidence": [
            {
                "source_file": "Daily Report.xlsx",
                "sha256": "a" * 64,
                "format": ".xlsx",
                "parser_engine": "openpyxl",
                "report_date_source": "Workbook Date(s) field",
                "report_date": "2026-07-19",
            }
        ],
        "_metric_evidence": {
            "recent_metric_scores": {"check_average": -2},
            "guest_count": 42,
        },
    }


def current_action() -> dict:
    action = metrics.enrich_management_signal(sample_signal())
    action.update(
        {
            "Action ID": "ABC123",
            "First Seen": date(2026, 7, 19),
            "Weeks Open": 1,
            "Manager Notes": "",
            "Signal State": "Current",
        }
    )
    return action


def test_stable_codes_and_evidence_id_are_deterministic() -> None:
    first = metrics.enrich_management_signal(sample_signal())
    second = metrics.enrich_management_signal(sample_signal())

    assert first["Action Code"] == "COACH_NOW"
    assert first["Reason Code"] == "SERVER_FALLING_BELOW_BENCHMARK"
    assert first["Evidence ID"] == second["Evidence ID"]
    assert first["Evidence Week Ends"] == "2026-07-12, 2026-07-19"
    sources = json.loads(first["Evidence Sources"])
    assert sources[0]["sha256"] == "a" * 64


def test_management_signal_matches_reviewed_semantic_golden() -> None:
    result = metrics.enrich_management_signal(sample_signal())
    fixture = Path(__file__).parent / "fixtures" / "management_signal_golden.json"
    expected = json.loads(fixture.read_text(encoding="utf-8"))

    assert {field: result[field] for field in expected} == expected


def test_action_focus_links_to_editable_board_and_evidence_is_read_only() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    action = current_action()
    metrics.write_action_tracking_sheet(
        wb, "Action Board", [action], editable=True
    )
    metrics.write_action_focus_sheet(wb, [action])
    metrics.write_evidence_detail_sheet(wb, [action], [])

    focus = wb["Action Focus"]
    evidence = wb["Evidence Detail"]
    assert focus["A6"].value == "High"
    assert focus["K6"].hyperlink.target == "#'Action Board'!C5"
    assert evidence["A5"].value == action["Evidence ID"]
    assert evidence["A5"].hyperlink.target == "#'Action Board'!C5"
    assert json.loads(evidence["L5"].value)[0]["parser_engine"] == "openpyxl"
    assert all(
        cell.protection.locked is not False
        for cell in evidence._cells.values()
    )


def test_verified_evidence_uses_live_editable_action_fields(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    action = current_action()
    metrics.write_action_tracking_sheet(
        wb, "Action Board", [action], editable=True
    )
    metrics.write_evidence_detail_sheet(wb, [action], [])
    board = wb["Action Board"]
    board["D5"] = "Blocked"
    board["E5"] = "Current Manager"
    board["F5"] = date(2026, 8, 6)
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    workbook_path = output_dir / "Red_Onion_Server_Master.xlsx"
    wb.save(workbook_path)
    wb.close()

    archive_dir = tmp_path / "archive"
    archive_dir.mkdir()
    manifest_path = archive_dir / "manifest.json"
    manifest_path.write_text("{}", encoding="utf-8")
    manifest_sha256 = "b" * 64
    manifest_payload = {
        "run_id": "run-1",
        "created_at_utc": "2026-07-23T12:00:00+00:00",
        "master_generated_content_sha256": "c" * 64,
        "provenance": {
            "git": {"commit": "d" * 40},
            "effective_config_sha256": "e" * 64,
        },
    }
    monkeypatch.setattr(
        metrics,
        "latest_integrity_manifest_path",
        lambda archive: manifest_path,
    )
    monkeypatch.setattr(
        metrics,
        "verify_integrity_anchor",
        lambda archive, anchor: (manifest_path.resolve(), manifest_sha256),
    )
    monkeypatch.setattr(
        metrics,
        "verify_integrity_state",
        lambda archive, output, manifest: (
            manifest_payload,
            manifest_sha256,
        ),
    )
    monkeypatch.setattr(
        metrics, "managed_master_workbook_path", lambda output: workbook_path
    )
    monkeypatch.setattr(
        metrics, "validate_management_workbook", lambda workbook, digest: None
    )
    args = Namespace(
        output_dir=str(output_dir),
        archive_dir=str(archive_dir),
        integrity_anchor_dir=str(tmp_path / "anchors"),
    )

    _, rows, _ = metrics.verified_evidence_source(args)

    assert rows[0]["Status"] == "Blocked"
    assert rows[0]["Owner"] == "Current Manager"
    assert metrics.as_date(rows[0]["Due Date"]) == date(2026, 8, 6)
    assert rows[0]["Action Code"] == action["Action Code"]
    assert rows[0]["Evidence ID"] == action["Evidence ID"]


def sample_package() -> metrics.ManagementEvidencePackageV1:
    record = metrics.EvidenceRecord(
        action_id="ABC123",
        evidence_id="EVIDENCE1",
        action_code="COACH_NOW",
        reason_code="SERVER_FALLING_BELOW_BENCHMARK",
        location="RC Richmond",
        person_or_area="Alex",
        priority="High",
        status="Open",
        owner="Pat Manager",
        due_date="2026-07-30",
        recommended_next_step="Review recent shifts.",
        why_it_matters="Watch: check average",
        evidence_week_ends="2026-07-12, 2026-07-19",
        evidence_sources=(
            {
                "source_file": "Daily Report.xlsx",
                "sha256": "a" * 64,
            },
        ),
        metric_evidence={"guest_count": 42},
        methodology_version=metrics.MANAGEMENT_METHODOLOGY_VERSION,
    )
    return metrics.ManagementEvidencePackageV1(
        source={
            "manifest_path": "manifest.json",
            "manifest_sha256": "b" * 64,
            "manifest_run_id": "run-1",
            "manifest_created_at_utc": "2026-07-23T12:00:00+00:00",
            "workbook_file": "Red_Onion_Server_Master.xlsx",
            "workbook_generated_content_sha256": "c" * 64,
            "generator_commit": "d" * 40,
            "effective_config_sha256": "e" * 64,
            "methodology_version": metrics.MANAGEMENT_METHODOLOGY_VERSION,
        },
        records=(record,),
        retention_delete_after="2027-07-23",
    )


def test_candidate_requires_exact_approval_and_promotes_identical_bytes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    package = sample_package()
    monkeypatch.setattr(
        metrics, "build_management_evidence_package", lambda args: package
    )
    candidate = tmp_path / "review" / "candidate.json"
    args = Namespace(archive_dir=str(tmp_path / "archive"))
    paths = metrics.stage_management_evidence(args, candidate)
    candidate_bytes = candidate.read_bytes()
    payload = json.loads(candidate_bytes)
    assert payload["retention"]["days"] == 365
    assert payload["distribution"]["automatic_upload"] is False
    assert payload["distribution"]["automatic_send"] is False
    assert payload["records"][0]["person_or_area"] == "Alex"

    template = json.loads(paths[2].read_text(encoding="utf-8"))
    template.update(
        {
            "approved_by": "Authorized Manager",
            "approved_at_utc": "2026-07-23T13:00:00+00:00",
            "purpose": "Approved AI management coaching review",
        }
    )
    approval = tmp_path / "review" / "approval.json"
    write_json_atomic(approval, template)
    (tmp_path / "archive").mkdir()
    monkeypatch.setattr(
        metrics,
        "verified_evidence_source",
        lambda args: (package.source, [], tmp_path / "master.xlsx"),
    )

    promoted = metrics.promote_approved_management_evidence(
        args, candidate, approval
    )

    assert promoted[0].read_bytes() == candidate_bytes
    receipt = json.loads(promoted[1].read_text(encoding="utf-8"))
    assert receipt["approved_by"] == "Authorized Manager"
    assert receipt["automatic_upload"] is False
    assert receipt["automatic_send"] is False


def test_approval_hash_mismatch_is_rejected_before_archive_creation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    package = sample_package()
    monkeypatch.setattr(
        metrics, "build_management_evidence_package", lambda args: package
    )
    args = Namespace(archive_dir=str(tmp_path / "archive"))
    candidate = tmp_path / "candidate.json"
    paths = metrics.stage_management_evidence(args, candidate)
    approval = json.loads(paths[2].read_text(encoding="utf-8"))
    approval.update(
        {
            "candidate_sha256": "0" * 64,
            "approved_by": "Authorized Manager",
            "approved_at_utc": "2026-07-23T13:00:00+00:00",
            "purpose": "Review",
        }
    )
    approval_path = tmp_path / "approval.json"
    write_json_atomic(approval_path, approval)

    with pytest.raises(metrics.IntegrityError, match="does not match"):
        metrics.promote_approved_management_evidence(
            args, candidate, approval_path
        )

    assert not (tmp_path / "archive").exists()
