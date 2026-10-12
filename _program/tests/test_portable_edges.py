"""Boundary checks that require real accepted history and report workbooks."""
from pathlib import Path
import shutil

import pytest

from red_onion_transactions import RecoveryRequired
import red_onion_weekly_metrics as weekly
import red_onion_workflow as workflow
from test_portable_workflow import (
    add_accepted_duplicate, arguments, make_workspace, output_hashes,
    published_workspace, published_copy, store_for,
)


def test_legacy_source_changed_after_capture_cannot_commit(published_workspace, tmp_path, monkeypatch):
    root = make_workspace(tmp_path / "Legacy migration")
    history = root / "Archive/processed-daily-reports"
    history.mkdir()
    for day, item in store_for(published_workspace).current()["days"].items():
        shutil.copy2(published_workspace / "Archive" / item["source"], history / f"Export {day}.xlsx")
    changed = next(history.glob("*.xlsx"))
    real_writer = weekly.write_public_workbook
    def changed_while_building(*args, **kwargs):
        result = real_writer(*args, **kwargs)
        changed.write_bytes(b"an archive replacement after capture")
        return result
    monkeypatch.setattr(weekly, "write_public_workbook", changed_while_building)
    args = arguments(root, rebuild_from_history=True)
    with pytest.raises(RecoveryRequired, match="History input changed"):
        workflow.run(args)
    assert store_for(root).current() is None
    assert output_hashes(root) == {}
    assert changed.read_bytes() == b"an archive replacement after capture"


def test_duplicate_restores_missing_current_public_report(published_copy):
    root = published_copy
    missing = root / "Finished Reports/Check_Wine_RVA101126.xlsx"
    missing.unlink()
    add_accepted_duplicate(root)
    args = arguments(root)
    assert len(workflow.run(args)) == 3
    assert args._run_attempt_recorder.outcome == "published"
    assert missing.is_file()


def test_portable_evidence_export_uses_saved_facts_and_current_management(published_copy):
    root = published_copy
    package = weekly.build_management_evidence_package(arguments(root)).to_dict()
    assert package["source"]["workbook_generated_content_digest_scheme"] == "portable-business-v1"
    assert package["source"]["manifest_run_id"] == store_for(root).current()["run_id"]
    assert package["distribution"]["automatic_send"] is False
    assert package["distribution"]["automatic_upload"] is False


def test_portable_health_check_is_read_only_and_does_not_claim_input_readiness(tmp_path):
    root = tmp_path / "Unavailable cloud workspace"
    args = arguments(root)
    result = weekly.build_health_check(args)
    assert result["overall"] == "Attention"
    assert "incoming-week completeness" in result["scope"]
    assert not root.exists()
