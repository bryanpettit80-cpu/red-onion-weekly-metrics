"""Explicit period choices must preserve the ordinary complete-week contract."""
from datetime import date
from pathlib import Path
import shutil

import pytest

import red_onion_weekly_metrics as weekly
import red_onion_workflow as workflow
from test_portable_workflow import (
    arguments, make_workspace, output_hashes, public_totals, published_workspace,
    published_copy, store_for, write_week,
)


@pytest.mark.parametrize("selection", [
    {"week_end": "2026-10-09"},
    {"week_start": "2026-10-07"},
])
def test_explicit_partial_week_cannot_publish_from_accepted_history(published_copy, selection):
    root = published_copy
    before = output_hashes(root)
    previous = store_for(root).current()
    args = arguments(root, **selection)

    with pytest.raises(ValueError, match="Tuesday|Sunday|complete|full"):
        workflow.run(args)

    assert output_hashes(root) == before
    assert store_for(root).current() == previous
    assert not list((root / "Daily Reports").glob("*.xlsx"))


def test_explicit_old_complete_week_regenerates_missing_reports_with_no_new_input(
    published_copy, tmp_path, monkeypatch,
):
    root = published_copy
    old_start, old_end = date(2026, 9, 29), date(2026, 10, 4)
    previous = store_for(root).current()
    import_dir = tmp_path / "Earlier source exports"
    import_dir.mkdir()
    for path in write_week(root, start=old_start):
        shutil.move(path, import_dir / path.name)
    import_args = arguments(root, migrate_history_only=True, migrate_history_from=[str(import_dir)])
    assert workflow.run(import_args) == []
    assert import_args._run_attempt_recorder.outcome == "history_imported"
    imported = store_for(root).current()
    assert imported["report_publication_id"] == previous["report_publication_id"]
    assert imported["period"] == previous["period"]
    latest_before = output_hashes(root)
    assert "Check_Wine_RVA100426.xlsx" not in latest_before

    # This boundary concerns period selection, not the already-covered master
    # generator. Preserve its real management tables while recording the call.
    current_master = root / "Finished Reports/Red_Onion_Server_Master.xlsx"
    master_periods = []
    def master_copy(records, output_path, config, source_dir, public_start, public_end):
        master_periods.append((public_start, public_end))
        shutil.copy2(current_master, output_path)
        return output_path
    monkeypatch.setattr(weekly, "write_master_workbook", master_copy)

    args = arguments(root, week_start=old_start.isoformat(), week_end=old_end.isoformat())
    outputs = workflow.run(args)

    assert args._run_attempt_recorder.outcome == "published"
    assert {p.name for p in outputs} == {
        "Check_Wine_RVA100426.xlsx", "Check_Wine_VB100426.xlsx", "Red_Onion_Server_Master.xlsx",
    }
    for code in ("RVA", "VB"):
        assert public_totals(root, code, week_end=old_end) == pytest.approx((9000, 900, 10, 900, .10))
        name = f"Check_Wine_{code}101126.xlsx"
        assert output_hashes(root)[name] == latest_before[name]
    assert master_periods == [(date(2026, 10, 6), date(2026, 10, 11))]
    assert store_for(root).current()["changed_weeks"] == ["2026-10-04"]
    assert not list((root / "Daily Reports").glob("*.xlsx"))


@pytest.mark.parametrize("source_days", [1, 6])
def test_import_preserves_publication_then_next_ordinary_run_refreshes_accepted_history(
    published_copy, tmp_path, monkeypatch, source_days,
):
    root = published_copy
    previous = store_for(root).current()
    old_start, old_end = date(2026, 9, 29), date(2026, 10, 4)
    import_dir = tmp_path / "Earlier accepted history"
    import_dir.mkdir()
    for index, path in enumerate(write_week(root, start=old_start)):
        if index < source_days:
            shutil.move(path, import_dir / path.name)
        else:
            path.unlink()
    before = output_hashes(root)
    import_args = arguments(root, migrate_history_only=True, migrate_history_from=[str(import_dir)])
    assert workflow.run(import_args) == []
    imported = store_for(root).current()
    assert imported["report_publication_id"] == previous["report_publication_id"]
    assert imported["period"] == previous["period"]
    assert output_hashes(root) == before
    assert len(imported["days"]) == 6 + source_days

    current_master = root / "Finished Reports/Red_Onion_Server_Master.xlsx"
    generated_dates = []
    def master_copy(records, output_path, config, source_dir, public_start, public_end):
        generated_dates.extend({r.report_date for r in records})
        shutil.copy2(current_master, output_path)
        return output_path
    monkeypatch.setattr(weekly, "write_master_workbook", master_copy)

    args = arguments(root)
    outputs = workflow.run(args)
    assert args._run_attempt_recorder.outcome == "published"
    assert "Red_Onion_Server_Master.xlsx" in {p.name for p in outputs}
    assert old_start in generated_dates
    final = store_for(root).current()
    assert final["report_publication_id"] != previous["report_publication_id"]
    assert final["period"] == previous["period"]
    assert not final.get("pending_report_refresh")
    assert not final.get("pending_report_weeks")
    for code in ("RVA", "VB"):
        old_public = root / "Finished Reports" / f"Check_Wine_{code}{old_end:%m%d%y}.xlsx"
        assert old_public.exists() is (source_days == 6)
        if source_days == 6:
            assert public_totals(root, code, week_end=old_end) == pytest.approx((9000, 900, 10, 900, .10))
    assert not list((root / "Daily Reports").glob("*.xlsx"))


def test_first_import_can_publish_on_the_next_ordinary_run_without_a_prior_master(
    published_workspace, tmp_path, monkeypatch,
):
    root = make_workspace(tmp_path / "Fresh imported workspace")
    import_dir = tmp_path / "Initial source exports"
    import_dir.mkdir()
    for path in write_week(root):
        shutil.move(path, import_dir / path.name)
    args = arguments(root, migrate_history_only=True, migrate_history_from=[str(import_dir)])
    assert workflow.run(args) == []
    imported = store_for(root).current()
    assert imported["report_publication_id"] is None
    assert imported["period"] is None
    assert not imported["published_outputs"]

    template = published_workspace / "Finished Reports/Red_Onion_Server_Master.xlsx"
    def master_copy(records, output_path, config, source_dir, public_start, public_end):
        shutil.copy2(template, output_path)
        return output_path
    monkeypatch.setattr(weekly, "write_master_workbook", master_copy)
    publish_args = arguments(root)
    outputs = workflow.run(publish_args)
    assert publish_args._run_attempt_recorder.outcome == "published"
    assert len(outputs) == 3
    assert store_for(root).current()["report_publication_id"] is not None
    assert not store_for(root).current().get("pending_report_refresh")
