from __future__ import annotations

from argparse import Namespace
from datetime import date
from pathlib import Path

import pytest

import red_onion_weekly_metrics as metrics


def make_record(day: date, location: str = "RC Richmond") -> metrics.MetricRecord:
    return metrics.MetricRecord(
        source_file=f"Daily Report {day.isoformat()}.xls",
        report_date=day,
        location=location,
        raw_user_name="Server One",
        display_name="Server One",
        is_location_total=False,
        gross_sales=100.0,
        guest_count=10.0,
        check_average=10.0,
        wine_sales=20.0,
        wine_pct=0.2,
        rate_of_sale_by_guest_count=1.0,
        average_ticket_time_seconds=120.0,
    )


def args_for(tmp_path: Path) -> Namespace:
    return Namespace(
        input_dir=str(tmp_path / "Daily Reports"),
        output_dir=str(tmp_path / "Output"),
        archive_dir=str(tmp_path / "Archive - Old Files"),
        config=str(tmp_path / "missing-config.json"),
        week_start=None,
        week_end=None,
    )


def test_zero_active_files_stop_with_clear_message(tmp_path: Path) -> None:
    args = args_for(tmp_path)
    Path(args.input_dir).mkdir(parents=True)

    with pytest.raises(FileNotFoundError, match="No active daily .xls reports"):
        metrics.run(args)


def test_mixed_active_weeks_stop_with_exact_file_names(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    args = args_for(tmp_path)
    input_dir = Path(args.input_dir)
    input_dir.mkdir(parents=True)
    first = input_dir / "Daily Report - TM - 06-07-2026.xls"
    second = input_dir / "Daily Report - TM - 06-14-2026.xls"
    first.write_text("first", encoding="utf-8")
    second.write_text("second", encoding="utf-8")

    def fake_parse(path: Path, config: dict) -> list[metrics.MetricRecord]:
        if path == first:
            return [make_record(date(2026, 6, 7))]
        return [make_record(date(2026, 6, 14))]

    monkeypatch.setattr(metrics, "parse_daily_report", fake_parse)

    with pytest.raises(ValueError) as exc_info:
        metrics.run(args)

    message = str(exc_info.value)
    assert "more than one Tuesday-Sunday operating week" in message
    assert first.name in message
    assert second.name in message
    assert first.exists()
    assert second.exists()


def test_successful_run_archives_active_files_and_keeps_master_history(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    args = args_for(tmp_path)
    input_dir = Path(args.input_dir)
    archive_dir = Path(args.archive_dir)
    output_dir = Path(args.output_dir)
    input_dir.mkdir(parents=True)
    archived_source = (
        archive_dir
        / "processed-daily-reports"
        / "week-ending-2026-06-07"
        / "Daily Report archived.xls"
    )
    archived_source.parent.mkdir(parents=True)
    active_source = input_dir / "Daily Report active.xls"
    archived_source.write_text("archived", encoding="utf-8")
    active_source.write_text("active", encoding="utf-8")
    master_records: list[metrics.MetricRecord] = []

    def fake_parse(path: Path, config: dict) -> list[metrics.MetricRecord]:
        if path == archived_source:
            return [make_record(date(2026, 6, 7), "RC Virginia Beach")]
        return [make_record(date(2026, 6, 14), "RC Richmond")]

    def fake_public(location, records, output_path, config, start, end):
        output_file = output_path / f"public-{location}.xlsx"
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text("public", encoding="utf-8")
        return output_file

    def fake_master(records, output_path, config, input_path, start, end):
        master_records.extend(records)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text("master", encoding="utf-8")
        return output_path

    monkeypatch.setattr(metrics, "parse_daily_report", fake_parse)
    monkeypatch.setattr(metrics, "write_public_workbook", fake_public)
    monkeypatch.setattr(metrics, "write_master_workbook", fake_master)

    generated = metrics.run(args)

    assert not active_source.exists()
    assert (
        archive_dir
        / "processed-daily-reports"
        / "week-ending-2026-06-14"
        / active_source.name
    ).exists()
    assert {record.report_date for record in master_records} == {
        date(2026, 6, 7),
        date(2026, 6, 14),
    }
    assert output_dir / "Red_Onion_Server_Master.xlsx" in generated


def test_failed_run_leaves_active_files_in_place(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    args = args_for(tmp_path)
    input_dir = Path(args.input_dir)
    input_dir.mkdir(parents=True)
    active_source = input_dir / "Daily Report active.xls"
    active_source.write_text("active", encoding="utf-8")

    def fake_parse(path: Path, config: dict) -> list[metrics.MetricRecord]:
        return [make_record(date(2026, 6, 14))]

    def fake_public(location, records, output_path, config, start, end):
        output_path.mkdir(parents=True, exist_ok=True)
        output_file = output_path / f"public-{location}.xlsx"
        output_file.write_text("public", encoding="utf-8")
        return output_file

    def failing_master(records, output_path, config, input_path, start, end):
        raise RuntimeError("workbook failed")

    monkeypatch.setattr(metrics, "parse_daily_report", fake_parse)
    monkeypatch.setattr(metrics, "write_public_workbook", fake_public)
    monkeypatch.setattr(metrics, "write_master_workbook", failing_master)

    with pytest.raises(RuntimeError, match="workbook failed"):
        metrics.run(args)

    assert active_source.exists()


def test_duplicate_archive_handling_preserves_different_same_name_files(tmp_path: Path) -> None:
    source_dir = tmp_path / "Daily Reports"
    archive_root = tmp_path / "Archive - Old Files"
    source_dir.mkdir()
    source = source_dir / "Daily Report active.xls"
    source.write_text("new content", encoding="utf-8")
    existing = (
        archive_root
        / "processed-daily-reports"
        / "week-ending-2026-06-14"
        / source.name
    )
    existing.parent.mkdir(parents=True)
    existing.write_text("old content", encoding="utf-8")

    archived = metrics.archive_processed_files([source], archive_root, date(2026, 6, 14))

    assert not source.exists()
    assert existing.read_text(encoding="utf-8") == "old content"
    assert archived == [existing.with_name("Daily Report active (1).xls")]
    assert archived[0].read_text(encoding="utf-8") == "new content"


def test_duplicate_archive_handling_removes_identical_active_copy(tmp_path: Path) -> None:
    source_dir = tmp_path / "Daily Reports"
    archive_root = tmp_path / "Archive - Old Files"
    source_dir.mkdir()
    source = source_dir / "Daily Report active.xls"
    source.write_text("same content", encoding="utf-8")
    existing = (
        archive_root
        / "processed-daily-reports"
        / "week-ending-2026-06-14"
        / source.name
    )
    existing.parent.mkdir(parents=True)
    existing.write_text("same content", encoding="utf-8")

    archived = metrics.archive_processed_files([source], archive_root, date(2026, 6, 14))

    assert not source.exists()
    assert archived == [existing]
    assert existing.read_text(encoding="utf-8") == "same content"


def test_parser_defaults_point_to_operator_root_folders() -> None:
    args = metrics.build_parser().parse_args([])

    assert Path(args.input_dir).name == "Daily Reports"
    assert Path(args.output_dir).name == "Output"
    assert Path(args.archive_dir).name == "Archive - Old Files"
    assert Path(args.input_dir).parent == metrics.PROJECT_ROOT
    assert Path(args.config).parent == metrics.PROGRAM_DIR
