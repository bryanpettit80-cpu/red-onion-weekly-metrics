from __future__ import annotations

from datetime import date, time
from pathlib import Path

from openpyxl import Workbook
import pytest

import red_onion_weekly_metrics as metrics


def make_record(path: Path, day: date, *, gross_sales: float = 100.0) -> metrics.MetricRecord:
    return metrics.MetricRecord(
        source_file=path.name,
        report_date=day,
        location="RC Richmond",
        raw_user_name="Server One",
        display_name="Server One",
        is_location_total=False,
        gross_sales=gross_sales,
        guest_count=10.0,
        check_average=gross_sales / 10.0,
        wine_sales=20.0,
        wine_pct=20.0 / gross_sales,
        rate_of_sale_by_guest_count=0.2,
        average_ticket_time_seconds=1200.0,
    )


def test_daily_report_discovery_accepts_xls_and_xlsx_and_ignores_temporary_files(
    tmp_path: Path,
) -> None:
    expected_names = {
        "Daily Report first.xls",
        "Daily Report second.xlsx",
        "daily report third.XLSX",
    }
    for name in [
        *expected_names,
        "~$Daily Report temporary.xlsx",
        "Daily Report notes.csv",
        "Other workbook.xlsx",
    ]:
        (tmp_path / name).write_text("placeholder", encoding="utf-8")

    assert {path.name for path in metrics.daily_report_paths(tmp_path)} == expected_names


def test_valid_xlsx_uses_internal_business_date(tmp_path: Path) -> None:
    path = tmp_path / "Daily Report - TM - 07-12-2026.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Report(All)"
    worksheet.append(["Date(s): 07/11/2026"])
    worksheet.append(
        [
            "Store",
            "User",
            "Gross Sales",
            "Guest Count",
            "Check Average",
            "Wine Sales",
            "Rate of Sale by Guest Count",
            "Average Ticket Time",
        ]
    )
    worksheet.append(
        ["RC Richmond", "Server One", 100.0, 10.0, 10.0, 20.0, 0.2, time(0, 20)]
    )
    workbook.save(path)
    workbook.close()

    records = metrics.parse_daily_report(path, metrics.load_config(tmp_path / "missing.json"))

    assert metrics.daily_report_excel_engine(path) == "openpyxl"
    assert {record.report_date for record in records} == {date(2026, 7, 11)}
    assert records[0].gross_sales == 100.0


def test_archived_history_is_limited_to_canonical_processed_folder(tmp_path: Path) -> None:
    canonical = (
        tmp_path
        / metrics.CANONICAL_DAILY_ARCHIVE_FOLDER
        / "week-ending-2026-06-14"
        / "Daily Report canonical.xlsx"
    )
    backup = tmp_path / "pre-redesign" / "Daily Report backup.xls"
    canonical.parent.mkdir(parents=True)
    backup.parent.mkdir(parents=True)
    canonical.write_text("canonical", encoding="utf-8")
    backup.write_text("backup", encoding="utf-8")

    assert metrics.archived_daily_report_paths(tmp_path) == [canonical]


def test_copy_only_migration_is_semantic_and_idempotent(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    first_source = tmp_path / "legacy-one"
    second_source = tmp_path / "legacy-two"
    archive_root = tmp_path / "archive"
    first_source.mkdir()
    second_source.mkdir()
    xls_path = first_source / "Daily Report same-day.xls"
    xlsx_path = second_source / "Daily Report same-day-converted.xlsx"
    xls_path.write_text("legacy binary", encoding="utf-8")
    xlsx_path.write_text("converted binary with different bytes", encoding="utf-8")

    def fake_parse(path: Path, config: dict) -> list[metrics.MetricRecord]:
        return [make_record(path, date(2026, 6, 11))]

    monkeypatch.setattr(metrics, "parse_daily_report", fake_parse)

    first = metrics.migrate_history_files(
        [first_source, second_source], archive_root, metrics.DEFAULT_CONFIG
    )
    second = metrics.migrate_history_files(
        [first_source, second_source], archive_root, metrics.DEFAULT_CONFIG
    )

    expected_folder = (
        archive_root
        / metrics.CANONICAL_DAILY_ARCHIVE_FOLDER
        / "week-ending-2026-06-14"
    )
    assert first.copied_paths == (expected_folder / xlsx_path.name,)
    assert first.duplicate_files_ignored == 1
    assert first.business_dates_considered == 1
    assert second.copied_paths == ()
    assert xls_path.exists()
    assert xlsx_path.exists()
    assert len(metrics.archived_daily_report_paths(archive_root)) == 1


def test_migration_conflict_blocks_before_copying(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source_dir = tmp_path / "legacy"
    archive_root = tmp_path / "archive"
    source_dir.mkdir()
    first = source_dir / "Daily Report first.xls"
    second = source_dir / "Daily Report second.xlsx"
    first.write_text("first", encoding="utf-8")
    second.write_text("second", encoding="utf-8")

    def fake_parse(path: Path, config: dict) -> list[metrics.MetricRecord]:
        gross_sales = 100.0 if path == first else 125.0
        return [make_record(path, date(2026, 6, 11), gross_sales=gross_sales)]

    monkeypatch.setattr(metrics, "parse_daily_report", fake_parse)

    with pytest.raises(ValueError, match="Conflicting daily reports") as exc_info:
        metrics.migrate_history_files([source_dir], archive_root, metrics.DEFAULT_CONFIG)

    assert "2026-06-11" in str(exc_info.value)
    assert first.name in str(exc_info.value)
    assert second.name in str(exc_info.value)
    assert not metrics.canonical_daily_archive_dir(archive_root).exists()


def test_migration_reserves_distinct_destinations_for_same_named_reports(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    first_source = tmp_path / "legacy-one"
    second_source = tmp_path / "legacy-two"
    archive_root = tmp_path / "archive"
    first_source.mkdir()
    second_source.mkdir()
    first = first_source / "Daily Report same-name.xlsx"
    second = second_source / "Daily Report same-name.xlsx"
    first.write_text("2026-06-10", encoding="utf-8")
    second.write_text("2026-06-11", encoding="utf-8")

    def fake_parse(path: Path, config: dict) -> list[metrics.MetricRecord]:
        return [make_record(path, date.fromisoformat(path.read_text(encoding="utf-8")))]

    monkeypatch.setattr(metrics, "parse_daily_report", fake_parse)

    first_result = metrics.migrate_history_files(
        [first_source, second_source], archive_root, metrics.DEFAULT_CONFIG
    )
    second_result = metrics.migrate_history_files(
        [first_source, second_source], archive_root, metrics.DEFAULT_CONFIG
    )

    expected_folder = (
        archive_root
        / metrics.CANONICAL_DAILY_ARCHIVE_FOLDER
        / "week-ending-2026-06-14"
    )
    assert first_result.copied_paths == (
        expected_folder / first.name,
        expected_folder / "Daily Report same-name (1).xlsx",
    )
    assert first_result.business_dates_considered == 2
    assert second_result.copied_paths == ()
    assert len(metrics.archived_daily_report_paths(archive_root)) == 2


def test_migration_rechecks_all_destinations_before_copying(tmp_path: Path) -> None:
    first_source = tmp_path / "Daily Report first.xlsx"
    second_source = tmp_path / "Daily Report second.xlsx"
    first_destination = tmp_path / "archive" / first_source.name
    second_destination = tmp_path / "archive" / second_source.name
    first_source.write_text("first", encoding="utf-8")
    second_source.write_text("second", encoding="utf-8")
    second_destination.parent.mkdir(parents=True)
    second_destination.write_text("appeared after planning", encoding="utf-8")
    plan = metrics.HistoryMigrationPlan(
        copy_pairs=(
            (first_source, first_destination),
            (second_source, second_destination),
        ),
        effective_records_by_path={},
        duplicate_paths=(),
        business_dates=(),
    )

    with pytest.raises(RuntimeError, match="No files were copied"):
        metrics.apply_history_migration_plan(plan)

    assert not first_destination.exists()
    assert second_destination.read_text(encoding="utf-8") == "appeared after planning"


def test_history_cli_accepts_repeatable_sources_and_migration_only() -> None:
    args = metrics.build_parser().parse_args(
        [
            "--migrate-history-from",
            "first-folder",
            "--migrate-history-from",
            "second-folder",
            "--migrate-history-only",
        ]
    )

    assert args.migrate_history_from == ["first-folder", "second-folder"]
    assert args.migrate_history_only is True
