from __future__ import annotations

from argparse import Namespace
from datetime import date, timedelta
from pathlib import Path

from openpyxl import load_workbook
import pytest

import red_onion_weekly_metrics as metrics


def make_record(
    day: date,
    location: str = "RC Richmond",
    raw_user_name: str = "Server One",
    display_name: str | None = None,
    is_location_total: bool = False,
    gross_sales: float = 100.0,
    guest_count: float = 10.0,
    wine_sales: float = 20.0,
    rate: float = 1.0,
    ticket_seconds: float = 120.0,
    source_file: str | None = None,
) -> metrics.MetricRecord:
    if is_location_total and raw_user_name == "Server One":
        raw_user_name = ""
    if display_name is None:
        display_name = raw_user_name
    return metrics.MetricRecord(
        source_file=source_file or f"Daily Report {day.isoformat()}.xls",
        report_date=day,
        location=location,
        raw_user_name=raw_user_name,
        display_name=display_name,
        is_location_total=is_location_total,
        gross_sales=gross_sales,
        guest_count=guest_count,
        check_average=gross_sales / guest_count if guest_count else 0.0,
        wine_sales=wine_sales,
        wine_pct=wine_sales / gross_sales if gross_sales else 0.0,
        rate_of_sale_by_guest_count=rate,
        average_ticket_time_seconds=ticket_seconds,
    )


def weekly_row(
    week_end: date,
    *,
    location: str = "RC Richmond",
    display_name: str = "Server One",
    gross_sales: float = 1000.0,
    guest_count: float = 50.0,
    wine_sales: float = 100.0,
    rate: float = 2.0,
    ticket_seconds: float = 1200.0,
    active_days: int = 3,
) -> dict:
    return {
        "week_start": week_end - timedelta(days=metrics.OPERATING_WEEK_DAYS - 1),
        "week_end": week_end,
        "location": location,
        "raw_user_name": display_name,
        "display_name": display_name,
        "gross_sales": gross_sales,
        "guest_count": guest_count,
        "check_average": gross_sales / guest_count if guest_count else 0.0,
        "wine_sales": wine_sales,
        "wine_pct": wine_sales / gross_sales if gross_sales else 0.0,
        "rate_of_sale_by_guest_count": rate,
        "average_ticket_time_seconds": ticket_seconds,
        "active_days": active_days,
        "source_days": active_days,
        "source_files": "synthetic.xls",
    }


def ranked_row(week_end: date, *, ranks: tuple[int, int, int, int], **kwargs) -> dict:
    row = weekly_row(week_end, **kwargs)
    row.update(
        {
            "check_average_rank": ranks[0],
            "wine_pct_rank": ranks[1],
            "rate_rank": ranks[2],
            "ticket_time_rank": ranks[3],
        }
    )
    return row


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


def test_star_scoring_uses_all_metric_families_and_rank_movement(tmp_path: Path) -> None:
    config = metrics.load_config(tmp_path / "missing-config.json")
    prior_week = date(2026, 6, 7)
    latest_week = date(2026, 6, 14)
    ranked_rows = [
        ranked_row(
            prior_week,
            display_name="Alex Rising",
            gross_sales=1000,
            guest_count=50,
            wine_sales=80,
            rate=2.0,
            ticket_seconds=25 * 60,
            ranks=(5, 5, 5, 5),
        ),
        ranked_row(
            latest_week,
            display_name="Alex Rising",
            gross_sales=1400,
            guest_count=50,
            wine_sales=168,
            rate=1.5,
            ticket_seconds=18 * 60,
            ranks=(1, 1, 1, 1),
        ),
        ranked_row(
            prior_week,
            display_name="Jordan Falling",
            gross_sales=1500,
            guest_count=50,
            wine_sales=180,
            rate=1.5,
            ticket_seconds=18 * 60,
            ranks=(1, 1, 1, 1),
        ),
        ranked_row(
            latest_week,
            display_name="Jordan Falling",
            gross_sales=1000,
            guest_count=50,
            wine_sales=50,
            rate=2.0,
            ticket_seconds=25 * 60,
            ranks=(6, 6, 6, 6),
        ),
    ]

    stars = metrics.server_star_rows(metrics.server_week_trend_rows(ranked_rows), config)
    by_name = {row["display_name"]: row for row in stars}

    rising = by_name["Alex Rising"]
    assert rising["category"] == "Rising Star"
    assert rising["composite_score"] == 9
    assert rising["average_rank_movement"] == 4
    assert "Check avg +$8.00" in rising["why"]
    assert "Wine +4.0 pts" in rising["why"]
    assert "Rate -0.50" in rising["why"]
    assert "Ticket -7.0 min" in rising["why"]

    falling = by_name["Jordan Falling"]
    assert falling["category"] == "Falling Star"
    assert falling["composite_score"] == -10
    assert falling["average_rank_movement"] == -5


def test_star_rows_exclude_low_volume_admin_takeout_and_placeholders(tmp_path: Path) -> None:
    config = metrics.load_config(tmp_path / "missing-config.json")
    prior_week = date(2026, 6, 7)
    latest_week = date(2026, 6, 14)
    ranked_rows: list[dict] = []
    for name, guest_count, active_days in (
        ("AGM Manager", 50, 3),
        ("Takeout", 50, 3),
        ("1 Server Server", 50, 3),
        ("Low Signal Person", 5, 1),
        ("Real Person", 50, 3),
    ):
        ranked_rows.extend(
            [
                ranked_row(
                    prior_week,
                    display_name=name,
                    gross_sales=1000,
                    guest_count=guest_count,
                    wine_sales=100,
                    active_days=active_days,
                    ranks=(5, 5, 5, 5),
                ),
                ranked_row(
                    latest_week,
                    display_name=name,
                    gross_sales=1500,
                    guest_count=guest_count,
                    wine_sales=180,
                    rate=1.0,
                    ticket_seconds=10 * 60,
                    active_days=active_days,
                    ranks=(1, 1, 1, 1),
                ),
            ]
        )

    stars = metrics.server_star_rows(metrics.server_week_trend_rows(ranked_rows), config)

    assert [row["display_name"] for row in stars] == ["Real Person"]


def test_store_week_trend_deltas() -> None:
    prior_week = weekly_row(
        date(2026, 6, 7),
        gross_sales=1000,
        guest_count=50,
        wine_sales=100,
        rate=2.0,
        ticket_seconds=25 * 60,
        active_days=6,
    )
    latest_week = weekly_row(
        date(2026, 6, 14),
        gross_sales=1500,
        guest_count=60,
        wine_sales=180,
        rate=1.5,
        ticket_seconds=20 * 60,
        active_days=6,
    )

    trend_rows = metrics.weekly_metric_trend_rows([prior_week, latest_week], ("location",))
    latest = trend_rows[-1]

    assert latest["gross_sales_change"] == 500
    assert latest["guest_count_change"] == 10
    assert latest["check_average_change"] == 5
    assert latest["wine_pct_change"] == pytest.approx(0.02)
    assert latest["rate_change"] == -0.5
    assert latest["ticket_time_change_minutes"] == -5
    assert latest["trend_note"] == "Improving"


def test_all_stores_group_aggregation_and_trend_deltas() -> None:
    records = [
        make_record(
            date(2026, 6, 7),
            "RC Richmond",
            is_location_total=True,
            gross_sales=1000,
            guest_count=50,
            wine_sales=100,
            rate=2,
            ticket_seconds=600,
            source_file="richmond-prior.xls",
        ),
        make_record(
            date(2026, 6, 7),
            "RC Virginia Beach",
            is_location_total=True,
            gross_sales=500,
            guest_count=25,
            wine_sales=50,
            rate=4,
            ticket_seconds=1200,
            source_file="vb-prior.xls",
        ),
        make_record(
            date(2026, 6, 14),
            "RC Richmond",
            is_location_total=True,
            gross_sales=1200,
            guest_count=60,
            wine_sales=180,
            rate=1.5,
            ticket_seconds=600,
            source_file="richmond-latest.xls",
        ),
        make_record(
            date(2026, 6, 14),
            "RC Virginia Beach",
            is_location_total=True,
            gross_sales=900,
            guest_count=30,
            wine_sales=90,
            rate=3,
            ticket_seconds=900,
            source_file="vb-latest.xls",
        ),
    ]

    group_rows = metrics.group_weekly_rows(records)
    prior = group_rows[0]
    assert prior["group"] == "All Stores"
    assert prior["gross_sales"] == 1500
    assert prior["guest_count"] == 75
    assert prior["check_average"] == 20
    assert prior["wine_pct"] == pytest.approx(0.1)
    assert prior["rate_of_sale_by_guest_count"] == pytest.approx((2 * 50 + 4 * 25) / 75)
    assert prior["average_ticket_time_seconds"] == pytest.approx((600 * 50 + 1200 * 25) / 75)
    assert prior["active_days"] == 1

    trend_rows = metrics.weekly_metric_trend_rows(group_rows, ("group",))
    latest = trend_rows[-1]
    assert latest["gross_sales_change"] == 600
    assert latest["guest_count_change"] == 15
    assert latest["check_average_change"] == pytest.approx((2100 / 90) - 20)
    assert latest["wine_pct_change"] == pytest.approx((270 / 2100) - 0.1)


def test_master_workbook_contains_star_store_group_and_dashboard_sections(tmp_path: Path) -> None:
    config = metrics.load_config(tmp_path / "missing-config.json")
    records: list[metrics.MetricRecord] = []
    for report_date, gross, wine, rate, ticket_seconds in (
        (date(2026, 6, 7), 1000, 100, 2.0, 25 * 60),
        (date(2026, 6, 14), 1500, 180, 1.5, 20 * 60),
    ):
        for location in ("RC Richmond", "RC Virginia Beach"):
            records.append(
                make_record(
                    report_date,
                    location,
                    is_location_total=True,
                    gross_sales=gross * 2,
                    guest_count=100,
                    wine_sales=wine * 2,
                    rate=rate,
                    ticket_seconds=ticket_seconds,
                    source_file=f"{location}-{report_date}.xls",
                )
            )
            records.append(
                make_record(
                    report_date,
                    location,
                    raw_user_name="Server One",
                    gross_sales=gross,
                    guest_count=50,
                    wine_sales=wine,
                    rate=rate,
                    ticket_seconds=ticket_seconds,
                    source_file=f"{location}-{report_date}.xls",
                )
            )

    output_path = tmp_path / "Red_Onion_Server_Master.xlsx"
    metrics.write_master_workbook(
        records,
        output_path,
        config,
        tmp_path / "Daily Reports",
        date(2026, 6, 9),
        date(2026, 6, 14),
    )

    wb = load_workbook(output_path)
    assert "Rising & Falling Stars" in wb.sheetnames
    assert "Store Week Trends" in wb.sheetnames
    assert "Store Trend Summary" in wb.sheetnames
    assert "Group Week Trends" in wb.sheetnames
    assert "Group Trend Summary" in wb.sheetnames
    assert wb["_Dashboard Chart Data"].sheet_state == "hidden"

    dashboard_values = {
        value
        for row in wb["Dashboard"].iter_rows(values_only=True)
        for value in row
        if isinstance(value, str)
    }
    assert "Rising Stars" in dashboard_values
    assert "Falling Stars" in dashboard_values
    assert "Latest Store Trend Summary" in dashboard_values
    assert "All-Stores Group Trend" in dashboard_values
    assert "Data Quality Warnings" in dashboard_values
    assert "Short Week" in dashboard_values
