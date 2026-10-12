from __future__ import annotations

from dataclasses import replace
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
import pytest

from red_onion_config import ConfigError, validate_config_payload
from red_onion_ingestion import METRIC_HEADERS, parse_single_business_date
import red_onion_weekly_metrics as metrics


START = date(2026, 7, 21)
END = date(2026, 7, 26)


def write_source(path: Path, headers: list[str], *, metadata: str = "Date(s): 07/21/2026") -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report(All)"
    sheet.append([metadata])
    sheet.append(headers)
    values = {
        "Store": "RC Richmond", "User": "Alex", "": "Alex",
        "Gross Sales": 1000, "Guest Count": 100, "Check Average": 10,
        "Wine Sales": 100, "Rate of Sale by Guest Count": 5,
        "Average Ticket Time": "00:10:00", "Check Count": 20,
    }
    sheet.append([values.get(header) for header in headers])
    workbook.save(path)
    workbook.close()
    return path


def example_record(day: date = START, *, location: str = "RC Richmond", total: bool = False) -> metrics.MetricRecord:
    return metrics.MetricRecord(
        source_file="synthetic.xlsx", report_date=day, location=location,
        raw_user_name="" if total else "Alex", display_name="" if total else "Alex",
        is_location_total=total, gross_sales=1000, guest_count=100,
        check_average=10, wine_sales=100, wine_pct=0.1,
        rate_of_sale_by_guest_count=5, average_ticket_time_seconds=100,
        rate_available=True, ticket_time_available=True,
        check_count=10, check_count_available=True,
    )


def complete_week() -> list[metrics.MetricRecord]:
    return [
        example_record(START + timedelta(days=offset), location=location, total=total)
        for location in metrics.DEFAULT_CONFIG["locations"]
        for offset in range(6)
        for total in (False, True)
    ]


@pytest.mark.parametrize("field", [
    "Date(s): 07/21/2026", "Date(s): 7/21/26 - 07/21/2026",
    "Date(s): 07/21/2026 – 07/21/2026",
])
def test_single_explicit_day_and_equivalent_ranges(field: str) -> None:
    assert parse_single_business_date([field], "renamed.xlsx") == START


@pytest.mark.parametrize("fields", [
    ["Date(s): 07/21/2026 - 07/26/2026"],
    ["Date(s): 07/21/2026", "Date(s): 07/22/2026"],
    ["Date(s): 02/30/2026"], ["Date(s):"], ["No date metadata"],
])
def test_ambiguous_or_missing_dates_cannot_use_filename(fields: list[str]) -> None:
    with pytest.raises(ValueError, match="business date|Date\\(s\\)|filename"):
        parse_single_business_date(fields, "Daily Report 07-22-2026.xlsx")


def test_reordered_metric_and_identity_headers_preserve_business_values(tmp_path: Path) -> None:
    path = write_source(tmp_path / "Renamed export.xlsx", [
        "User", "Store", "Check Count", "Gross Sales", "Check Average", "Wine Sales",
        "Guest Count", "Average Ticket Time", "Rate of Sale by Guest Count",
    ])
    row = metrics.parse_daily_report(path, metrics.DEFAULT_CONFIG)[0]
    assert (row.raw_user_name, row.gross_sales, row.guest_count, row.check_average) == ("Alex", 1000, 100, 10)
    assert (row.wine_pct, row.check_count, row.average_ticket_time_seconds) == (0.1, 20, 600)


def test_current_two_location_type_blocks_use_first_populated_block(tmp_path: Path) -> None:
    path = tmp_path / "Two types.xlsx"
    headers = ["Store", "", *METRIC_HEADERS, "Check Count", *METRIC_HEADERS, "Check Count"]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report(All)"
    sheet.append(["Date(s): 07/21/2026"])
    sheet.append(headers)
    sheet.append(["RC Richmond", "Alex", *([None] * 7), 1000, 100, 10, 100, 5, "00:10:00", 20])
    workbook.save(path)
    workbook.close()
    row = metrics.parse_daily_report(path, metrics.DEFAULT_CONFIG)[0]
    assert (row.gross_sales, row.guest_count, row.check_count) == (1000, 100, 20)


@pytest.mark.parametrize("headers", [
    ["Store", "User", "Gross Sales", "Guest Count", "Wine Sales"],
    ["Store", "User", *METRIC_HEADERS, "Guest Count"],
    ["Store", "User", "Server", *METRIC_HEADERS],
])
def test_missing_or_ambiguous_columns_fail_before_metrics(tmp_path: Path, headers: list[str]) -> None:
    path = write_source(tmp_path / "Invalid structure.xlsx", headers)
    with pytest.raises(ValueError, match="headers|User column"):
        metrics.parse_daily_report(path, metrics.DEFAULT_CONFIG)


def test_rename_resave_and_formatting_do_not_change_normalized_facts(tmp_path: Path) -> None:
    original = write_source(tmp_path / "Daily Report.xlsx", ["Store", "User", *METRIC_HEADERS, "Check Count"])
    renamed = tmp_path / "Tuesday Richmond.xlsx"
    workbook = load_workbook(original)
    workbook.active.column_dimensions["A"].width = 35
    workbook.save(renamed)
    workbook.close()
    (tmp_path / "~$open.xlsx").write_bytes(b"Excel temporary lock")
    (tmp_path / "Read me.txt").write_text("operator note", encoding="utf-8")
    paths = metrics.daily_report_paths(tmp_path)
    assert paths == [original, renamed]
    reports = metrics.read_reports_by_path(paths, metrics.DEFAULT_CONFIG)
    resolution = metrics.resolve_report_duplicates(reports)
    assert len(resolution.records_by_path) == 1
    assert len(resolution.duplicate_paths) == 1


def test_unrecognized_excel_content_is_not_silently_skipped(tmp_path: Path) -> None:
    path = tmp_path / "Unexpected workbook.xlsx"
    workbook = Workbook()
    workbook.save(path)
    workbook.close()
    assert metrics.daily_report_paths(tmp_path) == [path]
    with pytest.raises(ValueError, match="Unexpected workbook.xlsx.*Report\\(All\\)"):
        metrics.read_reports_by_path([path], metrics.DEFAULT_CONFIG)


def test_richer_optional_context_cannot_hide_overlapping_check_correction() -> None:
    rich = [example_record(location=location, total=total) for location in metrics.DEFAULT_CONFIG["locations"] for total in (False, True)]
    partial = [replace(row, check_count=30 if row.location == "RC Richmond" else 0,
                       check_count_available=row.location == "RC Richmond") for row in rich]
    for records in (rich, partial):
        metrics.validate_daily_location_reconciliation(records, metrics.DEFAULT_CONFIG)
    with pytest.raises(ValueError, match="Conflicting daily reports"):
        metrics.resolve_report_duplicates({Path("rich.xlsx"): rich, Path("partial.xlsx"): partial})


def test_richer_compatible_context_is_accepted_once() -> None:
    rich = [example_record()]
    legacy = [replace(rich[0], check_count=0, check_count_available=False)]
    resolution = metrics.resolve_report_duplicates({Path("rich.xlsx"): rich, Path("legacy.xls"): legacy})
    assert list(resolution.records_by_path) == [Path("rich.xlsx")]


def test_complete_reporting_set_requires_twelve_actual_restaurant_days() -> None:
    records = complete_week()
    coverage = metrics.validate_complete_reporting_set(records, metrics.DEFAULT_CONFIG, START, END)
    assert coverage.ready and len(coverage.expected) == len(coverage.received) == 12
    missing = START + timedelta(days=2)
    records = [row for row in records if (row.location, row.report_date) != ("RC Richmond", missing)]
    with pytest.raises(ValueError, match="RC Richmond 2026-07-23"):
        metrics.validate_complete_reporting_set(records, metrics.DEFAULT_CONFIG, START, END)
    _, locations = metrics.weekly_rollups(records)
    readiness = metrics.latest_week_readiness(records, locations, metrics.DEFAULT_CONFIG)
    assert not readiness.ready
    assert readiness.missing_restaurant_days == (("RC Richmond", missing),)


def test_zero_total_is_receipt_without_creating_person_activity() -> None:
    records = complete_week()
    zero_day = START + timedelta(days=2)
    records = [replace(row, gross_sales=0, guest_count=0, wine_sales=0, check_count=0)
               if (row.location, row.report_date) == ("RC Richmond", zero_day) else row for row in records]
    servers, locations = metrics.weekly_rollups(records)
    richmond = next(row for row in locations if row["location"] == "RC Richmond")
    server = next(row for row in servers if row["location"] == "RC Richmond")
    assert (richmond["source_days"], richmond["active_days"], server["active_days"]) == (6, 5, 5)
    assert metrics.validate_complete_reporting_set(records, metrics.DEFAULT_CONFIG, START, END).ready


def test_explicit_closure_counts_as_coverage_without_fabricating_sales() -> None:
    closed_day = START + timedelta(days=2)
    config = validate_config_payload({"closed_dates": {"RC Richmond": [closed_day.isoformat()]}})
    records = [row for row in complete_week() if (row.location, row.report_date) != ("RC Richmond", closed_day)]
    coverage = metrics.validate_complete_reporting_set(records, config, START, END)
    assert (len(coverage.received), len(coverage.closed)) == (11, 1)
    _, locations = metrics.weekly_rollups(records, config)
    richmond = next(row for row in locations if row["location"] == "RC Richmond")
    assert (richmond["source_days"], richmond["active_days"], richmond["gross_sales"]) == (6, 5, 5000)
    assert metrics.latest_week_readiness(records, locations, config).ready
    with pytest.raises(ValueError, match="declared closed but source activity exists"):
        metrics.validate_complete_reporting_set(complete_week(), config, START, END)


def test_whole_week_closure_has_no_invented_financial_row(tmp_path: Path) -> None:
    config = validate_config_payload({"closed_dates": {"RC Richmond": [
        (START + timedelta(days=offset)).isoformat() for offset in range(6)
    ]}})
    records = [row for row in complete_week() if row.location == "RC Virginia Beach"]
    assert metrics.validate_complete_reporting_set(records, config, START, END).ready
    servers, locations = metrics.weekly_rollups(records, config)
    assert {row["location"] for row in locations} == {"RC Virginia Beach"}
    _, full = metrics.full_week_ends_by_location(
        locations, config["locations"], closed_dates=config["closed_dates"],
    )
    assert full == {END}
    model = metrics.build_performance_consistency_model(servers, locations, config)
    assert model["week_ends"] == (END,)
    path = metrics.write_public_workbook("RC Richmond", records, tmp_path, config, START, END)
    workbook = load_workbook(path)
    try:
        assert "No daily source; declared closed" in workbook["Report(All)"]["A1"].value
        assert workbook["Report(All)"]["C3"].value is None
    finally:
        workbook.close()


@pytest.mark.parametrize("closures", [
    {"Unknown": ["2026-07-21"]}, {"RC Richmond": ["2026-07-20"]},
    {"RC Richmond": ["07/21/2026"]}, {"RC Richmond": ["2026-07-21", "2026-07-21"]},
    {"RC Richmond": "2026-07-21"}, {"RC Richmond": ["2026-02-30"]},
])
def test_closure_rules_are_explicit_and_validated(closures: object) -> None:
    with pytest.raises(ConfigError):
        validate_config_payload({"closed_dates": closures})


def test_partial_public_preview_names_local_missing_day(tmp_path: Path) -> None:
    missing = START + timedelta(days=2)
    records = [row for row in complete_week() if (row.location, row.report_date) != ("RC Richmond", missing)]
    path = metrics.write_public_workbook("RC Richmond", records, tmp_path, metrics.DEFAULT_CONFIG, START, END)
    workbook = load_workbook(path)
    try:
        assert "PRELIMINARY - Missing: 2026-07-23" in workbook["Report(All)"]["A1"].value
        assert workbook["Report(All)"]["E2"].value == "Sales per Guest"
    finally:
        workbook.close()


def test_independent_business_known_answer_for_six_days_two_restaurants() -> None:
    a = example_record()
    b = replace(a, raw_user_name="Blair", display_name="Blair", gross_sales=500,
                guest_count=50, wine_sales=50, check_count=30,
                rate_of_sale_by_guest_count=10, average_ticket_time_seconds=300)
    daily = metrics.aggregate_records([a, b], ("location",))[0]
    expected_daily = {
        "gross_sales": 1500, "guest_count": 150, "check_count": 40, "wine_sales": 150,
        "check_average": 10, "wine_pct": 0.1, "sales_per_check": 37.5,
        "guests_per_check": 3.75, "rate_of_sale_by_guest_count": 6,
        "average_ticket_time_seconds": 250,
    }
    for key, expected in expected_daily.items():
        assert daily[key] == pytest.approx(expected)
    all_records = [replace(row, location=location, report_date=START + timedelta(days=offset))
                   for location in metrics.DEFAULT_CONFIG["locations"] for offset in range(6) for row in (a, b)]
    stores = metrics.aggregate_records(all_records, ("location",))
    group = metrics.aggregate_records(all_records, ())[0]
    for row in stores:
        assert (row["gross_sales"], row["guest_count"], row["check_count"], row["wine_sales"]) == (9000, 900, 240, 900)
        for key in ("check_average", "wine_pct", "sales_per_check", "guests_per_check", "rate_of_sale_by_guest_count", "average_ticket_time_seconds"):
            assert row[key] == pytest.approx(expected_daily[key])
    assert (group["gross_sales"], group["guest_count"], group["check_count"], group["wine_sales"]) == (18000, 1800, 480, 1800)
