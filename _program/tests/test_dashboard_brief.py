from __future__ import annotations

from datetime import date, timedelta

from openpyxl import Workbook

import red_onion_weekly_metrics as metrics


METRIC_VALUES = {
    "gross_sales": 12000.0,
    "guest_count": 500.0,
    "check_average": 24.0,
    "wine_pct": 0.12,
    "rate_of_sale_by_guest_count": 0.18,
    "average_ticket_time_seconds": 72 * 60.0,
}


def record_for(day: date) -> metrics.MetricRecord:
    return metrics.MetricRecord(
        source_file=f"daily-{day.isoformat()}.xlsx",
        report_date=day,
        location="RC Richmond",
        raw_user_name="",
        display_name="",
        is_location_total=True,
        gross_sales=100.0,
        guest_count=10.0,
        check_average=10.0,
        wine_sales=10.0,
        wine_pct=0.10,
        rate_of_sale_by_guest_count=0.20,
        average_ticket_time_seconds=3600.0,
    )


def weekly_location_row(
    week_end: date, location: str, source_days: int = metrics.OPERATING_WEEK_DAYS
) -> dict:
    return {
        "week_start": week_end - timedelta(days=metrics.OPERATING_WEEK_DAYS - 1),
        "week_end": week_end,
        "location": location,
        **METRIC_VALUES,
        "active_days": source_days,
        "source_days": source_days,
    }


def management_item(entity: str, *, group: bool = False) -> dict:
    benchmark = {**METRIC_VALUES, "gross_sales": 11000.0, "guest_count": 480.0}
    return {
        "entity": "All Stores" if group else entity,
        "latest": dict(METRIC_VALUES),
        "prior": dict(METRIC_VALUES),
        "benchmark_values": benchmark,
        "benchmark_sources": {field: "4-week baseline" for field in METRIC_VALUES},
        "prior_changes": {field: 0.0 for field in METRIC_VALUES},
        "benchmark_changes": {
            field: METRIC_VALUES[field] - benchmark[field] for field in METRIC_VALUES
        },
        "priority": "Monitor",
        "status": "On Track / Mixed",
        "recommended_focus": "Continue monitoring the current service plan.",
    }


def action(
    entity_key: str,
    person: str,
    priority: str,
    *,
    owner: str = "Pat Manager",
) -> dict:
    return {
        "Entity Key": entity_key,
        "Priority": priority,
        "Status": "Open",
        "Location": "RC Richmond",
        "Person / Area": person,
        "Action": "Coach Now",
        "Signal": "Traffic opportunity",
        "Why It Matters": "Guest traffic is below baseline | supporting detail",
        "Recommended Next Step": "Review the next two shifts",
        "Owner": owner,
        "Due Date": date(2026, 7, 15),
    }


def dashboard_inputs(source_days: int = metrics.OPERATING_WEEK_DAYS):
    week_end = date(2026, 7, 12)
    records = [
        record_for(week_end - timedelta(days=offset))
        for offset in range(metrics.OPERATING_WEEK_DAYS)
    ]
    if source_days < metrics.OPERATING_WEEK_DAYS:
        records = [record for record in records if record.report_date != date(2026, 7, 11)]
    weekly_locations = [
        weekly_location_row(week_end, location, source_days)
        for location in ("RC Richmond", "RC Virginia Beach")
    ]
    stores = [management_item(location) for location in ("RC Richmond", "RC Virginia Beach")]
    groups = [management_item("All Stores", group=True)]
    actions = [
        action("server|alex", "Alex", "High"),
        action("server|alex", "Alex duplicate", "Medium"),
        action("server|blair", "Blair", "Medium"),
        action("store|richmond", "RC Richmond", "Review"),
        action("server|casey", "Casey", "Recognize"),
    ]
    return records, weekly_locations, stores, groups, actions


def test_dashboard_is_one_screen_and_deduplicates_actions() -> None:
    wb = Workbook()
    records, weekly_locations, stores, groups, actions = dashboard_inputs()

    metrics.write_management_dashboard_sheet(
        wb, records, weekly_locations, [], [], stores, groups, actions
    )

    ws = wb["Dashboard"]
    values = {
        value
        for row in ws.iter_rows(min_row=1, max_row=24, min_col=1, max_col=12, values_only=True)
        for value in row
        if isinstance(value, str)
    }
    assert {"Reports Received", "Traffic vs Benchmark", "High-Priority Actions"}.issubset(values)
    assert {"TOP THREE ACTIONS", "STORE SNAPSHOT", "RECOGNITION / REPLICATE"}.issubset(values)
    assert ws["A7"].value == "6 of 6"
    assert ws["I7"].value == 3
    assert ws["C13"].value == "Alex | RC Richmond"
    assert ws["C14"].value == "Blair | RC Richmond"
    assert ws["C15"].value == "RC Richmond"
    assert "Alex duplicate" not in values
    assert "Casey" in ws["A23"].value
    assert len(ws._charts) == 0
    assert ws.max_row == 24
    assert "$A$1:$L$24" in str(ws.print_area)
    assert ws.freeze_panes == "A5"
    assert ws.sheet_view.zoomScale == 90


def test_incomplete_week_pauses_comparisons_actions_and_recognition() -> None:
    wb = Workbook()
    records, weekly_locations, stores, groups, actions = dashboard_inputs(source_days=5)

    metrics.write_management_dashboard_sheet(
        wb, records, weekly_locations, [], [], stores, groups, actions
    )

    ws = wb["Dashboard"]
    assert ws["A7"].value == "5 of 6"
    assert ws["E7"].value == "PAUSED"
    assert ws["I7"].value == "PAUSED"
    assert "Jul 11" in ws["A4"].value
    assert "Actions are paused" in ws["A13"].value
    assert "Recognition is paused" in ws["A23"].value
    assert ws["C19"].value == "Preliminary"
    assert ws["C20"].value == "Preliminary"
    assert len(ws._charts) == 0


def test_missing_configured_location_pauses_dashboard_without_stale_metrics() -> None:
    wb = Workbook()
    records, _weekly_locations, stores, groups, actions = dashboard_inputs()
    latest_week_end = date(2026, 7, 12)
    weekly_locations = [
        weekly_location_row(latest_week_end, "RC Richmond"),
        weekly_location_row(latest_week_end - timedelta(days=7), "RC Virginia Beach"),
    ]

    metrics.write_management_dashboard_sheet(
        wb,
        records,
        weekly_locations,
        [],
        [],
        stores,
        groups,
        actions,
        metrics.DEFAULT_CONFIG,
    )

    dashboard = wb["Dashboard"]
    assert dashboard["A7"].value == "6 of 6"
    assert dashboard["E7"].value == "PAUSED"
    assert dashboard["I7"].value == "PAUSED"
    assert "PRELIMINARY" in dashboard["A4"].value
    assert "RC Virginia Beach data" in dashboard["A4"].value
    assert dashboard["A19"].value == "RC Richmond"
    assert dashboard["A20"].value == "RC Virginia Beach"
    assert dashboard["C20"].value == "Missing"
    assert dashboard["E20"].value is None
    assert dashboard["F20"].value is None
    assert dashboard["G20"].value is None

    metrics.write_management_data_quality_sheet(
        wb, weekly_locations, metrics.DEFAULT_CONFIG
    )
    data_quality = wb["Data Quality"]
    assert "incomplete" in data_quality["A3"].value.lower()
    assert data_quality["B7"].value == "RC Virginia Beach"
    assert data_quality["E7"].value == "Missing"


def test_scorecard_charts_use_latest_eight_complete_weeks() -> None:
    wb = Workbook()
    first_week_end = date(2026, 5, 10)
    complete_week_ends = [first_week_end + timedelta(days=7 * offset) for offset in range(9)]
    partial_week_end = complete_week_ends[-1] + timedelta(days=7)
    weekly_locations = [
        weekly_location_row(week_end, location)
        for week_end in complete_week_ends
        for location in ("RC Richmond", "RC Virginia Beach")
    ]
    weekly_locations.extend(
        weekly_location_row(partial_week_end, location, 5)
        for location in ("RC Richmond", "RC Virginia Beach")
    )
    weekly_groups = [
        {"week_end": week_end, "guest_count": 900 + index * 10}
        for index, week_end in enumerate([*complete_week_ends, partial_week_end])
    ]

    metrics.write_store_group_scorecards_sheet(
        wb,
        [
            management_item("All Stores", group=True),
            management_item("RC Richmond"),
            management_item("RC Virginia Beach"),
        ],
        metrics.DEFAULT_CONFIG,
        weekly_locations,
        weekly_groups,
    )

    scorecards = wb["Store & Group Scorecards"]
    chart_data = wb["_Dashboard Chart Data"]
    assert len(scorecards._charts) == 2
    assert chart_data.sheet_state == "hidden"
    assert [chart_data.cell(row=row, column=1).value for row in range(2, 10)] == [
        week_end.strftime("%m/%d") for week_end in complete_week_ends[-8:]
    ]
    assert chart_data["A10"].value is None
    assert partial_week_end.strftime("%m/%d") not in {
        chart_data.cell(row=row, column=1).value for row in range(2, 12)
    }
    assert [series.tx.v for series in scorecards._charts[0].series] == [
        "Richmond",
        "Virginia Beach",
    ]
    assert scorecards._charts[0].series[0].cat.strRef.f.endswith("$A$2:$A$9")
    assert scorecards._charts[1].series[0].cat.strRef.f.endswith("$F$2:$F$9")
