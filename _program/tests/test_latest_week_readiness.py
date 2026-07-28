from __future__ import annotations

from datetime import date, timedelta
import json

from openpyxl import Workbook

import red_onion_weekly_metrics as metrics


LATEST_WEEK_END = date(2026, 7, 12)


def metric_record(report_date: date) -> metrics.MetricRecord:
    return metrics.MetricRecord(
        source_file=f"daily-{report_date.isoformat()}.xlsx",
        report_date=report_date,
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


def report_records(*, missing: date | None = None) -> list[metrics.MetricRecord]:
    week_start = LATEST_WEEK_END - timedelta(days=metrics.OPERATING_WEEK_DAYS - 1)
    return [
        metric_record(week_start + timedelta(days=offset))
        for offset in range(metrics.OPERATING_WEEK_DAYS)
        if week_start + timedelta(days=offset) != missing
    ]


def weekly_location(
    location: str,
    *,
    week_end: date = LATEST_WEEK_END,
    source_days: int = metrics.OPERATING_WEEK_DAYS,
) -> dict:
    return {
        "week_start": week_end - timedelta(days=metrics.OPERATING_WEEK_DAYS - 1),
        "week_end": week_end,
        "location": location,
        "gross_sales": 12000.0,
        "guest_count": 500.0,
        "check_average": 24.0,
        "wine_sales": 1440.0,
        "wine_pct": 0.12,
        "rate_of_sale_by_guest_count": 0.18,
        "average_ticket_time_seconds": 72 * 60.0,
        "active_days": source_days,
        "source_days": source_days,
    }


def server_candidate() -> dict:
    return {
        "prominent": True,
        "action": "Recognize & Replicate",
        "priority": "Recognize",
        "location": "RC Richmond",
        "raw_user_name": "Alex",
        "display_name": "Alex",
        "momentum": "Rising",
        "performance_level": "Above Benchmark",
        "long_term_direction": "Improving",
        "long_term_history_label": "Full",
        "positive_drivers": ["Check avg +$5"],
        "negative_drivers": [],
        "guest_count": 80.0,
        "active_days": metrics.OPERATING_WEEK_DAYS,
        "confidence": "High",
        "week_end": LATEST_WEEK_END,
        "recommended_next_step": "Recognize the improvement.",
    }


def entity_candidate(entity: str, *, week_end: date = LATEST_WEEK_END) -> dict:
    latest = {
        **weekly_location(entity, week_end=week_end),
        "week_end": week_end,
    }
    benchmark = {
        **latest,
        "gross_sales": 14000.0,
        "guest_count": 600.0,
    }
    fields = [field for field, _, _ in metrics.MANAGEMENT_METRICS]
    return {
        "entity": entity,
        "latest": latest,
        "prior": latest,
        "baseline": benchmark,
        "baseline_weeks": 4,
        "benchmark_values": {field: benchmark[field] for field in fields},
        "benchmark_sources": {field: "4-week baseline" for field in fields},
        "prior_changes": {field: 0.0 for field in fields},
        "benchmark_changes": {
            field: latest[field] - benchmark[field] for field in fields
        },
        "priority": "High",
        "status": "Traffic Watch",
        "recommended_focus": "Review traffic drivers.",
    }


def group_candidate() -> dict:
    item = entity_candidate("All Stores")
    item["latest"]["group"] = "All Stores"
    return item


def configured_locations(*, source_days: int = metrics.OPERATING_WEEK_DAYS) -> list[dict]:
    return [
        weekly_location("RC Richmond", source_days=source_days),
        weekly_location("RC Virginia Beach", source_days=source_days),
    ]


def test_missing_daily_report_suppresses_new_performance_and_recognition_actions() -> None:
    missing_date = date(2026, 7, 11)
    locations = configured_locations(source_days=metrics.OPERATING_WEEK_DAYS - 1)
    readiness = metrics.latest_week_readiness(
        report_records(missing=missing_date), locations, metrics.DEFAULT_CONFIG
    )

    signals = metrics.build_management_action_signals(
        [server_candidate()],
        [entity_candidate("RC Richmond")],
        [group_candidate()],
        locations,
        readiness,
    )

    assert readiness.ready is False
    assert readiness.missing_dates == (missing_date,)
    assert signals
    assert {signal["Action"] for signal in signals} == {"Data Quality"}
    assert {signal["Location"] for signal in signals} == {
        "RC Richmond",
        "RC Virginia Beach",
    }


def test_missing_configured_store_suppresses_actions_and_blanks_stale_scorecard() -> None:
    prior_week_end = LATEST_WEEK_END - timedelta(days=7)
    locations = [
        weekly_location("RC Richmond"),
        weekly_location("RC Virginia Beach", week_end=prior_week_end),
    ]
    readiness = metrics.latest_week_readiness(
        report_records(), locations, metrics.DEFAULT_CONFIG
    )
    stale_store = entity_candidate("RC Virginia Beach", week_end=prior_week_end)
    stale_store["latest"]["gross_sales"] = 999999.0

    signals = metrics.build_management_action_signals(
        [server_candidate()],
        [entity_candidate("RC Richmond"), stale_store],
        [group_candidate()],
        locations,
        readiness,
    )

    assert readiness.ready is False
    assert readiness.missing_dates == ()
    assert readiness.location_gaps == ("RC Virginia Beach",)
    assert {signal["Action"] for signal in signals} == {"Data Quality"}
    assert [signal["Location"] for signal in signals] == ["RC Virginia Beach"]

    wb = Workbook()
    metrics.write_store_group_scorecards_sheet(
        wb,
        [group_candidate(), entity_candidate("RC Richmond"), stale_store],
        metrics.DEFAULT_CONFIG,
        locations,
        [],
        readiness,
    )

    scorecards = wb["Store & Group Scorecards"]
    assert "Preliminary" in scorecards["A4"].value
    assert "Preliminary" in scorecards["A17"].value
    assert scorecards["A30"].value.startswith("RC Virginia Beach | Missing")
    assert scorecards["B32"].value is None
    assert scorecards["F32"].value == "No current-week data"
    assert scorecards["G32"].value == "Missing"
    assert 999999.0 not in {
        cell.value for row in scorecards.iter_rows() for cell in row
    }


def test_incomplete_week_carries_manual_active_action_as_paused() -> None:
    missing_date = date(2026, 7, 11)
    locations = configured_locations(source_days=metrics.OPERATING_WEEK_DAYS - 1)
    readiness = metrics.latest_week_readiness(
        report_records(missing=missing_date), locations, metrics.DEFAULT_CONFIG
    )
    prior = {
        "Action ID": "A1B2C3D4E5F6",
        "Entity Key": "server|rc richmond|alex|coaching",
        "Priority": "High",
        "Status": "In Progress",
        "Owner": "Pat Manager",
        "Due Date": date(2026, 7, 15),
        "Location": "RC Richmond",
        "Person / Area": "Alex",
        "Action": "Coach Now",
        "Signal": "Falling / Below Benchmark",
        "Last Seen": date(2026, 7, 5),
        "Manager Notes": "Review Friday",
        "First Seen": date(2026, 6, 28),
        "Weeks Open": 2,
        "Confidence": "High",
        "Signal State": "Current",
        "Evidence ID": "STALE-EVIDENCE",
        "Action Code": "COACH_NOW",
        "Reason Code": "SERVER_FALLING_BELOW_BENCHMARK",
        "Evidence Week Ends": "2026-06-28, 2026-07-05",
        "Evidence Sources": "[]",
        "Metric Evidence": '{"guest_count":42}',
        "Methodology Version": metrics.MANAGEMENT_METHODOLOGY_VERSION,
    }
    stale_data_quality = {
        "Action ID": "D1E2F3A4B5C6",
        "Entity Key": "data-quality|rc richmond|short-week",
        "Priority": "Review",
        "Status": "Open",
        "Location": "RC Richmond",
        "Person / Area": "RC Richmond",
        "Action": "Data Quality",
        "Signal": "Incomplete Latest Week",
        "Last Seen": date(2026, 7, 5),
        "First Seen": date(2026, 7, 5),
        "Weeks Open": 1,
        "Confidence": "Low Sample",
        "Signal State": "Current",
    }

    current, history = metrics.merge_management_actions(
        [], {"active_actions": [prior, stale_data_quality], "action_history": []}, readiness
    )

    assert history == []
    assert len(current) == 2
    current_by_id = {row["Action ID"]: row for row in current}
    paused = current_by_id[prior["Action ID"]]
    cleared_follow_up = current_by_id[stale_data_quality["Action ID"]]
    assert cleared_follow_up["Status"] == "Review Needed"
    assert cleared_follow_up["Review Disposition"] == "Pending Review"
    assert cleared_follow_up["Signal State"] == "Cleared / Follow-up Required"
    assert "never resolved" in cleared_follow_up["Recommended Next Step"]
    assert paused["Status"] == "Review Needed"
    assert paused["Owner"] == "Pat Manager"
    assert paused["Due Date"] == date(2026, 7, 15)
    assert paused["Context Notes"] == "Review Friday"
    assert paused["Priority"] == "Paused"
    assert paused["Action"] == "Paused Carryover"
    assert paused["Signal"].startswith("PAUSED / CARRYOVER")
    assert paused["Why It Matters"].startswith("Prior action retained")
    assert "manual assignment on hold" in paused["Recommended Next Step"]
    assert paused["Peer Comparison"] == "Preliminary"
    assert paused["Recent Movement"] == "Not Evaluated"
    assert paused["Evidence Status"] == "Paused"
    assert paused["Review Disposition"] == "Pending Review"
    assert paused["Signal State"] == "Paused / Carryover"
    assert paused["Action Code"] == "PAUSED_CARRYOVER"
    assert (
        paused["Reason Code"]
        == "LATEST_WEEK_INCOMPLETE_PRIOR_ACTION_RETAINED"
    )
    assert paused["Evidence ID"] != "STALE-EVIDENCE"
    assert "2026-07-12" in paused["Evidence Week Ends"]
    paused_evidence = json.loads(paused["Metric Evidence"])[
        "paused_carryover"
    ]
    assert paused_evidence["latest_week_end"] == "2026-07-12"
    assert paused_evidence["missing_dates"] == ["2026-07-11"]
    assert "Jul 11" in paused_evidence["missing_text"]
    assert paused["Last Seen"] == date(2026, 7, 12)
    assert paused["First Seen"] == date(2026, 6, 28)
    assert paused["Weeks Open"] == 3

    repeated, _ = metrics.merge_management_actions(
        [], {"active_actions": current, "action_history": history}, readiness
    )
    repeated_by_id = {row["Action ID"]: row for row in repeated}
    repeated_paused = repeated_by_id[prior["Action ID"]]
    assert repeated_paused["Last Seen"] == date(2026, 7, 12)
    assert repeated_paused["First Seen"] == date(2026, 6, 28)
    assert repeated_paused["Weeks Open"] == 3
    assert repeated_paused["Evidence ID"] == paused["Evidence ID"]


def test_backdated_incomplete_week_does_not_rewind_paused_action_dates() -> None:
    readiness = metrics.latest_week_readiness(
        report_records(missing=date(2026, 7, 11)),
        configured_locations(source_days=metrics.OPERATING_WEEK_DAYS - 1),
        metrics.DEFAULT_CONFIG,
    )
    prior = {
        "Action ID": "A1B2C3D4E5F6",
        "Entity Key": "server|rc richmond|alex|coaching",
        "Priority": "High",
        "Status": "In Progress",
        "Location": "RC Richmond",
        "Person / Area": "Alex",
        "Action": "Coach Now",
        "Signal": "Falling / Below Benchmark",
        "First Seen": date(2026, 7, 19),
        "Last Seen": date(2026, 7, 26),
        "Weeks Open": 2,
    }

    current, _ = metrics.merge_management_actions(
        [], {"active_actions": [prior], "action_history": []}, readiness
    )

    assert current[0]["First Seen"] == date(2026, 7, 19)
    assert current[0]["Last Seen"] == date(2026, 7, 26)
    assert current[0]["Weeks Open"] == 2


def test_complete_week_shared_readiness_preserves_action_signal_behavior() -> None:
    locations = configured_locations()
    readiness = metrics.latest_week_readiness(
        report_records(), locations, metrics.DEFAULT_CONFIG
    )
    args = (
        [server_candidate()],
        [entity_candidate("RC Richmond")],
        [group_candidate()],
        locations,
    )

    legacy_signals = metrics.build_management_action_signals(*args)
    shared_signals = metrics.build_management_action_signals(*args, readiness)

    assert readiness.ready is True
    assert shared_signals == legacy_signals
    assert {signal["Action"] for signal in shared_signals} == {
        "Recognize & Replicate",
        "Store Review",
        "Group Review",
    }
