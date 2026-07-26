from __future__ import annotations

from copy import deepcopy
from datetime import date

from openpyxl import Workbook
import pytest

import red_onion_weekly_metrics as metrics


def record(
    name: str,
    *,
    gross_sales: float,
    guests: float,
    checks: float,
    rate: float,
    ticket_seconds: float,
    check_count_available: bool = True,
) -> metrics.MetricRecord:
    return metrics.MetricRecord(
        source_file="Daily Report - synthetic.xlsx",
        report_date=date(2026, 7, 23),
        location="RC Richmond",
        raw_user_name=name,
        display_name=name,
        is_location_total=False,
        gross_sales=gross_sales,
        guest_count=guests,
        check_average=gross_sales / guests if guests else 0.0,
        wine_sales=0.0,
        wine_pct=0.0,
        rate_of_sale_by_guest_count=rate,
        average_ticket_time_seconds=ticket_seconds,
        rate_available=True,
        ticket_time_available=True,
        check_count=checks,
        check_count_available=check_count_available,
    )


def test_operational_rollup_uses_harmonic_ros_and_check_weighted_ticket_time() -> None:
    source = [
        record(
            "A",
            gross_sales=1_000,
            guests=100,
            checks=10,
            rate=5,
            ticket_seconds=100,
        ),
        record(
            "B",
            gross_sales=500,
            guests=50,
            checks=30,
            rate=10,
            ticket_seconds=300,
        ),
    ]

    rollup = metrics.aggregate_records(source, ("location",))[0]

    # ROS = total opportunities / total inferred qualifying sales.
    assert rollup["rate_of_sale_by_guest_count"] == pytest.approx(150 / (20 + 5))
    # Ticket Time is weighted by checks, not guests.
    assert rollup["average_ticket_time_seconds"] == pytest.approx(
        (100 * 10 + 300 * 30) / 40
    )
    assert rollup["check_count"] == 40
    assert rollup["check_count_available"] is True
    assert rollup["sales_per_check"] == pytest.approx(37.5)
    assert rollup["guests_per_check"] == pytest.approx(3.75)
    assert rollup["ticket_time_weight_basis"] == "Check Count"


def test_all_dict_rollup_paths_use_the_same_operational_semantics() -> None:
    rows = [
        {
            "report_date": date(2026, 7, 22),
            "gross_sales": 1_000.0,
            "guest_count": 100.0,
            "wine_sales": 0.0,
            "rate_of_sale_by_guest_count": 5.0,
            "average_ticket_time_seconds": 100.0,
            "rate_available": True,
            "ticket_time_available": True,
            "check_count": 10.0,
            "check_count_available": True,
        },
        {
            "report_date": date(2026, 7, 23),
            "gross_sales": 500.0,
            "guest_count": 50.0,
            "wine_sales": 0.0,
            "rate_of_sale_by_guest_count": 10.0,
            "average_ticket_time_seconds": 300.0,
            "rate_available": True,
            "ticket_time_available": True,
            "check_count": 30.0,
            "check_count_available": True,
        },
    ]

    weekly = metrics.aggregate_weekly_rows(rows)
    from_daily = metrics.weekly_row_from_daily_records(rows)

    assert weekly is not None and from_daily is not None
    for result in (weekly, from_daily):
        assert result["rate_of_sale_by_guest_count"] == pytest.approx(6.0)
        assert result["average_ticket_time_seconds"] == pytest.approx(250.0)
        assert result["check_count"] == 40
        assert result["sales_per_check"] == pytest.approx(37.5)
        assert result["guests_per_check"] == pytest.approx(3.75)


def test_legacy_rows_never_fall_back_to_guest_weighted_ticket_time() -> None:
    legacy = [
        record(
            "A",
            gross_sales=1_000,
            guests=100,
            checks=0,
            rate=5,
            ticket_seconds=100,
            check_count_available=False,
        ),
        record(
            "B",
            gross_sales=500,
            guests=50,
            checks=0,
            rate=10,
            ticket_seconds=300,
            check_count_available=False,
        ),
    ]

    rollup = metrics.aggregate_records(legacy, ("location",))[0]

    assert rollup["check_count_available"] is False
    assert rollup["ticket_time_available"] is False
    assert rollup["average_ticket_time_seconds"] == 0
    assert rollup["ticket_time_weight_basis"].startswith("Unavailable")


def test_hidden_summaries_leave_unavailable_ticket_time_blank() -> None:
    legacy_week = metrics.aggregate_records(
        [
            record(
                "A",
                gross_sales=1_000,
                guests=100,
                checks=0,
                rate=5,
                ticket_seconds=100,
                check_count_available=False,
            )
        ],
        ("location", "raw_user_name", "display_name"),
        {
            "week_start": date(2026, 7, 14),
            "week_end": date(2026, 7, 19),
        },
    )

    summary = metrics.trend_summary_rows(legacy_week)[0]

    assert summary["weighted_ticket_time_seconds"] is None
    assert summary["latest_ticket_time_seconds"] is None
    assert summary["ticket_time_change_minutes"] is None
    assert metrics.duration_fraction(None) is None


def test_data_quality_does_not_call_partial_context_complete() -> None:
    workbook = Workbook()
    config = deepcopy(metrics.DEFAULT_CONFIG)
    config["locations"] = {"RC Richmond": {"short_code": "RVA"}}
    weekly_location = {
        "week_start": date(2026, 7, 14),
        "week_end": date(2026, 7, 19),
        "location": "RC Richmond",
        "active_days": 6,
        "source_days": 6,
        "check_count": 100.0,
        "check_count_available": True,
        "ticket_time_available": False,
        "ticket_time_weight_basis": (
            "Unavailable (Check Count missing or incomplete)"
        ),
    }

    metrics.write_management_data_quality_sheet(
        workbook,
        [weekly_location],
        config,
    )

    quality = workbook["Data Quality"]
    assert quality["F6"].value.startswith("Unavailable")
    assert quality["H6"].value == (
        "People review eligible; Ticket Time unavailable"
    )
