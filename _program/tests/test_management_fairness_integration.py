from __future__ import annotations

from copy import deepcopy
from datetime import date, timedelta
from types import SimpleNamespace

import pytest

import red_onion_weekly_metrics as metrics


def weekly_row(
    week_end: date,
    person: str,
    *,
    check_average: float,
    wine_pct: float,
    rate: float,
    ticket_seconds: float,
) -> dict:
    guests = 100.0
    gross_sales = check_average * guests
    return {
        "week_start": week_end - timedelta(days=5),
        "week_end": week_end,
        "location": "RC Richmond",
        "raw_user_name": person,
        "display_name": person,
        "gross_sales": gross_sales,
        "guest_count": guests,
        "check_average": check_average,
        "wine_sales": gross_sales * wine_pct,
        "wine_pct": wine_pct,
        "rate_of_sale_by_guest_count": rate,
        "average_ticket_time_seconds": ticket_seconds,
        "active_days": 6,
        "source_days": 6,
        "rate_available": True,
        "ticket_time_available": True,
        "source_evidence": [],
        "daily_records": [],
    }


def synthetic_history() -> tuple[list[dict], list[dict], dict]:
    first_week = date(2026, 6, 14)
    rows: list[dict] = []
    locations: list[dict] = []
    for index in range(6):
        week_end = first_week + timedelta(days=7 * index)
        improved = index >= 4
        rows.append(
            weekly_row(
                week_end,
                "entity-focus",
                check_average=35.0 if improved else 10.0,
                wine_pct=0.20 if improved else 0.01,
                rate=0.10 if improved else 0.40,
                ticket_seconds=3000.0 if improved else 7200.0,
            )
        )
        for peer_index in range(6):
            rows.append(
                weekly_row(
                    week_end,
                    f"entity-peer-{peer_index}",
                    check_average=20.0,
                    wine_pct=0.08,
                    rate=0.25,
                    ticket_seconds=5000.0,
                )
            )
        locations.append(
            {
                **weekly_row(
                    week_end,
                    "location-total",
                    check_average=20.0,
                    wine_pct=0.08,
                    rate=0.25,
                    ticket_seconds=5000.0,
                ),
                "location": "RC Richmond",
            }
        )
    config = deepcopy(metrics.DEFAULT_CONFIG)
    config["locations"] = {"RC Richmond": {"short_code": "RVA"}}
    for family in (
        "management_score_thresholds",
        "management_peer_score_thresholds",
    ):
        config[family] = {
            "check_average": {
                "neutral": 2.5,
                "strong": 5.0,
                "lower_is_better": False,
            },
            "wine_pct": {
                "neutral": 0.005,
                "strong": 0.01,
                "lower_is_better": False,
            },
            "rate_of_sale_by_guest_count": {
                "neutral": 0.005,
                "strong": 0.01,
                "lower_is_better": True,
            },
            "average_ticket_time_seconds": {
                "neutral": 150.0,
                "strong": 300.0,
                "lower_is_better": True,
            },
        }
    return rows, locations, config


def focal_result(
    rows: list[dict],
    locations: list[dict],
    config: dict,
    *,
    ranked_rows: list[dict] | None = None,
    targets: dict | None = None,
) -> dict:
    output = metrics.management_server_rows(
        rows,
        locations,
        ranked_rows or [],
        targets or {},
        config,
    )
    return next(
        row for row in output if row["raw_user_name"] == "entity-focus"
    )


def test_rank_and_excluded_rows_cannot_change_a_person_signal() -> None:
    rows, locations, config = synthetic_history()
    baseline = focal_result(rows, locations, config)
    ranked = [
        {
            **row,
            "check_average_rank": 999,
            "wine_pct_rank": 999,
            "rate_rank": 999,
            "ticket_time_rank": 999,
        }
        for row in rows
    ]
    contaminated = list(rows)
    for location_row in locations:
        contaminated.append(
            weekly_row(
                location_row["week_end"],
                "Banquet excluded extreme",
                check_average=10000.0,
                wine_pct=0.99,
                rate=0.0001,
                ticket_seconds=1.0,
            )
        )
    changed = focal_result(
        contaminated,
        locations,
        config,
        ranked_rows=ranked,
    )

    for field in (
        "action",
        "momentum",
        "performance_level",
        "candidate_polarity",
        "composite_score",
        "peer_composite_score",
    ):
        assert changed[field] == baseline[field]
    assert changed["rank_modifier"] == baseline["rank_modifier"] == 0
    assert changed["average_rank_movement"] is None


def test_person_targets_do_not_replace_leave_one_out_peer_reference() -> None:
    rows, locations, config = synthetic_history()
    baseline = focal_result(rows, locations, config)
    absurd_targets = {
        "RC Richmond": {
            field: -999999.0 for field in metrics.SERVER_TREND_FIELDS
        }
    }
    targeted = focal_result(
        rows,
        locations,
        config,
        targets=absurd_targets,
    )

    assert targeted["action"] == baseline["action"]
    assert targeted["performance_level"] == baseline["performance_level"]
    assert targeted["benchmark_values"] == baseline["benchmark_values"]
    assert set(targeted["benchmark_sources"].values()) == {
        "Same-store prior-four-week median"
    }
    assert targeted["peer_cohort_size"] == 6


def test_context_only_metrics_cannot_change_a_person_action() -> None:
    rows, locations, config = synthetic_history()
    baseline = focal_result(rows, locations, config)
    changed_rows = deepcopy(rows)
    for row in changed_rows:
        if row["raw_user_name"] == "entity-focus":
            row["rate_of_sale_by_guest_count"] = 99_999.0
            row["average_ticket_time_seconds"] = 999_999.0
            row["rate_available"] = False
            row["ticket_time_available"] = False
            row["check_count"] = 1.0
            row["check_count_available"] = True
    changed = focal_result(changed_rows, locations, config)

    for field in (
        "action",
        "momentum",
        "performance_level",
        "candidate_polarity",
        "composite_score",
        "peer_composite_score",
        "persistence_reason",
    ):
        assert changed[field] == baseline[field]
    assert tuple(metrics.SERVER_TREND_FIELDS) == (
        "check_average",
        "wine_pct",
    )


def test_unavailable_rate_and_ticket_do_not_feed_long_term_scoring(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    rows, locations, config = synthetic_history()
    monkeypatch.setattr(
        metrics,
        "SERVER_TREND_FIELDS",
        (*metrics.SERVER_PERSON_ACTION_FIELDS, *metrics.SERVER_CONTEXT_FIELDS),
    )
    for row in rows:
        if row["raw_user_name"] == "entity-focus":
            row["rate_of_sale_by_guest_count"] = 99_999.0
            row["average_ticket_time_seconds"] = 999_999.0
            row["rate_available"] = False
            row["ticket_time_available"] = False

    result = focal_result(rows, locations, config)

    for field in metrics.SERVER_CONTEXT_FIELDS:
        assert result["long_term_changes"][field] is None
        assert result["long_term_metric_scores"][field] is None
    assert result["long_term_direction"] == "Not Evaluated"


def test_unavailable_baseline_metric_is_a_data_issue_not_eligible(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    rows, locations, config = synthetic_history()
    latest_week_end = max(row["week_end"] for row in locations)
    current = next(
        row
        for row in rows
        if row["raw_user_name"] == "entity-focus"
        and row["week_end"] == latest_week_end
    )
    full_by_location = {
        "RC Richmond": {row["week_end"] for row in locations}
    }
    original_aggregate = metrics.aggregate_weekly_rows

    def aggregate_with_unavailable_baseline(selected):
        result = original_aggregate(selected)
        if result is not None:
            result = dict(result)
            result["wine_pct"] = float("nan")
        return result

    monkeypatch.setattr(
        metrics,
        "aggregate_weekly_rows",
        aggregate_with_unavailable_baseline,
    )

    evaluated = metrics.evaluate_server_week_signal(
        current,
        rows,
        full_by_location,
        config,
    )

    assert evaluated["metric_scores"]["wine_pct"] is None
    assert evaluated["candidate_qualified"] is False
    assert evaluated["evidence_status"] == "Data Issue"


def test_store_shock_only_failure_is_not_labeled_day_sensitive(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    rows, locations, config = synthetic_history()
    monkeypatch.setattr(
        metrics,
        "assess_common_store_shock",
        lambda *_args, **_kwargs: SimpleNamespace(
            guard_passed=False,
            common_store_shock=True,
            comparable_metrics=("check_average", "wine_pct"),
            differentiating_metrics=(),
            relative_scores=(),
            reason="common_store_shock",
        ),
    )

    result = focal_result(rows, locations, config)

    assert result["action"] == "Context Review"
    assert result["confidence"] == "Store Shock"
    assert result["persistence_reason"] == "store_shock_guard_not_passed"
    assert result["stability_result"] == (
        "Store-shock guard not passed; day-removal stability not applicable"
    )


def test_low_volume_server_keeps_descriptive_wow_but_remains_monitor() -> None:
    rows, locations, config = synthetic_history()
    latest_week_end = max(row["week_end"] for row in locations)
    latest = next(
        row
        for row in rows
        if row["raw_user_name"] == "entity-focus"
        and row["week_end"] == latest_week_end
    )
    latest.update(
        {
            "gross_sales": 800.0,
            "guest_count": 20.0,
            "check_average": 40.0,
            "wine_sales": 200.0,
            "wine_pct": 0.25,
            "active_days": 2,
        }
    )

    result = focal_result(rows, locations, config)

    assert result["confidence"] == "Limited Volume"
    assert result["momentum"] == "Not Evaluated"
    assert result["descriptive_recent_movement"] == "Upward"
    assert result["previous_week_end"] == latest_week_end - timedelta(days=7)
    assert result["week_over_week_changes"] == pytest.approx(
        {
            "check_average": 5.0,
            "wine_pct": 0.05,
        }
    )
    assert result["week_over_week_movement"] == "Improving"
    assert result["action"] == "Monitor"
    assert result["prominent"] is False
    assert (
        metrics.build_management_action_signals(
            [result],
            [],
            [],
            locations,
        )
        == []
    )


def test_missing_immediately_prior_calendar_week_leaves_wow_blank() -> None:
    rows, locations, config = synthetic_history()
    latest_week_end = max(row["week_end"] for row in locations)
    immediately_prior = latest_week_end - timedelta(days=7)
    rows = [
        row
        for row in rows
        if not (
            row["raw_user_name"] == "entity-focus"
            and row["week_end"] == immediately_prior
        )
    ]

    result = focal_result(rows, locations, config)

    assert result["previous_week_end"] is None
    assert result["week_over_week_changes"] == {
        "check_average": None,
        "wine_pct": None,
    }
    assert result["week_over_week_movement"] == "No prior week"


@pytest.mark.parametrize(
    "value",
    [None, "", "not-a-number", float("nan"), float("inf")],
)
def test_missing_rate_values_are_unavailable_not_favorable_zero(value) -> None:
    parsed, available = metrics.optional_float(value)
    assert parsed == 0.0
    assert available is False


@pytest.mark.parametrize(
    "value",
    [None, "", "not-a-time", float("nan"), float("inf")],
)
def test_missing_ticket_values_are_unavailable_not_favorable_zero(value) -> None:
    parsed, available = metrics.optional_ticket_time_seconds(value)
    assert parsed == 0.0
    assert available is False


def metric_record(
    *,
    is_total: bool,
    gross_sales: float,
    guests: float,
    wine_sales: float,
) -> metrics.MetricRecord:
    return metrics.MetricRecord(
        source_file="synthetic.xls",
        report_date=date(2026, 7, 14),
        location="RC Richmond",
        raw_user_name="" if is_total else "entity-01",
        display_name="total" if is_total else "entity-01",
        is_location_total=is_total,
        gross_sales=gross_sales,
        guest_count=guests,
        check_average=gross_sales / guests,
        wine_sales=wine_sales,
        wine_pct=wine_sales / gross_sales,
        rate_of_sale_by_guest_count=0.20,
        average_ticket_time_seconds=4800.0,
    )


def test_daily_reconciliation_fails_closed_on_material_mismatch() -> None:
    records = [
        metric_record(
            is_total=True,
            gross_sales=100.0,
            guests=10.0,
            wine_sales=5.0,
        ),
        metric_record(
            is_total=False,
            gross_sales=99.98,
            guests=9.0,
            wine_sales=4.98,
        ),
    ]

    with pytest.raises(ValueError, match="Daily reconciliation failed"):
        metrics.validate_daily_location_reconciliation(
            records,
            metrics.DEFAULT_CONFIG,
        )
