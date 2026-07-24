from __future__ import annotations

from datetime import date, timedelta
import json
from pathlib import Path

import pytest

from red_onion_config import DEFAULT_CONFIG
import red_onion_model_validation as model_validation
import red_onion_weekly_metrics as weekly


def _all_metric_values(values: list[float]) -> dict[str, list[float]]:
    return {
        "check_average": list(values),
        "wine_pct": [value / 1000 for value in values],
        "rate_of_sale_by_guest_count": [value / 1000 for value in values],
        "average_ticket_time_seconds": [value * 60 for value in values],
    }


def test_calibration_freezes_separate_movement_and_peer_families() -> None:
    config = {
        **DEFAULT_CONFIG,
        "management_threshold_calibration": {
            **DEFAULT_CONFIG["management_threshold_calibration"],
            "version": "test-v1",
        },
    }

    result = model_validation.calibrate_deviation_sets(
        _all_metric_values([1, 2, 3, 4]),
        _all_metric_values([2, 4, 6, 8]),
        config,
        calibration_start="2026-01-06",
        calibration_end="2026-03-29",
    )

    fragment = result["config_fragment"]
    assert fragment["management_score_thresholds"]["check_average"] == {
        "neutral": 3.5,
        "strong": 5.0,
        "lower_is_better": False,
    }
    assert fragment["management_peer_score_thresholds"]["check_average"] == {
        "neutral": 6.5,
        "strong": 7.5,
        "lower_is_better": False,
    }
    metadata = fragment["management_threshold_calibration"]
    assert metadata["movement_observation_count"] == 16
    assert metadata["peer_observation_count"] == 16
    assert metadata["version"] == "test-v1"
    serialized = json.dumps(result)
    assert "raw_user_name" not in serialized
    assert "display_name" not in serialized


def test_calibration_requires_each_metric_family_to_have_evidence() -> None:
    movement = _all_metric_values([1, 2, 3])
    peer = _all_metric_values([1, 2, 3])
    peer["wine_pct"] = []

    with pytest.raises(ValueError, match="peer deviations.*wine_pct"):
        model_validation.calibrate_deviation_sets(
            movement,
            peer,
            DEFAULT_CONFIG,
            calibration_start=date(2026, 1, 6),
            calibration_end=date(2026, 3, 29),
        )


def test_backtest_summary_reports_rates_without_entity_details() -> None:
    first = date(2026, 1, 11)
    observations = [
        {
            "_entity_key": "entity-01",
            "location": "Store A",
            "week_end": first,
            "qualified": True,
            "action": "Context Review",
            "candidate_polarity": "positive",
            "stability_result": "Stable under every active-day removal",
        },
        {
            "_entity_key": "entity-01",
            "location": "Store A",
            "week_end": first + timedelta(days=7),
            "qualified": True,
            "action": "Coaching Prompt",
            "candidate_polarity": "negative",
            "stability_result": "Stable under every active-day removal",
        },
        {
            "_entity_key": "entity-02",
            "location": "Store A",
            "week_end": first,
            "qualified": True,
            "action": "Monitor",
            "candidate_polarity": "none",
            "stability_result": "Not applicable",
        },
        {
            "_entity_key": "entity-03",
            "location": "Store B",
            "week_end": first,
            "qualified": True,
            "action": "Monitor",
            "candidate_polarity": "none",
            "stability_result": "Not applicable",
        },
        {
            "_entity_key": "entity-04",
            "location": "Store B",
            "week_end": first,
            "qualified": False,
            "action": "Context Review",
            "candidate_polarity": "positive",
            "stability_result": "Sensitive",
        },
    ]

    result = model_validation.summarize_backtest_observations(observations)

    assert result["qualified_person_weeks"] == 4
    assert result["overall_review_action_rate"] == pytest.approx(0.5)
    assert result["maximum_store_week_review_action_rate"] == pytest.approx(1.0)
    assert result["prompt_stability_rate"] == pytest.approx(1.0)
    assert result["consecutive_candidate_pairs"] == 1
    assert result["consecutive_reversal_rate"] == pytest.approx(1.0)
    serialized = json.dumps(result)
    assert "entity-01" not in serialized
    assert "Store A" not in serialized


def test_report_directory_ingest_is_read_only_and_reconciles(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    report = tmp_path / "Daily Report - TM - 01-07-2026.xls"
    original = b"synthetic-read-only-source"
    report.write_bytes(original)
    business_date = date(2026, 1, 6)
    records = [
        weekly.MetricRecord(
            source_file=report.name,
            report_date=business_date,
            location="Store A",
            raw_user_name="entity-01",
            display_name="entity-01",
            is_location_total=False,
            gross_sales=100.0,
            guest_count=10.0,
            check_average=10.0,
            wine_sales=5.0,
            wine_pct=0.05,
            rate_of_sale_by_guest_count=0.1,
            average_ticket_time_seconds=600.0,
        ),
        weekly.MetricRecord(
            source_file=report.name,
            report_date=business_date,
            location="Store A",
            raw_user_name="",
            display_name="Store Total",
            is_location_total=True,
            gross_sales=100.0,
            guest_count=10.0,
            check_average=10.0,
            wine_sales=5.0,
            wine_pct=0.05,
            rate_of_sale_by_guest_count=0.1,
            average_ticket_time_seconds=600.0,
        ),
    ]
    monkeypatch.setattr(
        weekly,
        "read_reports_by_path",
        lambda paths, config: {Path(next(iter(paths))): records},
    )
    config = {
        **DEFAULT_CONFIG,
        "locations": {"Store A": {"short_code": "A"}},
    }
    before_entries = sorted(path.name for path in tmp_path.iterdir())

    dataset = model_validation.load_validated_report_directories(
        [tmp_path], config
    )

    assert dataset.source_report_count == 1
    assert dataset.business_dates == (business_date,)
    assert len(dataset.records) == 2
    assert report.read_bytes() == original
    assert sorted(path.name for path in tmp_path.iterdir()) == before_entries
