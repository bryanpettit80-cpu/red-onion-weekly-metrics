from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

import red_onion_config as config_module
import red_onion_weekly_metrics as metrics


def write_config(path: Path, payload: object) -> Path:
    path.write_text(json.dumps(payload), encoding="utf-8")
    return path


def test_missing_config_uses_valid_defaults(tmp_path: Path) -> None:
    config = metrics.load_config(tmp_path / "missing.json")

    assert config == config_module.DEFAULT_CONFIG
    assert config is not config_module.DEFAULT_CONFIG


def test_partial_config_is_deep_merged(tmp_path: Path) -> None:
    path = write_config(
        tmp_path / "config.json",
        {
            "management_score_thresholds": {
                "wine_pct": {
                    "neutral": 0.006,
                    "strong": 0.012,
                    "lower_is_better": False,
                }
            }
        },
    )

    config = metrics.load_config(path)

    assert config["management_score_thresholds"]["wine_pct"]["neutral"] == 0.006
    assert (
        config["management_score_thresholds"]["check_average"]["neutral"]
        == metrics.DEFAULT_CONFIG["management_score_thresholds"]["check_average"][
            "neutral"
        ]
    )


def test_weekly_area_and_shared_number_mappings_are_deep_merged(tmp_path: Path) -> None:
    path = write_config(
        tmp_path / "config.json",
        {
            "weekly_area_name_patterns": {"Wine Dinners": ["Special Event"]},
            "weekly_shared_number_areas": {"7070": "Wine Dinners"},
        },
    )

    config = metrics.load_config(path)

    assert config["weekly_area_name_patterns"]["Wine Dinners"] == ["Special Event"]
    assert config["weekly_area_name_patterns"]["Bar"] == ["Bar"]
    assert config["weekly_shared_number_areas"] == {"7070": "Wine Dinners"}


def test_validated_string_values_are_retained_in_normalized_form(
    tmp_path: Path,
) -> None:
    path = write_config(
        tmp_path / "config.json",
        {
            "locations": {
                " RC Richmond ": {"short_code": " RVA "},
            },
            "dashboard_exclude_name_contains": [" Banquet ", " Server "],
            "public_name_aliases": {" Special POS ": " Dining Room "},
            "weekly_area_name_patterns": {
                "Wine Dinners": [" Special Event "],
            },
            "weekly_shared_number_areas": {" 7070 ": " Wine Dinners "},
            "management_threshold_calibration": {
                "calibration_start": " 2026-03-24 ",
                "calibration_end": " 2026-07-19 ",
                "version": " 2026.07-v3 ",
            },
        },
    )

    config = metrics.load_config(path)

    assert " RC Richmond " not in config["locations"]
    assert config["locations"]["RC Richmond"]["short_code"] == "RVA"
    assert config["dashboard_exclude_name_contains"] == ["Banquet", "Server"]
    assert config["public_name_aliases"]["Special POS"] == "Dining Room"
    assert config["weekly_area_name_patterns"]["Wine Dinners"] == [
        "Special Event"
    ]
    assert config["weekly_shared_number_areas"] == {"7070": "Wine Dinners"}
    assert config["management_threshold_calibration"]["calibration_start"] == (
        "2026-03-24"
    )
    assert config["management_threshold_calibration"]["calibration_end"] == (
        "2026-07-19"
    )
    assert config["management_threshold_calibration"]["version"] == "2026.07-v3"


@pytest.mark.parametrize(
    ("payload", "message"),
    [
        ({"surprise": True}, "Unknown configuration field: surprise"),
        ({"locations": {}}, "at least one location"),
        (
            {
                "locations": {
                    "One": {"short_code": "X"},
                    "Two": {"short_code": "x"},
                }
            },
            "short_code values must be unique",
        ),
        (
            {"locations": {"New Site": {"short_code": "RVA"}}},
            "short_code values must be unique",
        ),
        (
            {"locations": {"rc richmond": {"short_code": "NEW"}}},
            "names must be unique ignoring case",
        ),
        (
            {"management_materiality": {"wine_pct": -0.1}},
            "management_materiality.wine_pct must be greater than zero",
        ),
        (
            {
                "management_score_thresholds": {
                    "wine_pct": {
                        "neutral": 0.02,
                        "strong": 0.01,
                        "lower_is_better": False,
                    }
                }
            },
            "strong must be greater than neutral",
        ),
        (
            {
                "management_peer_score_thresholds": {
                    "check_average": {
                        "neutral": 5.0,
                        "strong": 5.0,
                        "lower_is_better": False,
                    }
                }
            },
            "strong must be greater than neutral",
        ),
        (
            {"dashboard_long_term_full_weeks": 7},
            "must equal two times",
        ),
        (
            {"management_materiality": {"sales_pct": 1.5}},
            "management_materiality.sales_pct cannot exceed 1",
        ),
        (
            {
                "management_peer_reference": {
                    "leave_one_person_out": False,
                }
            },
            "leave_one_person_out must be true",
        ),
        (
            {
                "management_signal_persistence": {
                    "require_recurring_driver": False,
                }
            },
            "require_recurring_driver must be true",
        ),
        (
            {
                "management_signal_persistence": {
                    "require_leave_one_active_day_stability": False,
                }
            },
            "require_leave_one_active_day_stability must be true",
        ),
        (
            {"weekly_area_name_patterns": {"Roof": ["Roof"]}},
            "Unknown configuration field: weekly_area_name_patterns.Roof",
        ),
        (
            {"weekly_shared_number_areas": {"70": "Wine Dinners"}},
            "four-digit POS numbers",
        ),
        (
            {"weekly_shared_number_areas": {"7070": "Roof"}},
            "must name a configured weekly area",
        ),
    ],
)
def test_invalid_config_is_rejected(
    tmp_path: Path, payload: object, message: str
) -> None:
    path = write_config(tmp_path / "config.json", payload)

    with pytest.raises(config_module.ConfigError, match=message):
        metrics.load_config(path)


def test_duplicate_json_key_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "config.json"
    path.write_text(
        '{"public_min_guest_count": 1, "public_min_guest_count": 2}',
        encoding="utf-8",
    )

    with pytest.raises(config_module.ConfigError, match="Duplicate configuration field"):
        metrics.load_config(path)


def test_standard_library_config_cli_has_precise_failure(tmp_path: Path) -> None:
    path = write_config(tmp_path / "config.json", {"locations": {"One": {"typo": "X"}}})

    result = subprocess.run(
        [sys.executable, str(Path(config_module.__file__)), "--config", str(path)],
        capture_output=True,
        text=True,
    )

    assert result.returncode != 0
    assert "Unknown configuration field: locations.One.typo" in result.stderr
