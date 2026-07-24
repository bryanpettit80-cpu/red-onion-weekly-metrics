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
