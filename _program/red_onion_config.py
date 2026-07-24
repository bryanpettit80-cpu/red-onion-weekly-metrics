from __future__ import annotations

import argparse
import copy
from datetime import date
import json
import math
from pathlib import Path
from typing import Any


DEFAULT_CONFIG: dict[str, Any] = {
    "locations": {
        "RC Richmond": {"short_code": "RVA"},
        "RC Virginia Beach": {"short_code": "VB"},
    },
    "public_min_guest_count": 1,
    "master_min_guest_count_for_rankings": 1,
    "dashboard_min_guest_count_for_trends": 25,
    "dashboard_min_active_days_for_trends": 3,
    "dashboard_min_prior_full_weeks": 2,
    "dashboard_min_prior_guest_count": 50,
    "dashboard_baseline_full_weeks": 4,
    "dashboard_long_term_full_weeks": 8,
    "dashboard_long_term_block_weeks": 4,
    "dashboard_long_term_full_min_recent_guests": 100,
    "dashboard_long_term_full_min_earlier_guests": 100,
    "dashboard_long_term_developing_min_total_weeks": 6,
    "dashboard_long_term_developing_min_recent_weeks": 3,
    "dashboard_long_term_developing_min_earlier_weeks": 2,
    "dashboard_long_term_developing_min_recent_guests": 75,
    "dashboard_long_term_developing_min_earlier_guests": 50,
    "dashboard_exclude_name_contains": ["Banquet", "Server"],
    "dashboard_exclude_exact_names": ["Bar", "Patio", "Banquet", "Takeout"],
    "management_score_thresholds": {
        "check_average": {"neutral": 11.5, "strong": 18.5, "lower_is_better": False},
        "wine_pct": {"neutral": 0.041, "strong": 0.057, "lower_is_better": False},
        "rate_of_sale_by_guest_count": {
            "neutral": 0.019,
            "strong": 0.027,
            "lower_is_better": True,
        },
        "average_ticket_time_seconds": {
            "neutral": 720.0,
            "strong": 1140.0,
            "lower_is_better": True,
        },
    },
    "management_peer_score_thresholds": {
        "check_average": {"neutral": 11.0, "strong": 16.5, "lower_is_better": False},
        "wine_pct": {"neutral": 0.041, "strong": 0.058, "lower_is_better": False},
        "rate_of_sale_by_guest_count": {
            "neutral": 0.019,
            "strong": 0.028,
            "lower_is_better": True,
        },
        "average_ticket_time_seconds": {
            "neutral": 840.0,
            "strong": 1080.0,
            "lower_is_better": True,
        },
    },
    "management_peer_reference": {
        "prior_full_weeks": 4,
        "min_prior_full_weeks": 3,
        "min_distinct_peers_per_week": 5,
        "min_peer_server_weeks": 20,
        "statistic": "median",
        "leave_one_person_out": True,
    },
    "management_signal_persistence": {
        "qualified_weeks": 2,
        "require_recurring_driver": True,
        "require_leave_one_active_day_stability": True,
    },
    "management_threshold_calibration": {
        "method": "r7-absolute-deviation",
        "neutral_quantile": 0.75,
        "strong_quantile": 0.9,
        "calibration_start": "2026-04-28",
        "calibration_end": "2026-07-19",
        "movement_observation_count": 338,
        "peer_observation_count": 306,
        "version": "2026.07-v3",
    },
    "management_min_entity_baseline_weeks": 2,
    "management_materiality": {
        "sales_pct": 0.05,
        "guest_pct": 0.05,
        "check_average": 2.5,
        "wine_pct": 0.005,
        "rate": 0.005,
        "ticket_minutes": 2.5,
    },
    "public_name_aliases": {
        "Bar 1 Bar 1": "Bar",
        "Bar Server": "Bar",
        "BarPatio Bartender Patio": "Patio",
    },
    "public_exclude_name_contains": [
        "Banquet",
        "Takeout",
        "Server Server",
        "Jonathan Josephs",
        "Sean Kelly",
        "Bryan Pettit",
        "Christina Rivera",
        "Paul Sorensen",
        "Paula Friedrich",
        "Cicily McFadden",
        "AGM",
        "manager",
    ],
}

POSITIVE_INTEGER_FIELDS = {
    "public_min_guest_count",
    "master_min_guest_count_for_rankings",
    "dashboard_min_guest_count_for_trends",
    "dashboard_min_active_days_for_trends",
    "dashboard_min_prior_full_weeks",
    "dashboard_min_prior_guest_count",
    "dashboard_baseline_full_weeks",
    "dashboard_long_term_full_weeks",
    "dashboard_long_term_block_weeks",
    "dashboard_long_term_full_min_recent_guests",
    "dashboard_long_term_full_min_earlier_guests",
    "dashboard_long_term_developing_min_total_weeks",
    "dashboard_long_term_developing_min_recent_weeks",
    "dashboard_long_term_developing_min_earlier_weeks",
    "dashboard_long_term_developing_min_recent_guests",
    "dashboard_long_term_developing_min_earlier_guests",
    "management_min_entity_baseline_weeks",
}


class ConfigError(ValueError):
    """Raised when a configuration file violates the supported schema."""


def _object_without_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ConfigError(f"Duplicate configuration field: {key}.")
        result[key] = value
    return result


def _read_user_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_object_without_duplicate_keys,
            parse_constant=lambda value: (_ for _ in ()).throw(
                ConfigError(f"Non-finite JSON number is not allowed: {value}.")
            ),
        )
    except ConfigError:
        raise
    except json.JSONDecodeError as exc:
        raise ConfigError(
            f"Invalid configuration JSON at line {exc.lineno}, column {exc.colno}: "
            f"{exc.msg}."
        ) from exc
    except (OSError, UnicodeError) as exc:
        raise ConfigError(f"Could not read configuration {path}: {exc}.") from exc
    if not isinstance(payload, dict):
        raise ConfigError("Configuration root must be a JSON object.")
    return payload


def _reject_unknown_fields(
    payload: dict[str, Any], allowed: set[str], path: str
) -> None:
    unknown = sorted(set(payload) - allowed)
    if unknown:
        field = f"{path}.{unknown[0]}" if path else unknown[0]
        raise ConfigError(f"Unknown configuration field: {field}.")


def _require_mapping(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ConfigError(f"{path} must be a JSON object.")
    return value


def _require_finite_number(value: Any, path: str, *, positive: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ConfigError(f"{path} must be a finite number.")
    number = float(value)
    if not math.isfinite(number):
        raise ConfigError(f"{path} must be a finite number.")
    if positive and number <= 0:
        raise ConfigError(f"{path} must be greater than zero.")
    return number


def _require_positive_integer(value: Any, path: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ConfigError(f"{path} must be a positive integer.")
    return value


def _require_string(value: Any, path: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"{path} must be a non-empty string.")
    return value.strip()


def _validate_string_list(value: Any, path: str) -> list[str]:
    if not isinstance(value, list):
        raise ConfigError(f"{path} must be a JSON array of strings.")
    normalized = [_require_string(item, f"{path}[{index}]") for index, item in enumerate(value)]
    folded = [item.casefold() for item in normalized]
    if len(folded) != len(set(folded)):
        raise ConfigError(f"{path} contains duplicate values.")
    return normalized


def _deep_merge(base: dict[str, Any], updates: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in updates.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def _validate_effective_locations(locations_value: Any) -> None:
    locations = _require_mapping(locations_value, "locations")
    if not locations:
        raise ConfigError("locations must define at least one location.")
    normalized_names: list[str] = []
    short_codes: list[str] = []
    for location, settings_value in locations.items():
        normalized_names.append(_require_string(location, "locations key").casefold())
        settings = _require_mapping(settings_value, f"locations.{location}")
        _reject_unknown_fields(settings, {"short_code"}, f"locations.{location}")
        if "short_code" not in settings:
            raise ConfigError(f"locations.{location}.short_code is required.")
        short_codes.append(
            _require_string(
                settings["short_code"], f"locations.{location}.short_code"
            ).casefold()
        )
    if len(normalized_names) != len(set(normalized_names)):
        raise ConfigError("locations names must be unique ignoring case.")
    if len(short_codes) != len(set(short_codes)):
        raise ConfigError("locations short_code values must be unique.")


def validate_config_payload(payload: dict[str, Any]) -> dict[str, Any]:
    _reject_unknown_fields(payload, set(DEFAULT_CONFIG), "")

    if "locations" in payload:
        locations = _require_mapping(payload["locations"], "locations")
        if not locations:
            raise ConfigError("locations must define at least one location.")
        for location, settings_value in locations.items():
            _require_string(location, "locations key")
            settings = _require_mapping(settings_value, f"locations.{location}")
            _reject_unknown_fields(settings, {"short_code"}, f"locations.{location}")
            if "short_code" not in settings:
                raise ConfigError(f"locations.{location}.short_code is required.")
            _require_string(
                settings["short_code"], f"locations.{location}.short_code"
            )

    for field in POSITIVE_INTEGER_FIELDS & set(payload):
        _require_positive_integer(payload[field], field)

    for field in (
        "dashboard_exclude_name_contains",
        "dashboard_exclude_exact_names",
        "public_exclude_name_contains",
    ):
        if field in payload:
            _validate_string_list(payload[field], field)

    if "public_name_aliases" in payload:
        aliases = _require_mapping(payload["public_name_aliases"], "public_name_aliases")
        for source, target in aliases.items():
            _require_string(source, "public_name_aliases key")
            _require_string(target, f"public_name_aliases.{source}")

    for threshold_family in (
        "management_score_thresholds",
        "management_peer_score_thresholds",
    ):
        if threshold_family not in payload:
            continue
        thresholds = _require_mapping(payload[threshold_family], threshold_family)
        allowed_metrics = set(DEFAULT_CONFIG[threshold_family])
        _reject_unknown_fields(thresholds, allowed_metrics, threshold_family)
        for metric, settings_value in thresholds.items():
            settings = _require_mapping(
                settings_value, f"{threshold_family}.{metric}"
            )
            _reject_unknown_fields(
                settings,
                {"neutral", "strong", "lower_is_better"},
                f"{threshold_family}.{metric}",
            )
            effective_settings = _deep_merge(
                DEFAULT_CONFIG[threshold_family][metric], settings
            )
            neutral = _require_finite_number(
                effective_settings["neutral"],
                f"{threshold_family}.{metric}.neutral",
                positive=True,
            )
            strong = _require_finite_number(
                effective_settings["strong"],
                f"{threshold_family}.{metric}.strong",
                positive=True,
            )
            if strong <= neutral:
                raise ConfigError(
                    f"{threshold_family}.{metric}.strong must be greater than neutral."
                )
            if metric in {"wine_pct", "rate_of_sale_by_guest_count"} and strong > 1:
                raise ConfigError(
                    f"{threshold_family}.{metric}.strong cannot exceed 1."
                )
            if not isinstance(effective_settings["lower_is_better"], bool):
                raise ConfigError(
                    f"{threshold_family}.{metric}.lower_is_better must be boolean."
                )

    if "management_peer_reference" in payload:
        peer = _require_mapping(
            payload["management_peer_reference"], "management_peer_reference"
        )
        allowed = set(DEFAULT_CONFIG["management_peer_reference"])
        _reject_unknown_fields(peer, allowed, "management_peer_reference")
        effective_peer = _deep_merge(DEFAULT_CONFIG["management_peer_reference"], peer)
        for field in (
            "prior_full_weeks",
            "min_prior_full_weeks",
            "min_distinct_peers_per_week",
            "min_peer_server_weeks",
        ):
            _require_positive_integer(
                effective_peer[field], f"management_peer_reference.{field}"
            )
        if (
            effective_peer["min_prior_full_weeks"]
            > effective_peer["prior_full_weeks"]
        ):
            raise ConfigError(
                "management_peer_reference.min_prior_full_weeks cannot exceed "
                "prior_full_weeks."
            )
        if effective_peer["statistic"] != "median":
            raise ConfigError(
                "management_peer_reference.statistic must be 'median'."
            )
        if not isinstance(effective_peer["leave_one_person_out"], bool):
            raise ConfigError(
                "management_peer_reference.leave_one_person_out must be boolean."
            )
        if effective_peer["leave_one_person_out"] is not True:
            raise ConfigError(
                "management_peer_reference.leave_one_person_out must be true for "
                "methodology 2026.07-v3."
            )

    if "management_signal_persistence" in payload:
        persistence = _require_mapping(
            payload["management_signal_persistence"],
            "management_signal_persistence",
        )
        allowed = set(DEFAULT_CONFIG["management_signal_persistence"])
        _reject_unknown_fields(
            persistence, allowed, "management_signal_persistence"
        )
        effective_persistence = _deep_merge(
            DEFAULT_CONFIG["management_signal_persistence"], persistence
        )
        _require_positive_integer(
            effective_persistence["qualified_weeks"],
            "management_signal_persistence.qualified_weeks",
        )
        if effective_persistence["qualified_weeks"] != 2:
            raise ConfigError(
                "management_signal_persistence.qualified_weeks must be 2 for "
                "methodology 2026.07-v3."
            )
        for field in (
            "require_recurring_driver",
            "require_leave_one_active_day_stability",
        ):
            if not isinstance(effective_persistence[field], bool):
                raise ConfigError(
                    f"management_signal_persistence.{field} must be boolean."
                )
            if effective_persistence[field] is not True:
                raise ConfigError(
                    f"management_signal_persistence.{field} must be true for "
                    "methodology 2026.07-v3."
                )

    if "management_threshold_calibration" in payload:
        calibration = _require_mapping(
            payload["management_threshold_calibration"],
            "management_threshold_calibration",
        )
        allowed = set(DEFAULT_CONFIG["management_threshold_calibration"])
        _reject_unknown_fields(
            calibration, allowed, "management_threshold_calibration"
        )
        effective_calibration = _deep_merge(
            DEFAULT_CONFIG["management_threshold_calibration"], calibration
        )
        if effective_calibration["method"] != "r7-absolute-deviation":
            raise ConfigError(
                "management_threshold_calibration.method must be "
                "'r7-absolute-deviation'."
            )
        neutral_q = _require_finite_number(
            effective_calibration["neutral_quantile"],
            "management_threshold_calibration.neutral_quantile",
            positive=True,
        )
        strong_q = _require_finite_number(
            effective_calibration["strong_quantile"],
            "management_threshold_calibration.strong_quantile",
            positive=True,
        )
        if not (neutral_q < strong_q < 1):
            raise ConfigError(
                "management_threshold_calibration quantiles must satisfy "
                "0 < neutral_quantile < strong_quantile < 1."
            )
        for count_field in (
            "movement_observation_count",
            "peer_observation_count",
        ):
            observation_count = effective_calibration[count_field]
            if (
                isinstance(observation_count, bool)
                or not isinstance(observation_count, int)
                or observation_count < 0
            ):
                raise ConfigError(
                    f"management_threshold_calibration.{count_field} must be a "
                    "non-negative integer."
                )
        for field in ("calibration_start", "calibration_end"):
            value = _require_string(
                effective_calibration[field],
                f"management_threshold_calibration.{field}",
            )
            try:
                date.fromisoformat(value)
            except ValueError as exc:
                raise ConfigError(
                    f"management_threshold_calibration.{field} must be YYYY-MM-DD."
                ) from exc
        _require_string(
            effective_calibration["version"],
            "management_threshold_calibration.version",
        )

    if "management_materiality" in payload:
        materiality = _require_mapping(
            payload["management_materiality"], "management_materiality"
        )
        allowed_materiality = set(DEFAULT_CONFIG["management_materiality"])
        _reject_unknown_fields(
            materiality, allowed_materiality, "management_materiality"
        )
        for field, value in materiality.items():
            number = _require_finite_number(
                value, f"management_materiality.{field}", positive=True
            )
            if field in {"sales_pct", "guest_pct", "wine_pct", "rate"} and number > 1:
                raise ConfigError(
                    f"management_materiality.{field} cannot exceed 1."
                )

    merged = _deep_merge(DEFAULT_CONFIG, payload)
    _validate_effective_locations(merged["locations"])
    if (
        merged["dashboard_long_term_full_weeks"]
        != merged["dashboard_long_term_block_weeks"] * 2
    ):
        raise ConfigError(
            "dashboard_long_term_full_weeks must equal two times "
            "dashboard_long_term_block_weeks."
        )
    if (
        merged["dashboard_long_term_developing_min_total_weeks"]
        > merged["dashboard_long_term_full_weeks"]
    ):
        raise ConfigError(
            "dashboard_long_term_developing_min_total_weeks cannot exceed "
            "dashboard_long_term_full_weeks."
        )
    if (
        merged["dashboard_long_term_developing_min_recent_weeks"]
        > merged["dashboard_long_term_block_weeks"]
        or merged["dashboard_long_term_developing_min_earlier_weeks"]
        > merged["dashboard_long_term_block_weeks"]
    ):
        raise ConfigError(
            "Developing recent/earlier week minimums cannot exceed "
            "dashboard_long_term_block_weeks."
        )
    return merged


def load_config(path: Path) -> dict[str, Any]:
    return validate_config_payload(_read_user_config(path))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Validate Red Onion weekly metrics configuration."
    )
    parser.add_argument("--config", required=True, help="Path to configuration JSON.")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    path = Path(args.config).resolve()
    try:
        config = load_config(path)
    except Exception as exc:
        raise SystemExit(f"ERROR: {exc}") from exc
    print(
        "Configuration valid: "
        f"{len(config['locations'])} location(s), {len(config)} supported top-level fields."
    )


if __name__ == "__main__":
    main()
