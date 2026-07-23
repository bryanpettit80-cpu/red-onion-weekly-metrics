from __future__ import annotations

import argparse
import copy
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
        "check_average": {"neutral": 2.5, "strong": 5.0, "lower_is_better": False},
        "wine_pct": {"neutral": 0.005, "strong": 0.01, "lower_is_better": False},
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
    },
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

    if "management_score_thresholds" in payload:
        thresholds = _require_mapping(
            payload["management_score_thresholds"], "management_score_thresholds"
        )
        allowed_metrics = set(DEFAULT_CONFIG["management_score_thresholds"])
        _reject_unknown_fields(thresholds, allowed_metrics, "management_score_thresholds")
        for metric, settings_value in thresholds.items():
            settings = _require_mapping(
                settings_value, f"management_score_thresholds.{metric}"
            )
            _reject_unknown_fields(
                settings,
                {"neutral", "strong", "lower_is_better"},
                f"management_score_thresholds.{metric}",
            )
            effective_settings = _deep_merge(
                DEFAULT_CONFIG["management_score_thresholds"][metric], settings
            )
            neutral = _require_finite_number(
                effective_settings["neutral"],
                f"management_score_thresholds.{metric}.neutral",
                positive=True,
            )
            strong = _require_finite_number(
                effective_settings["strong"],
                f"management_score_thresholds.{metric}.strong",
                positive=True,
            )
            if strong < neutral:
                raise ConfigError(
                    f"management_score_thresholds.{metric}.strong must be at least neutral."
                )
            if metric in {"wine_pct", "rate_of_sale_by_guest_count"} and strong > 1:
                raise ConfigError(
                    f"management_score_thresholds.{metric}.strong cannot exceed 1."
                )
            if not isinstance(effective_settings["lower_is_better"], bool):
                raise ConfigError(
                    f"management_score_thresholds.{metric}.lower_is_better must be boolean."
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
