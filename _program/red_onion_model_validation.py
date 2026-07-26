"""Maintainer-only calibration and backtesting for management signals.

The ordinary weekly workflow imports none of this module.  It is deliberately
read-only: report directories are parsed and validated in memory, and the CLI
prints one JSON document without creating or changing files.
"""

from __future__ import annotations

import argparse
import copy
from dataclasses import dataclass
from datetime import date
import json
from pathlib import Path
import sys
from typing import Any, Iterable, Mapping, Sequence

from red_onion_config import DEFAULT_CONFIG, load_config
from red_onion_fairness import CandidatePolarity, calibrate_hybrid_band
import red_onion_weekly_metrics as weekly


METRIC_SPECS: dict[str, dict[str, float]] = {
    "check_average": {
        "business_neutral": 2.5,
        "business_strong": 5.0,
        "increment": 0.5,
    },
    "wine_pct": {
        "business_neutral": 0.005,
        "business_strong": 0.01,
        "increment": 0.001,
    },
}

QUALIFIED_EVIDENCE_STATUSES = frozenset({"Eligible", "Stable", "Sensitive"})
PROMPT_ACTIONS = frozenset({"Coaching Prompt", "Recognition Prompt"})
REVIEW_OR_ACTIONS = frozenset(
    {"Context Review", "Coaching Prompt", "Recognition Prompt"}
)
POLARITIES = frozenset(
    {CandidatePolarity.POSITIVE.value, CandidatePolarity.NEGATIVE.value}
)


@dataclass(frozen=True)
class ReportDataset:
    """Validated in-memory inventory from one or more report directories."""

    records: tuple[weekly.MetricRecord, ...]
    source_report_count: int
    duplicate_report_count: int
    business_dates: tuple[date, ...]


def _coerce_date(value: date | str | None, label: str) -> date | None:
    if value is None or isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value))
    except ValueError as exc:
        raise ValueError(f"{label} must be YYYY-MM-DD") from exc


def load_validated_report_directories(
    report_directories: Sequence[Path | str],
    config: dict[str, Any],
) -> ReportDataset:
    """Parse, deduplicate, and reconcile report directories without writing."""

    if not report_directories:
        raise ValueError("At least one report directory is required.")
    paths_by_key: dict[str, Path] = {}
    for raw_directory in report_directories:
        directory = Path(raw_directory).expanduser()
        if not directory.exists() or not directory.is_dir():
            raise ValueError(f"Report directory does not exist: {directory}")
        for path in weekly.find_daily_report_paths(directory, recursive=True):
            key = str(path.resolve()).casefold()
            paths_by_key.setdefault(key, path)
    paths = [paths_by_key[key] for key in sorted(paths_by_key)]
    if not paths:
        raise FileNotFoundError(
            "No supported Daily Report workbooks were found in the supplied directories."
        )

    parsed = weekly.read_reports_by_path(paths, config)
    resolution = weekly.resolve_report_duplicates(parsed)
    records = weekly.flatten_report_records(resolution.records_by_path)
    weekly.validate_daily_location_reconciliation(records, config)
    return ReportDataset(
        records=tuple(records),
        source_report_count=len(paths),
        duplicate_report_count=len(resolution.duplicate_paths),
        business_dates=resolution.business_dates,
    )


def _calibration_settings(config: Mapping[str, Any]) -> tuple[float, float, str]:
    configured = config.get("management_threshold_calibration", {})
    neutral_quantile = float(configured.get("neutral_quantile", 0.75))
    strong_quantile = float(configured.get("strong_quantile", 0.90))
    version = str(configured.get("version", "maintainer-calibration")).strip()
    if not version:
        raise ValueError("management_threshold_calibration.version cannot be blank")
    return neutral_quantile, strong_quantile, version


def collect_qualified_deviations(
    weekly_server_rows: Sequence[dict[str, Any]],
    weekly_location_rows: Sequence[dict[str, Any]],
    config: dict[str, Any],
    *,
    start: date | str | None = None,
    end: date | str | None = None,
) -> dict[str, dict[str, list[float]]]:
    """Collect action-eligible movement and peer deviations by metric."""

    start_date = _coerce_date(start, "start")
    end_date = _coerce_date(end, "end")
    if start_date and end_date and start_date > end_date:
        raise ValueError("start cannot be after end")
    full_by_location, _ = weekly.full_week_ends_by_location(
        list(weekly_location_rows)
    )
    deviations = {
        "movement": {metric: [] for metric in METRIC_SPECS},
        "peer": {metric: [] for metric in METRIC_SPECS},
    }
    for row in sorted(
        weekly_server_rows,
        key=lambda item: (
            item["week_end"],
            str(item.get("location") or ""),
            str(item.get("raw_user_name") or ""),
        ),
    ):
        week_end = row["week_end"]
        if (
            (start_date is not None and week_end < start_date)
            or (end_date is not None and week_end > end_date)
            or weekly.dashboard_excluded(row, config)
        ):
            continue
        evaluation = weekly.evaluate_server_week_signal(
            row,
            list(weekly_server_rows),
            full_by_location,
            config,
        )
        base_qualified = bool(
            evaluation["full_latest"]
            and evaluation["current_sample_eligible"]
            and evaluation["self_history_eligible"]
        )
        if not base_qualified:
            continue
        for metric in METRIC_SPECS:
            movement_value = evaluation["changes"].get(metric)
            if movement_value is not None:
                deviations["movement"][metric].append(float(movement_value))
            reference = evaluation["peer_references"].get(metric)
            peer_value = evaluation["peer_changes"].get(metric)
            if (
                reference is not None
                and reference.sufficient
                and peer_value is not None
            ):
                deviations["peer"][metric].append(float(peer_value))
    return deviations


def calibrate_deviation_sets(
    movement_deviations: Mapping[str, Sequence[float]],
    peer_deviations: Mapping[str, Sequence[float]],
    config: dict[str, Any],
    *,
    calibration_start: date | str,
    calibration_end: date | str,
) -> dict[str, Any]:
    """Return frozen config families plus aggregate, anonymized diagnostics."""

    start_date = _coerce_date(calibration_start, "calibration_start")
    end_date = _coerce_date(calibration_end, "calibration_end")
    if start_date is None or end_date is None:
        raise ValueError("calibration_start and calibration_end are required")
    if start_date > end_date:
        raise ValueError("calibration_start cannot be after calibration_end")
    neutral_quantile, strong_quantile, version = _calibration_settings(config)

    family_inputs = {
        "movement": movement_deviations,
        "peer": peer_deviations,
    }
    config_keys = {
        "movement": "management_score_thresholds",
        "peer": "management_peer_score_thresholds",
    }
    calibrated_families: dict[str, dict[str, dict[str, Any]]] = {}
    family_diagnostics: dict[str, dict[str, Any]] = {}
    family_observation_counts: dict[str, int] = {}
    for comparator, inputs in family_inputs.items():
        thresholds: dict[str, dict[str, Any]] = {}
        metric_diagnostics: dict[str, dict[str, Any]] = {}
        observation_count = 0
        for metric, spec in METRIC_SPECS.items():
            values = list(inputs.get(metric, ()))
            if not values:
                raise ValueError(
                    f"No qualified {comparator} deviations were available for {metric}."
                )
            result = calibrate_hybrid_band(
                values,
                business_neutral=spec["business_neutral"],
                business_strong=spec["business_strong"],
                increment=spec["increment"],
                neutral_quantile=neutral_quantile,
                strong_quantile=strong_quantile,
            )
            lower_is_better = bool(
                config.get(config_keys[comparator], {})
                .get(metric, DEFAULT_CONFIG[config_keys[comparator]][metric])
                .get("lower_is_better", False)
            )
            thresholds[metric] = {
                "neutral": result.band.neutral,
                "strong": result.band.strong,
                "lower_is_better": lower_is_better,
            }
            metric_diagnostics[metric] = {
                "observation_count": result.observation_count,
                "empirical_neutral": result.empirical_neutral,
                "empirical_strong": result.empirical_strong,
                "business_neutral": result.business_neutral,
                "business_strong": result.business_strong,
                "increment": result.increment,
                "calibrated_neutral": result.band.neutral,
                "calibrated_strong": result.band.strong,
            }
            observation_count += result.observation_count
        calibrated_families[config_keys[comparator]] = thresholds
        family_diagnostics[comparator] = {
            "observation_count": observation_count,
            "observation_unit": "qualified metric deviation",
            "metrics": metric_diagnostics,
        }
        family_observation_counts[comparator] = observation_count

    calibration_metadata = {
        "method": "r7-absolute-deviation",
        "neutral_quantile": neutral_quantile,
        "strong_quantile": strong_quantile,
        "calibration_start": start_date.isoformat(),
        "calibration_end": end_date.isoformat(),
        "movement_observation_count": family_observation_counts["movement"],
        "peer_observation_count": family_observation_counts["peer"],
        "version": version,
    }
    return {
        "config_fragment": {
            **calibrated_families,
            "management_threshold_calibration": calibration_metadata,
        },
        "diagnostics": {
            "method": "R-7 absolute qualified deviations with business minima",
            "calibration_start": start_date.isoformat(),
            "calibration_end": end_date.isoformat(),
            "families": family_diagnostics,
        },
    }


def calibrate_weekly_rows(
    weekly_server_rows: Sequence[dict[str, Any]],
    weekly_location_rows: Sequence[dict[str, Any]],
    config: dict[str, Any],
    *,
    calibration_start: date | str,
    calibration_end: date | str,
) -> dict[str, Any]:
    """Calibrate both comparator families from validated weekly rollups."""

    deviations = collect_qualified_deviations(
        weekly_server_rows,
        weekly_location_rows,
        config,
        start=calibration_start,
        end=calibration_end,
    )
    return calibrate_deviation_sets(
        deviations["movement"],
        deviations["peer"],
        config,
        calibration_start=calibration_start,
        calibration_end=calibration_end,
    )


def summarize_backtest_observations(
    observations: Iterable[Mapping[str, Any]],
) -> dict[str, Any]:
    """Summarize decision behavior without returning person-level identifiers."""

    selected = [dict(item) for item in observations]
    qualified = [item for item in selected if bool(item.get("qualified"))]
    review_or_action = [
        item for item in qualified if item.get("action") in REVIEW_OR_ACTIONS
    ]
    store_week_groups: dict[tuple[str, date], list[dict[str, Any]]] = {}
    for item in qualified:
        key = (str(item.get("location") or ""), item["week_end"])
        store_week_groups.setdefault(key, []).append(item)
    store_week_rates = [
        sum(item.get("action") in REVIEW_OR_ACTIONS for item in rows) / len(rows)
        for rows in store_week_groups.values()
        if rows
    ]

    prompts = [item for item in qualified if item.get("action") in PROMPT_ACTIONS]
    stable_prompts = [
        item
        for item in prompts
        if str(item.get("stability_result") or "").startswith(
            "Stable under every active-day removal"
        )
    ]

    by_entity: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for item in qualified:
        entity_key = (
            str(item.get("location") or ""),
            str(item.get("_entity_key") or ""),
        )
        by_entity.setdefault(entity_key, []).append(item)
    consecutive_pairs = 0
    consecutive_reversals = 0
    for rows in by_entity.values():
        rows.sort(key=lambda item: item["week_end"])
        for earlier, later in zip(rows, rows[1:]):
            if (later["week_end"] - earlier["week_end"]).days != 7:
                continue
            earlier_polarity = str(earlier.get("candidate_polarity") or "")
            later_polarity = str(later.get("candidate_polarity") or "")
            if earlier_polarity not in POLARITIES or later_polarity not in POLARITIES:
                continue
            consecutive_pairs += 1
            consecutive_reversals += earlier_polarity != later_polarity

    qualified_count = len(qualified)
    prompt_count = len(prompts)
    return {
        "evaluated_person_weeks": len(selected),
        "qualified_person_weeks": qualified_count,
        "review_or_action_person_weeks": len(review_or_action),
        "overall_review_action_rate": (
            len(review_or_action) / qualified_count if qualified_count else None
        ),
        "store_week_group_count": len(store_week_groups),
        "maximum_store_week_review_action_rate": (
            max(store_week_rates) if store_week_rates else None
        ),
        "prompt_count": prompt_count,
        "stable_prompt_count": len(stable_prompts),
        "prompt_stability_rate": (
            len(stable_prompts) / prompt_count if prompt_count else None
        ),
        "consecutive_candidate_pairs": consecutive_pairs,
        "consecutive_reversals": consecutive_reversals,
        "consecutive_reversal_rate": (
            consecutive_reversals / consecutive_pairs
            if consecutive_pairs
            else None
        ),
    }


def backtest_weekly_rows(
    weekly_server_rows: Sequence[dict[str, Any]],
    weekly_location_rows: Sequence[dict[str, Any]],
    config: dict[str, Any],
    *,
    start: date | str | None = None,
    end: date | str | None = None,
) -> dict[str, Any]:
    """Replay each historical week and return aggregate model diagnostics."""

    start_date = _coerce_date(start, "start")
    end_date = _coerce_date(end, "end")
    if start_date and end_date and start_date > end_date:
        raise ValueError("start cannot be after end")
    week_ends = sorted(
        {
            row["week_end"]
            for row in weekly_server_rows
            if (start_date is None or row["week_end"] >= start_date)
            and (end_date is None or row["week_end"] <= end_date)
        }
    )
    observations: list[dict[str, Any]] = []
    for week_end in week_ends:
        server_history = [
            row for row in weekly_server_rows if row["week_end"] <= week_end
        ]
        location_history = [
            row for row in weekly_location_rows if row["week_end"] <= week_end
        ]
        outputs = weekly.management_server_rows(
            server_history,
            location_history,
            [],
            {},
            config,
        )
        for row in outputs:
            observations.append(
                {
                    "_entity_key": str(row.get("raw_user_name") or ""),
                    "location": str(row.get("location") or ""),
                    "week_end": row["week_end"],
                    "qualified": row.get("confidence")
                    in QUALIFIED_EVIDENCE_STATUSES,
                    "action": str(row.get("action") or "Monitor"),
                    "candidate_polarity": str(
                        row.get("candidate_polarity") or ""
                    ),
                    "stability_result": str(
                        row.get("stability_result") or ""
                    ),
                }
            )
    diagnostics = summarize_backtest_observations(observations)
    diagnostics["weeks_replayed"] = len(week_ends)
    return diagnostics


def _config_with_fragment(
    config: dict[str, Any],
    fragment: Mapping[str, Any],
) -> dict[str, Any]:
    calibrated = copy.deepcopy(config)
    for key in (
        "management_score_thresholds",
        "management_peer_score_thresholds",
        "management_threshold_calibration",
    ):
        calibrated[key] = copy.deepcopy(fragment[key])
    return calibrated


def validate_report_directories(
    report_directories: Sequence[Path | str],
    config: dict[str, Any],
    *,
    calibration_start: date | str,
    calibration_end: date | str,
) -> dict[str, Any]:
    """Run read-only calibration and backtesting over validated reports."""

    start_date = _coerce_date(calibration_start, "calibration_start")
    end_date = _coerce_date(calibration_end, "calibration_end")
    if start_date is None or end_date is None:
        raise ValueError("calibration_start and calibration_end are required")
    dataset = load_validated_report_directories(report_directories, config)
    selected_records = [
        record
        for record in dataset.records
        if start_date <= record.report_date <= end_date
    ]
    if not selected_records:
        raise ValueError("No report rows fall inside the calibration window.")
    server_rows, location_rows = weekly.weekly_rollups(selected_records)
    calibrated = calibrate_weekly_rows(
        server_rows,
        location_rows,
        config,
        calibration_start=start_date,
        calibration_end=end_date,
    )
    calibrated_config = _config_with_fragment(
        config, calibrated["config_fragment"]
    )
    backtest = backtest_weekly_rows(
        server_rows,
        location_rows,
        calibrated_config,
        start=start_date,
        end=end_date,
    )
    observed_dates = sorted({record.report_date for record in selected_records})
    return {
        "status": "ok",
        "config_fragment": calibrated["config_fragment"],
        "diagnostics": {
            "source": {
                "source_report_count": dataset.source_report_count,
                "duplicate_report_count": dataset.duplicate_report_count,
                "business_date_count": len(observed_dates),
                "observed_start": observed_dates[0].isoformat(),
                "observed_end": observed_dates[-1].isoformat(),
                "location_count": len(
                    {record.location for record in selected_records}
                ),
            },
            "calibration": calibrated["diagnostics"],
            "backtest": backtest,
        },
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Read-only maintainer calibration and backtest; prints JSON only."
        )
    )
    parser.add_argument(
        "report_directories",
        nargs="+",
        type=Path,
        help="One or more directories containing Daily Report workbooks.",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=Path(__file__).with_name("red_onion_config.json"),
    )
    parser.add_argument("--start", help="Calibration start date (YYYY-MM-DD).")
    parser.add_argument("--end", help="Calibration end date (YYYY-MM-DD).")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        config = load_config(args.config)
        configured = config.get("management_threshold_calibration", {})
        start = args.start or configured.get("calibration_start")
        end = args.end or configured.get("calibration_end")
        if not start or not end:
            raise ValueError(
                "Supply --start and --end or configure the calibration window."
            )
        payload = validate_report_directories(
            args.report_directories,
            config,
            calibration_start=start,
            calibration_end=end,
        )
        exit_code = 0
    except Exception as exc:
        payload = {
            "status": "error",
            "error_type": type(exc).__name__,
            "message": str(exc),
        }
        exit_code = 1
    sys.stdout.write(json.dumps(payload, indent=2, sort_keys=True) + "\n")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
