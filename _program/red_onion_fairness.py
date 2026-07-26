"""Pure statistical and fairness helpers for Red Onion coaching signals.

This module deliberately contains no workbook, filesystem, pandas, or service
dependencies.  The reporting workflow can therefore test the decision rules
independently from source parsing and presentation.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from enum import Enum
import math
from statistics import median
from typing import Iterable, Mapping, Sequence


class RecentMovement(str, Enum):
    """Supported recent-movement labels."""

    UPWARD = "Upward"
    DOWNWARD = "Downward"
    STABLE = "Stable"
    NOT_EVALUATED = "Not Evaluated"


class PeerComparison(str, Enum):
    """Supported peer-comparison labels."""

    ABOVE = "Above Peer Reference"
    WITHIN = "Within Peer Range"
    BELOW = "Below Peer Reference"
    UNAVAILABLE = "Reference Unavailable"


class CandidatePolarity(str, Enum):
    """Direction of a qualified coaching-signal candidate."""

    POSITIVE = "positive"
    NEGATIVE = "negative"
    NONE = "none"


class PromptAction(str, Enum):
    """Human-review action produced by the persistence gate."""

    MONITOR = "Monitor"
    CONTEXT_REVIEW = "Context Review"
    COACHING_PROMPT = "Coaching Prompt"
    RECOGNITION_PROMPT = "Recognition Prompt"


@dataclass(frozen=True)
class MetricBand:
    """Neutral and strong absolute movement thresholds for one metric."""

    neutral: float
    strong: float

    def __post_init__(self) -> None:
        neutral = _finite_float(self.neutral, "neutral")
        strong = _finite_float(self.strong, "strong")
        if neutral < 0:
            raise ValueError("neutral must be non-negative")
        if strong <= neutral:
            raise ValueError("strong must be greater than neutral")


@dataclass(frozen=True)
class HybridCalibration:
    """Frozen result of hybrid business-minimum and empirical calibration."""

    band: MetricBand
    observation_count: int
    empirical_neutral: float
    empirical_strong: float
    business_neutral: float
    business_strong: float
    increment: float
    neutral_quantile: float
    strong_quantile: float
    quantile_method: str = "R-7"


@dataclass(frozen=True)
class PeerObservation:
    """One server-week metric value eligible for peer-reference evaluation."""

    person_id: str
    location: str
    week_end: date
    value: float | None
    qualified: bool = True
    excluded: bool = False


@dataclass(frozen=True)
class PeerReference:
    """Diagnostic result for a leave-one-person-out same-store reference."""

    value: float | None
    sufficient: bool
    reason: str
    requested_weeks: tuple[date, ...]
    usable_weeks: tuple[date, ...]
    peer_week_count: int
    distinct_peer_count: int
    peer_counts_by_week: tuple[tuple[date, int], ...]


@dataclass(frozen=True)
class CandidateAssessment:
    """Metric and label evidence for a positive or negative candidate."""

    polarity: CandidatePolarity
    eligible: bool
    composite_score: int
    agreeing_drivers: tuple[str, ...]
    opposing_drivers: tuple[str, ...]
    reason: str


@dataclass(frozen=True)
class StoreShockAssessment:
    """Whether a candidate materially differs from the current store movement."""

    guard_passed: bool
    common_store_shock: bool
    comparable_metrics: tuple[str, ...]
    differentiating_metrics: tuple[str, ...]
    relative_scores: tuple[tuple[str, int], ...]
    reason: str


@dataclass(frozen=True)
class WeeklyCandidateSignal:
    """Weekly inputs used by the symmetric two-week persistence gate."""

    week_end: date
    polarity: CandidatePolarity
    drivers: tuple[str, ...]
    qualified: bool
    leave_one_day_stable: bool
    store_shock_guard_passed: bool


@dataclass(frozen=True)
class PersistenceDecision:
    """Review action and audit detail from the two-week persistence gate."""

    action: PromptAction
    escalated: bool
    polarity: CandidatePolarity
    recurring_drivers: tuple[str, ...]
    reason: str


def _finite_float(value: object, name: str) -> float:
    if isinstance(value, bool):
        raise TypeError(f"{name} must be a finite number")
    try:
        result = float(value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise TypeError(f"{name} must be a finite number") from exc
    if not math.isfinite(result):
        raise ValueError(f"{name} must be finite")
    return result


def _optional_finite_float(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        result = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    return result if math.isfinite(result) else None


def _coerce_enum(value: object, enum_type: type[Enum], name: str) -> Enum:
    if isinstance(value, enum_type):
        return value
    normalized = str(value).strip().casefold()
    for item in enum_type:
        if str(item.value).casefold() == normalized:
            return item
    choices = ", ".join(repr(item.value) for item in enum_type)
    raise ValueError(f"{name} must be one of: {choices}")


def r7_quantile(values: Sequence[float], probability: float) -> float:
    """Return the R/NumPy default (R-7) linearly interpolated quantile."""

    probability_value = _finite_float(probability, "probability")
    if not 0 <= probability_value <= 1:
        raise ValueError("probability must be between 0 and 1")
    ordered = sorted(_finite_float(value, "quantile value") for value in values)
    if not ordered:
        raise ValueError("values must contain at least one observation")
    if len(ordered) == 1:
        return ordered[0]

    position = (len(ordered) - 1) * probability_value
    lower_index = math.floor(position)
    upper_index = math.ceil(position)
    if lower_index == upper_index:
        return ordered[lower_index]
    fraction = position - lower_index
    return ordered[lower_index] + fraction * (
        ordered[upper_index] - ordered[lower_index]
    )


def round_half_up_to_increment(value: float, increment: float) -> float:
    """Round to the nearest positive increment, resolving ties away from zero."""

    numeric_value = _finite_float(value, "value")
    numeric_increment = _finite_float(increment, "increment")
    if numeric_increment <= 0:
        raise ValueError("increment must be greater than zero")
    try:
        decimal_value = Decimal(str(numeric_value))
        decimal_increment = Decimal(str(numeric_increment))
        units = (decimal_value / decimal_increment).quantize(
            Decimal("1"),
            rounding=ROUND_HALF_UP,
        )
    except (InvalidOperation, ZeroDivisionError) as exc:
        raise ValueError("value and increment must support decimal rounding") from exc
    return float(units * decimal_increment)


def calibrate_hybrid_band(
    deviations: Sequence[float],
    *,
    business_neutral: float,
    business_strong: float,
    increment: float,
    neutral_quantile: float = 0.75,
    strong_quantile: float = 0.90,
) -> HybridCalibration:
    """Calibrate frozen neutral/strong bands from absolute qualified deviations."""

    business_neutral_value = _finite_float(
        business_neutral,
        "business_neutral",
    )
    business_strong_value = _finite_float(business_strong, "business_strong")
    increment_value = _finite_float(increment, "increment")
    neutral_probability = _finite_float(neutral_quantile, "neutral_quantile")
    strong_probability = _finite_float(strong_quantile, "strong_quantile")
    if business_neutral_value < 0:
        raise ValueError("business_neutral must be non-negative")
    if business_strong_value < business_neutral_value:
        raise ValueError("business_strong must not be below business_neutral")
    if increment_value <= 0:
        raise ValueError("increment must be greater than zero")
    if not 0 <= neutral_probability < strong_probability <= 1:
        raise ValueError(
            "quantiles must satisfy 0 <= neutral_quantile < strong_quantile <= 1"
        )

    absolute_deviations = [
        abs(_finite_float(value, "deviation")) for value in deviations
    ]
    if not absolute_deviations:
        raise ValueError("deviations must contain at least one observation")
    empirical_neutral = r7_quantile(
        absolute_deviations,
        neutral_probability,
    )
    empirical_strong = r7_quantile(
        absolute_deviations,
        strong_probability,
    )
    neutral = round_half_up_to_increment(
        max(business_neutral_value, empirical_neutral),
        increment_value,
    )
    strong = round_half_up_to_increment(
        max(business_strong_value, empirical_strong),
        increment_value,
    )
    if strong <= neutral:
        strong = round_half_up_to_increment(
            neutral + increment_value,
            increment_value,
        )

    return HybridCalibration(
        band=MetricBand(neutral=neutral, strong=strong),
        observation_count=len(absolute_deviations),
        empirical_neutral=empirical_neutral,
        empirical_strong=empirical_strong,
        business_neutral=business_neutral_value,
        business_strong=business_strong_value,
        increment=increment_value,
        neutral_quantile=neutral_probability,
        strong_quantile=strong_probability,
    )


def score_metric(
    deviation: float | None,
    band: MetricBand,
    *,
    higher_is_better: bool,
) -> int | None:
    """Score a current-minus-reference deviation from -2 through +2.

    ``None`` and non-finite inputs are unavailable rather than favorable zeroes.
    """

    numeric_deviation = _optional_finite_float(deviation)
    if numeric_deviation is None:
        return None
    if not isinstance(higher_is_better, bool):
        raise TypeError("higher_is_better must be a bool")
    benefit_deviation = (
        numeric_deviation if higher_is_better else -numeric_deviation
    )
    absolute_deviation = abs(benefit_deviation)
    if benefit_deviation == 0 or absolute_deviation < band.neutral:
        return 0
    magnitude = 2 if absolute_deviation >= band.strong else 1
    return magnitude if benefit_deviation > 0 else -magnitude


def leave_one_out_same_store_peer_reference(
    observations: Iterable[PeerObservation],
    *,
    focal_person_id: str,
    location: str,
    prior_week_ends: Sequence[date] | None = None,
    max_prior_weeks: int = 4,
    min_usable_weeks: int = 3,
    min_distinct_peers_per_week: int = 5,
    min_peer_weeks: int = 20,
) -> PeerReference:
    """Build a pooled same-store median without the focal person's observations."""

    if max_prior_weeks < 1:
        raise ValueError("max_prior_weeks must be at least 1")
    if min_usable_weeks < 1 or min_usable_weeks > max_prior_weeks:
        raise ValueError("min_usable_weeks must be between 1 and max_prior_weeks")
    if min_distinct_peers_per_week < 1:
        raise ValueError("min_distinct_peers_per_week must be at least 1")
    if min_peer_weeks < 1:
        raise ValueError("min_peer_weeks must be at least 1")

    all_observations = tuple(observations)
    if prior_week_ends is None:
        requested_weeks = sorted(
            {
                item.week_end
                for item in all_observations
                if isinstance(item.week_end, date)
            }
        )[-max_prior_weeks:]
    else:
        if any(not isinstance(week_end, date) for week_end in prior_week_ends):
            raise TypeError("prior_week_ends must contain date values")
        requested_weeks = sorted(set(prior_week_ends))[-max_prior_weeks:]
    requested_week_set = set(requested_weeks)

    focal_key = focal_person_id.strip().casefold()
    location_key = location.strip().casefold()
    grouped: dict[date, dict[str, list[float]]] = {
        week_end: {} for week_end in requested_weeks
    }
    for item in all_observations:
        if (
            item.week_end not in requested_week_set
            or item.location.strip().casefold() != location_key
            or item.person_id.strip().casefold() == focal_key
            or not item.qualified
            or item.excluded
        ):
            continue
        value = _optional_finite_float(item.value)
        if value is None:
            continue
        peer_key = item.person_id.strip().casefold()
        if not peer_key:
            continue
        grouped[item.week_end].setdefault(peer_key, []).append(value)

    peer_counts_by_week = tuple(
        (week_end, len(grouped[week_end])) for week_end in requested_weeks
    )
    usable_weeks = tuple(
        week_end
        for week_end, peer_count in peer_counts_by_week
        if peer_count >= min_distinct_peers_per_week
    )
    pooled_values: list[float] = []
    distinct_peers: set[str] = set()
    for week_end in usable_weeks:
        for peer_key, values in grouped[week_end].items():
            pooled_values.append(float(median(values)))
            distinct_peers.add(peer_key)

    if len(usable_weeks) < min_usable_weeks:
        sufficient = False
        reason = "insufficient_usable_weeks"
    elif len(pooled_values) < min_peer_weeks:
        sufficient = False
        reason = "insufficient_peer_weeks"
    else:
        sufficient = True
        reason = "available"
    reference_value = float(median(pooled_values)) if sufficient else None
    return PeerReference(
        value=reference_value,
        sufficient=sufficient,
        reason=reason,
        requested_weeks=tuple(requested_weeks),
        usable_weeks=usable_weeks,
        peer_week_count=len(pooled_values),
        distinct_peer_count=len(distinct_peers),
        peer_counts_by_week=peer_counts_by_week,
    )


def candidate_polarity(
    movement: RecentMovement | str,
    peer_comparison: PeerComparison | str,
) -> CandidatePolarity:
    """Return the candidate direction implied by movement and peer comparison."""

    movement_value = _coerce_enum(
        movement,
        RecentMovement,
        "movement",
    )
    comparison_value = _coerce_enum(
        peer_comparison,
        PeerComparison,
        "peer_comparison",
    )
    if (
        movement_value is RecentMovement.UPWARD
        and comparison_value is PeerComparison.ABOVE
    ):
        return CandidatePolarity.POSITIVE
    if (
        movement_value is RecentMovement.DOWNWARD
        and comparison_value is PeerComparison.BELOW
    ):
        return CandidatePolarity.NEGATIVE
    return CandidatePolarity.NONE


def classify_candidate(
    movement: RecentMovement | str,
    peer_comparison: PeerComparison | str,
    metric_scores: Mapping[str, int | None],
    *,
    minimum_composite: int = 3,
    minimum_agreeing_metrics: int = 2,
    required_metrics: Iterable[str] | None = None,
) -> CandidateAssessment:
    """Apply label alignment plus composite and agreeing-metric gates."""

    if minimum_composite < 1:
        raise ValueError("minimum_composite must be at least 1")
    if minimum_agreeing_metrics < 1:
        raise ValueError("minimum_agreeing_metrics must be at least 1")
    required = (
        tuple(dict.fromkeys(required_metrics))
        if required_metrics is not None
        else tuple(metric_scores)
    )
    unavailable = [
        metric
        for metric in required
        if metric not in metric_scores or metric_scores[metric] is None
    ]
    normalized_scores: dict[str, int] = {}
    for metric, score in metric_scores.items():
        if score is None:
            continue
        if isinstance(score, bool) or not isinstance(score, int) or not -2 <= score <= 2:
            raise ValueError(f"score for {metric!r} must be an integer from -2 to 2")
        normalized_scores[metric] = score
    composite = sum(normalized_scores.values())
    base_polarity = candidate_polarity(movement, peer_comparison)
    if unavailable:
        return CandidateAssessment(
            polarity=CandidatePolarity.NONE,
            eligible=False,
            composite_score=composite,
            agreeing_drivers=(),
            opposing_drivers=(),
            reason="metric_unavailable",
        )
    if base_polarity is CandidatePolarity.NONE:
        return CandidateAssessment(
            polarity=CandidatePolarity.NONE,
            eligible=False,
            composite_score=composite,
            agreeing_drivers=(),
            opposing_drivers=(),
            reason="movement_peer_not_aligned",
        )

    direction = 1 if base_polarity is CandidatePolarity.POSITIVE else -1
    agreeing = tuple(
        sorted(metric for metric, score in normalized_scores.items() if score * direction > 0)
    )
    opposing = tuple(
        sorted(metric for metric, score in normalized_scores.items() if score * direction < 0)
    )
    if composite * direction < minimum_composite:
        reason = "composite_below_gate"
        eligible = False
    elif len(agreeing) < minimum_agreeing_metrics:
        reason = "insufficient_agreeing_metrics"
        eligible = False
    else:
        reason = "candidate"
        eligible = True
    return CandidateAssessment(
        polarity=base_polarity if eligible else CandidatePolarity.NONE,
        eligible=eligible,
        composite_score=composite,
        agreeing_drivers=agreeing,
        opposing_drivers=opposing,
        reason=reason,
    )


def assess_common_store_shock(
    polarity: CandidatePolarity | str,
    focal_metric_changes: Mapping[str, float | None],
    peer_median_changes: Mapping[str, float | None],
    bands: Mapping[str, MetricBand],
    higher_is_better: Mapping[str, bool],
    *,
    candidate_drivers: Iterable[str] | None = None,
) -> StoreShockAssessment:
    """Require at least one candidate driver to differ materially from peers."""

    polarity_value = _coerce_enum(
        polarity,
        CandidatePolarity,
        "polarity",
    )
    if polarity_value is CandidatePolarity.NONE:
        return StoreShockAssessment(
            guard_passed=False,
            common_store_shock=False,
            comparable_metrics=(),
            differentiating_metrics=(),
            relative_scores=(),
            reason="no_candidate",
        )
    selected_metrics = (
        tuple(dict.fromkeys(candidate_drivers))
        if candidate_drivers is not None
        else tuple(focal_metric_changes)
    )
    relative_scores: list[tuple[str, int]] = []
    for metric in selected_metrics:
        focal_value = _optional_finite_float(focal_metric_changes.get(metric))
        peer_value = _optional_finite_float(peer_median_changes.get(metric))
        if (
            focal_value is None
            or peer_value is None
            or metric not in bands
            or metric not in higher_is_better
        ):
            continue
        relative_score = score_metric(
            focal_value - peer_value,
            bands[metric],
            higher_is_better=higher_is_better[metric],
        )
        if relative_score is not None:
            relative_scores.append((metric, relative_score))

    direction = 1 if polarity_value is CandidatePolarity.POSITIVE else -1
    differentiating = tuple(
        sorted(
            metric
            for metric, score in relative_scores
            if score * direction > 0
        )
    )
    comparable = tuple(sorted(metric for metric, _ in relative_scores))
    if differentiating:
        return StoreShockAssessment(
            guard_passed=True,
            common_store_shock=False,
            comparable_metrics=comparable,
            differentiating_metrics=differentiating,
            relative_scores=tuple(sorted(relative_scores)),
            reason="materially_different_from_store",
        )
    if comparable:
        return StoreShockAssessment(
            guard_passed=False,
            common_store_shock=True,
            comparable_metrics=comparable,
            differentiating_metrics=(),
            relative_scores=tuple(sorted(relative_scores)),
            reason="common_store_movement",
        )
    return StoreShockAssessment(
        guard_passed=False,
        common_store_shock=False,
        comparable_metrics=(),
        differentiating_metrics=(),
        relative_scores=(),
        reason="store_comparison_unavailable",
    )


def leave_one_day_stability(
    full_week_polarity: CandidatePolarity | str,
    leave_one_day_polarities: Sequence[CandidatePolarity | str],
) -> bool:
    """Return true only when every active-day removal preserves the candidate."""

    full_value = _coerce_enum(
        full_week_polarity,
        CandidatePolarity,
        "full_week_polarity",
    )
    if full_value is CandidatePolarity.NONE or not leave_one_day_polarities:
        return False
    return all(
        _coerce_enum(value, CandidatePolarity, "leave_one_day_polarity")
        is full_value
        for value in leave_one_day_polarities
    )


def evaluate_two_week_persistence(
    current: WeeklyCandidateSignal,
    previous: WeeklyCandidateSignal | None,
    *,
    expected_interval_days: int = 7,
) -> PersistenceDecision:
    """Return a prompt only after two stable, comparable, consecutive candidates."""

    if expected_interval_days < 1:
        raise ValueError("expected_interval_days must be at least 1")
    current_polarity = _coerce_enum(
        current.polarity,
        CandidatePolarity,
        "current.polarity",
    )
    if not current.qualified or current_polarity is CandidatePolarity.NONE:
        return PersistenceDecision(
            action=PromptAction.MONITOR,
            escalated=False,
            polarity=CandidatePolarity.NONE,
            recurring_drivers=(),
            reason="current_week_not_qualified_candidate",
        )
    context_result = PersistenceDecision(
        action=PromptAction.CONTEXT_REVIEW,
        escalated=False,
        polarity=current_polarity,
        recurring_drivers=(),
        reason="first_qualified_candidate",
    )
    if previous is None:
        return context_result
    if (current.week_end - previous.week_end).days != expected_interval_days:
        return PersistenceDecision(
            **{
                **context_result.__dict__,
                "reason": "nonconsecutive_week",
            }
        )
    previous_polarity = _coerce_enum(
        previous.polarity,
        CandidatePolarity,
        "previous.polarity",
    )
    if not previous.qualified or previous_polarity is CandidatePolarity.NONE:
        return PersistenceDecision(
            **{
                **context_result.__dict__,
                "reason": "previous_week_not_qualified_candidate",
            }
        )
    if previous_polarity is not current_polarity:
        return PersistenceDecision(
            **{
                **context_result.__dict__,
                "reason": "candidate_direction_changed",
            }
        )
    recurring_drivers = tuple(
        sorted(set(current.drivers).intersection(previous.drivers))
    )
    if not recurring_drivers:
        return PersistenceDecision(
            **{
                **context_result.__dict__,
                "reason": "no_recurring_driver",
            }
        )
    if not current.leave_one_day_stable or not previous.leave_one_day_stable:
        return PersistenceDecision(
            action=PromptAction.CONTEXT_REVIEW,
            escalated=False,
            polarity=current_polarity,
            recurring_drivers=recurring_drivers,
            reason="day_sensitive",
        )
    if (
        not current.store_shock_guard_passed
        or not previous.store_shock_guard_passed
    ):
        return PersistenceDecision(
            action=PromptAction.CONTEXT_REVIEW,
            escalated=False,
            polarity=current_polarity,
            recurring_drivers=recurring_drivers,
            reason="store_shock_guard_not_passed",
        )
    action = (
        PromptAction.RECOGNITION_PROMPT
        if current_polarity is CandidatePolarity.POSITIVE
        else PromptAction.COACHING_PROMPT
    )
    return PersistenceDecision(
        action=action,
        escalated=True,
        polarity=current_polarity,
        recurring_drivers=recurring_drivers,
        reason="two_week_persistence_met",
    )


__all__ = [
    "CandidateAssessment",
    "CandidatePolarity",
    "HybridCalibration",
    "MetricBand",
    "PeerComparison",
    "PeerObservation",
    "PeerReference",
    "PersistenceDecision",
    "PromptAction",
    "RecentMovement",
    "StoreShockAssessment",
    "WeeklyCandidateSignal",
    "assess_common_store_shock",
    "calibrate_hybrid_band",
    "candidate_polarity",
    "classify_candidate",
    "evaluate_two_week_persistence",
    "leave_one_day_stability",
    "leave_one_out_same_store_peer_reference",
    "r7_quantile",
    "round_half_up_to_increment",
    "score_metric",
]
