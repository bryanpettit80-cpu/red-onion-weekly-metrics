from __future__ import annotations

from datetime import date, timedelta

import pytest

import red_onion_fairness as fairness


def test_r7_quantile_linearly_interpolates_and_rejects_bad_input() -> None:
    assert fairness.r7_quantile([1, 2, 3, 4], 0.75) == pytest.approx(3.25)
    assert fairness.r7_quantile([10], 0.90) == 10
    with pytest.raises(ValueError, match="at least one"):
        fairness.r7_quantile([], 0.75)
    with pytest.raises(ValueError, match="between 0 and 1"):
        fairness.r7_quantile([1], 1.1)


def test_half_up_increment_rounding_resolves_positive_and_negative_ties() -> None:
    assert fairness.round_half_up_to_increment(10.25, 0.50) == 10.50
    assert fairness.round_half_up_to_increment(-10.25, 0.50) == -10.50
    assert fairness.round_half_up_to_increment(0.0045, 0.001) == 0.005


def test_hybrid_calibration_uses_empirical_max_and_separates_equal_bands() -> None:
    empirical = fairness.calibrate_hybrid_band(
        [1, 2, 3, 4],
        business_neutral=1,
        business_strong=2,
        increment=0.5,
    )

    assert empirical.empirical_neutral == pytest.approx(3.25)
    assert empirical.empirical_strong == pytest.approx(3.7)
    assert empirical.band == fairness.MetricBand(neutral=3.5, strong=4.0)

    equal = fairness.calibrate_hybrid_band(
        [1, 1, 1],
        business_neutral=1,
        business_strong=1,
        increment=0.5,
    )
    assert equal.band == fairness.MetricBand(neutral=1.0, strong=1.5)


@pytest.mark.parametrize(
    ("deviation", "higher_is_better", "expected"),
    [
        (4.99, True, 0),
        (5.0, True, 1),
        (9.99, True, 1),
        (10.0, True, 2),
        (-10.0, True, -2),
        (-10.0, False, 2),
        (float("nan"), True, None),
        (None, True, None),
    ],
)
def test_signed_metric_scoring(
    deviation: float | None,
    higher_is_better: bool,
    expected: int | None,
) -> None:
    assert (
        fairness.score_metric(
            deviation,
            fairness.MetricBand(neutral=5, strong=10),
            higher_is_better=higher_is_better,
        )
        == expected
    )


def peer_observations(
    *,
    weeks: int = 4,
    peers_per_week: int = 6,
) -> tuple[list[fairness.PeerObservation], list[date]]:
    first_week = date(2026, 5, 3)
    week_ends = [first_week + timedelta(days=7 * index) for index in range(weeks)]
    rows = [
        fairness.PeerObservation(
            person_id=f"peer-{peer_index}",
            location="RC Richmond",
            week_end=week_end,
            value=float(peer_index),
        )
        for week_end in week_ends
        for peer_index in range(1, peers_per_week + 1)
    ]
    return rows, week_ends


def test_peer_reference_is_same_store_leave_one_person_out_and_sufficient() -> None:
    rows, week_ends = peer_observations()
    rows.extend(
        [
            fairness.PeerObservation(
                person_id="focus",
                location="RC Richmond",
                week_end=week_end,
                value=10_000,
            )
            for week_end in week_ends
        ]
    )
    rows.extend(
        [
            fairness.PeerObservation(
                person_id="other-store",
                location="RC Virginia Beach",
                week_end=week_end,
                value=-10_000,
            )
            for week_end in week_ends
        ]
    )

    result = fairness.leave_one_out_same_store_peer_reference(
        rows,
        focal_person_id="FOCUS",
        location="rc richmond",
        prior_week_ends=week_ends,
    )

    assert result.sufficient is True
    assert result.reason == "available"
    assert result.value == pytest.approx(3.5)
    assert result.peer_week_count == 24
    assert result.distinct_peer_count == 6
    assert result.usable_weeks == tuple(week_ends)


def test_peer_reference_fails_closed_when_week_or_peer_week_gates_fail() -> None:
    too_few_per_week, week_ends = peer_observations(peers_per_week=4)
    result = fairness.leave_one_out_same_store_peer_reference(
        too_few_per_week,
        focal_person_id="focus",
        location="RC Richmond",
        prior_week_ends=week_ends,
    )
    assert result.sufficient is False
    assert result.value is None
    assert result.reason == "insufficient_usable_weeks"

    enough_per_week, week_ends = peer_observations(weeks=3, peers_per_week=6)
    result = fairness.leave_one_out_same_store_peer_reference(
        enough_per_week,
        focal_person_id="focus",
        location="RC Richmond",
        prior_week_ends=week_ends,
    )
    assert result.sufficient is False
    assert result.peer_week_count == 18
    assert result.reason == "insufficient_peer_weeks"


def test_duplicate_peer_rows_do_not_overweight_a_person_week() -> None:
    rows, week_ends = peer_observations()
    rows.extend(
        [
            fairness.PeerObservation(
                person_id="peer-1",
                location="RC Richmond",
                week_end=week_ends[0],
                value=101,
            ),
            fairness.PeerObservation(
                person_id="excluded",
                location="RC Richmond",
                week_end=week_ends[0],
                value=-1_000,
                excluded=True,
            ),
            fairness.PeerObservation(
                person_id="unqualified",
                location="RC Richmond",
                week_end=week_ends[0],
                value=-1_000,
                qualified=False,
            ),
        ]
    )

    result = fairness.leave_one_out_same_store_peer_reference(
        rows,
        focal_person_id="focus",
        location="RC Richmond",
        prior_week_ends=week_ends,
    )

    assert result.sufficient is True
    assert result.peer_week_count == 24


def test_candidate_requires_label_alignment_composite_and_agreeing_metrics() -> None:
    result = fairness.classify_candidate(
        "Upward",
        "Above Peer Reference",
        {"check": 2, "wine": 1, "rate": 0, "ticket": 0},
    )
    assert result.eligible is True
    assert result.polarity is fairness.CandidatePolarity.POSITIVE
    assert result.agreeing_drivers == ("check", "wine")

    low_composite = fairness.classify_candidate(
        "Downward",
        "Below Peer Reference",
        {"check": -2, "wine": -1, "rate": 1, "ticket": 0},
    )
    assert low_composite.eligible is False
    assert low_composite.reason == "composite_below_gate"

    not_aligned = fairness.classify_candidate(
        "Upward",
        "Below Peer Reference",
        {"check": 2, "wine": 2, "rate": 0, "ticket": 0},
    )
    assert not_aligned.eligible is False
    assert not_aligned.reason == "movement_peer_not_aligned"


def test_candidate_fails_closed_when_a_required_metric_is_unavailable() -> None:
    result = fairness.classify_candidate(
        "Upward",
        "Above Peer Reference",
        {"check": 2, "wine": 2, "rate": None},
        required_metrics=("check", "wine", "rate", "ticket"),
    )

    assert result.eligible is False
    assert result.polarity is fairness.CandidatePolarity.NONE
    assert result.reason == "metric_unavailable"


def test_store_shock_guard_requires_material_person_peer_difference() -> None:
    bands = {
        "check": fairness.MetricBand(neutral=5, strong=10),
        "rate": fairness.MetricBand(neutral=0.02, strong=0.04),
    }
    directions = {"check": True, "rate": False}
    common = fairness.assess_common_store_shock(
        "positive",
        {"check": 12, "rate": -0.02},
        {"check": 10, "rate": -0.01},
        bands,
        directions,
    )
    assert common.guard_passed is False
    assert common.common_store_shock is True
    assert common.reason == "common_store_movement"

    distinct = fairness.assess_common_store_shock(
        "positive",
        {"check": 17, "rate": -0.04},
        {"check": 10, "rate": -0.01},
        bands,
        directions,
    )
    assert distinct.guard_passed is True
    assert distinct.common_store_shock is False
    assert distinct.differentiating_metrics == ("check", "rate")


def weekly_signal(
    week_end: date,
    *,
    polarity: fairness.CandidatePolarity = fairness.CandidatePolarity.NEGATIVE,
    drivers: tuple[str, ...] = ("check", "wine"),
    qualified: bool = True,
    stable: bool = True,
    shock_guard_passed: bool = True,
) -> fairness.WeeklyCandidateSignal:
    return fairness.WeeklyCandidateSignal(
        week_end=week_end,
        polarity=polarity,
        drivers=drivers,
        qualified=qualified,
        leave_one_day_stable=stable,
        store_shock_guard_passed=shock_guard_passed,
    )


def test_two_week_gate_is_symmetric_for_coaching_and_recognition() -> None:
    previous_week = date(2026, 7, 12)
    current_week = date(2026, 7, 19)
    coaching = fairness.evaluate_two_week_persistence(
        weekly_signal(current_week),
        weekly_signal(previous_week, drivers=("check",)),
    )
    assert coaching.action is fairness.PromptAction.COACHING_PROMPT
    assert coaching.escalated is True
    assert coaching.recurring_drivers == ("check",)

    recognition = fairness.evaluate_two_week_persistence(
        weekly_signal(
            current_week,
            polarity=fairness.CandidatePolarity.POSITIVE,
        ),
        weekly_signal(
            previous_week,
            polarity=fairness.CandidatePolarity.POSITIVE,
            drivers=("wine",),
        ),
    )
    assert recognition.action is fairness.PromptAction.RECOGNITION_PROMPT
    assert recognition.escalated is True


@pytest.mark.parametrize(
    ("previous", "current_changes", "reason"),
    [
        (None, {}, "first_qualified_candidate"),
        (
            {"week_end": date(2026, 7, 5)},
            {},
            "nonconsecutive_week",
        ),
        (
            {},
            {"drivers": ("ticket",)},
            "no_recurring_driver",
        ),
        (
            {"stable": False},
            {},
            "day_sensitive",
        ),
        (
            {"shock_guard_passed": False},
            {},
            "store_shock_guard_not_passed",
        ),
    ],
)
def test_two_week_gate_downgrades_unstable_evidence_to_context_review(
    previous: dict[str, object] | None,
    current_changes: dict[str, object],
    reason: str,
) -> None:
    current_values = {
        "week_end": date(2026, 7, 19),
        **current_changes,
    }
    current = weekly_signal(**current_values)
    if previous is None:
        previous_signal = None
    else:
        previous_values = {
            "week_end": date(2026, 7, 12),
            **previous,
        }
        previous_signal = weekly_signal(**previous_values)

    result = fairness.evaluate_two_week_persistence(current, previous_signal)

    assert result.action is fairness.PromptAction.CONTEXT_REVIEW
    assert result.escalated is False
    assert result.reason == reason


def test_leave_one_day_stability_requires_every_removal_to_preserve_polarity() -> None:
    assert fairness.leave_one_day_stability(
        "negative",
        [
            fairness.CandidatePolarity.NEGATIVE,
            fairness.CandidatePolarity.NEGATIVE,
        ],
    )
    assert not fairness.leave_one_day_stability(
        "negative",
        [
            fairness.CandidatePolarity.NEGATIVE,
            fairness.CandidatePolarity.NONE,
        ],
    )
    assert not fairness.leave_one_day_stability("negative", [])
