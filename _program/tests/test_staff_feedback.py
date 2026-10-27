from __future__ import annotations

from copy import deepcopy
from datetime import date, timedelta

import pytest

from red_onion_staff_feedback import _topic, build_staff_feedback, currency_display_value


END = date(2026, 10, 25)


@pytest.mark.parametrize("value,shown", [
    (100.625, 100.63), (-100.625, -100.63),
    (2.675, 2.68), (-2.675, -2.68),
    (5.504, 5.50), (5.505, 5.51), (-5.505, -5.51),
    (1.6549999999999998, 1.66), (-1.6549999999999998, -1.66),
    (5.504999999999995, 5.51), (5.5049, 5.50),
])
def test_currency_display_rounds_half_cents_away_from_zero(value: float, shown: float) -> None:
    assert currency_display_value(value) == shown


@pytest.mark.parametrize("gap,metric,shown,action", [
    (5.505, "spg", 5.51, "Recognition opportunity"),
    (-5.505, "spg", -5.51, "Coaching conversation"),
    (4.495, "spg", 4.50, "Near conversation threshold"),
    (-4.495, "spg", -4.50, "Near conversation threshold"),
    (.01655, "wine", 1.66, "Recognition opportunity"),
    (-.01655, "wine", -1.66, "Coaching conversation"),
    (.01345, "wine", 1.35, "Near conversation threshold"),
    (-.01345, "wine", -1.35, "Near conversation threshold"),
])
def test_half_cent_topic_boundaries_match_displayed_currency(
    gap: float, metric: str, shown: float, action: str,
) -> None:
    topic = _topic(gap, metric)
    assert topic["display_gap"] == shown
    assert topic["action"] == action


def person(name: str, *, week: date = END, spg: float = 100, wine: float = .10,
           guests: float = 50, days: int = 3, eligible: bool = False) -> dict:
    sales = spg * guests
    return {
        "week_end": week, "location": "Store A", "raw_user_name": name,
        "display_name": name, "source_raw_names": (name,),
        "gross_sales": sales, "guest_count": guests, "wine_sales": sales * wine,
        "active_days": days, "feedback_peer_eligible": eligible,
        "daily_records": [{"report_date": week - timedelta(days=1),
                           "gross_sales": sales, "guest_count": guests}],
    }


def peers(weeks=(END,), count=6) -> list[dict]:
    return [person(f"Peer {number}", week=week, eligible=True)
            for week in weeks for number in range(count)]


def find(model: dict, name: str) -> dict:
    return next(row for row in model["rows"] if row["raw_user_name"] == name)


def test_one_guest_subject_is_compared_and_keeps_caveat() -> None:
    model = build_staff_feedback(peers() + [person("Small", spg=250, guests=1, days=1)], [END], {})
    row = find(model, "Small")
    assert row["comparison_spg"] == 250
    assert row["peer_spg"] == 100
    assert row["average_spg_gap"] == 150
    assert row["data_coverage"] == "Early indication"
    assert row["overall_read"] == "Recognition opportunity"
    assert row["consistency"] == "One week only"
    assert "only 1 guest" in row["coverage_note"]
    assert "Early indication" in row["recognition_prompt"]


def test_opposite_topics_make_mixed_row_and_explicit_list_overlap() -> None:
    model = build_staff_feedback(peers() + [person("Mixed", spg=80, wine=.15)], [END], {})
    row = find(model, "Mixed")
    assert row["overall_read"] == "Mixed results"
    assert row in model["coaching_rows"] and row in model["recognition_rows"]
    assert row not in model["checkin_rows"]
    assert "spend per guest" in row["coaching_prompt"]
    assert "wine sales per $100" in row["recognition_prompt"]
    assert model["counts"]["mixed"] == 1


@pytest.mark.parametrize("gap", [-5.5, -5.0, -4.5, 4.5, 5.0, 5.5, 5.504])
def test_boundary_band_does_not_create_coaching_or_recognition(gap: float) -> None:
    model = build_staff_feedback(peers() + [person("Boundary", spg=100 + gap)], [END], {})
    row = find(model, "Boundary")
    assert row["spg_topic"]["action"] == "Near conversation threshold"
    assert row["overall_read"] == "Near conversation threshold"
    assert not row["coaching_prompt"] and not row["recognition_prompt"]
    assert row in model["checkin_rows"]


@pytest.mark.parametrize("gap,action", [(5.51, "Recognition opportunity"), (-5.51, "Coaching conversation")])
def test_beyond_boundary_uses_displayed_gap(gap: float, action: str) -> None:
    row = find(build_staff_feedback(peers() + [person("Outside", spg=100 + gap)], [END], {}), "Outside")
    assert row["spg_topic"]["display_gap"] == gap
    assert row["spg_topic"]["action"] == action


def test_near_wine_boundary_does_not_cancel_clear_sales_topic() -> None:
    row = find(build_staff_feedback(peers() + [person("Both", spg=110, wine=.115)], [END], {}), "Both")
    assert row["spg_topic"]["action"] == "Recognition opportunity"
    assert row["wine_topic"]["action"] == "Near conversation threshold"
    assert row["overall_read"] == "Recognition opportunity"
    assert "near the conversation threshold" in row["feedback_text"]


@pytest.mark.parametrize("gap,shown,action", [
    (.01651, 1.65, "Near conversation threshold"),
    (.0166, 1.66, "Recognition opportunity"),
    (-.01651, -1.65, "Near conversation threshold"),
    (-.0166, -1.66, "Coaching conversation"),
    (.01349, 1.35, "Near conversation threshold"),
    (.0134, 1.34, "Routine check-in"),
])
def test_wine_boundary_uses_two_decimal_dollars_per_hundred(gap: float, shown: float, action: str) -> None:
    row = find(build_staff_feedback(peers() + [person("Wine boundary", wine=.10 + gap)], [END], {}), "Wine boundary")
    assert row["wine_topic"]["display_gap"] == shown
    assert row["wine_topic"]["action"] == action


def test_missing_peers_and_zero_denominator_remain_unavailable() -> None:
    no_peers = find(build_staff_feedback(peers(count=4) + [person("Small pool")], [END], {}), "Small pool")
    assert no_peers["comparison_spg"] is None
    assert no_peers["overall_read"] == "Comparison unavailable"
    zero = person("Zero", guests=0)
    zero["gross_sales"] = 200
    model = build_staff_feedback(peers() + [zero], [END], {})
    row = find(model, "Zero")
    assert row["present_latest_week"]
    assert row["comparison_spg"] is None and row["observed_spg"] is None
    assert row["compared_guests"] == 0
    assert "Zero" in {item["raw_user_name"] for item in model["rows"]}


def test_missing_wine_does_not_hide_valid_sales_comparison_or_invent_zero() -> None:
    subject = person("No wine", spg=120)
    subject.pop("wine_sales")
    row = find(build_staff_feedback(peers() + [subject], [END], {}), "No wine")
    assert row["average_spg_gap"] == 20
    assert row["comparison_wine_pct"] is None and row["average_wine_gap"] is None
    assert row["observed_wine_pct"] is None
    assert row["spg_compared_weeks"] == 1 and row["wine_compared_weeks"] == 0


def test_nonfinite_subject_and_peers_cannot_produce_fabricated_comparison() -> None:
    bad = person("Invalid")
    bad["gross_sales"] = float("nan")
    peer_rows = peers(count=5)
    peer_rows[0]["guest_count"] = float("inf")
    model = build_staff_feedback(peer_rows + [bad, person("Valid")], [END], {})
    assert find(model, "Invalid")["comparison_spg"] is None
    assert find(model, "Valid")["comparison_spg"] is None


def test_alias_provenance_and_dates_are_preserved_without_input_mutation() -> None:
    subject = person("Resolved")
    subject.update(display_name="Display Name", source_raw_names=("Old POS", "New POS"))
    source = peers() + [subject]
    before = deepcopy(source)
    row = find(build_staff_feedback(source, [END], {}), "Resolved")
    assert row["display_name"] == "Display Name"
    assert row["source_raw_names"] == ("New POS", "Old POS")
    assert row["last_worked_date"] == END - timedelta(days=1)
    assert source == before


def test_matched_values_and_capped_weights_reconcile() -> None:
    weeks = [END - timedelta(days=7), END]
    source = peers(weeks) + [person("Weighted", week=weeks[0], spg=120, guests=100),
                            person("Weighted", week=weeks[1], spg=200, guests=1, days=1)]
    row = find(build_staff_feedback(source, weeks, {}), "Weighted")
    assert row["average_spg_gap"] == pytest.approx((20 * 50 + 100) / 51)
    assert row["comparison_spg"] - row["peer_spg"] == pytest.approx(row["average_spg_gap"])
    assert row["comparison_wine_pct"] - row["peer_wine_pct"] == pytest.approx(row["average_wine_gap"])
    assert row["weekly_spg_gap_sd"] == pytest.approx(80 / 2 ** .5)


def test_peer_reference_excludes_subject_other_stores_and_other_weeks() -> None:
    peer_rows = [person(f"Different {i}", spg=value, eligible=True)
                 for i, value in enumerate((100, 200, 300, 400, 500))]
    other_store = person("Other store", spg=5000, eligible=True)
    other_store["location"] = "Store B"
    other_week = person("Other week", week=END - timedelta(days=7), spg=5000, eligible=True)
    subject = person("Subject", spg=1000, eligible=True)
    model = build_staff_feedback(peer_rows + [other_store, other_week, subject], [END], {})
    row = find(model, "Subject")
    assert row["peer_spg"] == 300
    detail = next(item for item in model["detail_rows"] if item["raw_user_name"] == "Subject")
    assert detail["distinct_peer_count"] == 5
    assert "Other week" not in {item["raw_user_name"] for item in model["rows"]}


def test_partial_wine_uses_same_available_weeks_for_actual_peer_and_gap() -> None:
    weeks = [END - timedelta(days=7), END]
    first = person("Partial", week=weeks[0], spg=200, guests=2, days=1)
    first["wine_sales"] = None
    second = person("Partial", week=weeks[1], spg=100, guests=100, wine=.20)
    row = find(build_staff_feedback(peers(weeks) + [first, second], weeks, {}), "Partial")
    assert row["spg_compared_weeks"] == 2 and row["wine_compared_weeks"] == 1
    assert row["comparison_wine_pct"] == pytest.approx(.20)
    assert row["peer_wine_pct"] == pytest.approx(.10)
    assert row["average_wine_gap"] == pytest.approx(.10)
    assert row["comparison_spg"] - row["peer_spg"] == pytest.approx(row["average_spg_gap"])
    assert row["comparison_wine_pct"] - row["peer_wine_pct"] == pytest.approx(row["average_wine_gap"])


def test_priority_current_then_sample_then_repetition_not_one_guest_extreme() -> None:
    weeks = [END - timedelta(days=7 * offset) for offset in range(7)]
    source = peers(weeks)
    source += [person("Established current", week=week, spg=110) for week in weeks[:6]]
    source += [person("Established absent", week=week, spg=140) for week in weeks[1:]]
    source += [person("One guest extreme", spg=1000, guests=1, days=1)]
    model = build_staff_feedback(source, weeks, {})
    names = [row["raw_user_name"] for row in model["recognition_rows"]]
    assert names == ["Established current", "One guest extreme", "Established absent"]
    absent = find(model, "Established absent")
    assert not absent["present_latest_week"]
    assert absent["feedback_text"].startswith("No recorded activity in the latest week.")


def test_every_roster_person_has_a_row_and_a_conversation_list() -> None:
    source = peers() + [person("Coach", spg=80), person("Recognize", spg=120),
                        person("Routine"), person("Mixed", spg=80, wine=.15),
                        person("No denominator", guests=0)]
    model = build_staff_feedback(source, [END], {})
    assert {row["raw_user_name"] for row in model["rows"]} == {row["raw_user_name"] for row in source}
    listed = model["coaching_rows"] + model["recognition_rows"] + model["checkin_rows"]
    assert {row["raw_user_name"] for row in listed} == {row["raw_user_name"] for row in source}
    assert len(listed) == len(source) + model["counts"]["mixed"]


def test_single_guest_does_not_win_a_tie_between_early_samples() -> None:
    source = peers() + [person("One guest extreme", spg=1000, guests=1, days=1),
                        person("More observed guests", spg=110, guests=20, days=1)]
    model = build_staff_feedback(source, [END], {})
    assert [row["raw_user_name"] for row in model["recognition_rows"]] == [
        "More observed guests", "One guest extreme"
    ]


def test_duplicate_resolved_identity_week_is_rejected_and_empty_window_is_empty() -> None:
    with pytest.raises(ValueError, match="Duplicate resolved person/week"):
        build_staff_feedback([person("Duplicate"), person("Duplicate")], [END], {})
    assert build_staff_feedback([person("Outside")], [], {})["rows"] == []
