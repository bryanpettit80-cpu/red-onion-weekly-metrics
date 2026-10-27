"""Descriptive conversations for every resolved person in the selected weeks.

The caller supplies resolved person identities and the existing peer eligibility
flag. Small subject samples stay visible; an absent denominator or peer reference
stays unavailable. This module does not create or change Management Center actions.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date
from decimal import Decimal, ROUND_HALF_UP, localcontext
from math import isfinite
from statistics import median, stdev
from typing import Any, Iterable


STAFF_FEEDBACK_METHODOLOGY_VERSION = "2026.10-v3"
SPG_CONVERSATION_THRESHOLD = 5.0
WINE_CONVERSATION_THRESHOLD = 0.015
NEAR_THRESHOLD_FRACTION = 0.10
GUEST_WEIGHT_CAP = 50.0


def _number(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    try:
        result = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    return result if isfinite(result) else None


def currency_display_value(value: float) -> float:
    """Round a finite currency amount to cents as Excel displays half cents."""
    # Excel retains 15 significant digits. Remove binary scaling/subtraction
    # noise before cent rounding (for example, .01655 * 100 is just below 1.655).
    with localcontext() as context:
        context.prec = 15
        context.rounding = ROUND_HALF_UP
        amount = +Decimal(str(value))
    return float(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _topic(gap: float | None, metric: str) -> dict[str, Any]:
    wine = metric == "wine"
    threshold = WINE_CONVERSATION_THRESHOLD if wine else SPG_CONVERSATION_THRESHOLD
    display_threshold = threshold * 100 if wine else threshold
    display_gap = currency_display_value(gap * 100 if wine else gap) if gap is not None else None
    near = display_gap is not None and abs(abs(display_gap) - display_threshold) <= display_threshold * NEAR_THRESHOLD_FRACTION + 1e-9
    if display_gap is None:
        direction, action = "Unavailable", "Comparison unavailable"
    elif near:
        direction = "Higher" if display_gap > 0 else "Lower"
        action = "Near conversation threshold"
    elif display_gap >= display_threshold:
        direction, action = "Higher", "Recognition opportunity"
    elif display_gap <= -display_threshold:
        direction, action = "Lower", "Coaching conversation"
    else:
        direction, action = "Around", "Routine check-in"
    return {
        "metric": metric,
        "label": "Wine sales per $100" if wine else "Spend per guest",
        "direction": direction,
        "action": action,
        "gap": gap,
        "display_gap": display_gap,
        "near_cutoff": near,
        "threshold": threshold,
        "repeat_count": 0,
        "repeat_weeks": 0,
        "repeat_rate": 0.0,
    }


def _weighted(rows: Iterable[dict[str, Any]], field: str) -> float | None:
    selected = [row for row in rows if row.get(field) is not None and row["guest_weight"] > 0]
    weight = sum(row["guest_weight"] for row in selected)
    return sum(row[field] * row["guest_weight"] for row in selected) / weight if weight else None


def _priority(row: dict[str, Any], action: str) -> tuple[Any, ...]:
    topics = [row[key] for key in ("spg_topic", "wine_topic") if row[key]["action"] == action]
    evidence = max(
        ((topic["repeat_count"], topic["repeat_rate"], round(abs(topic["display_gap"]) / (1.5 if topic["metric"] == "wine" else 5.0), 2)) for topic in topics),
        default=(0, 0.0, 0.0),
    )
    coverage = {"Established": 3, "Building": 2, "Early indication": 1, "No comparison": 0}[row["data_coverage"]]
    # A single-guest total must not win an otherwise comparable conversation
    # priority merely because its ratio is extreme. The visible guest count is
    # the basis for this ordering, not an undisclosed performance score.
    one_guest_total = row["compared_guests"] <= 1
    return (not row["present_latest_week"], -coverage, one_guest_total, -evidence[0], -evidence[1], -evidence[2], row["location"], row["display_name"].casefold(), row["raw_user_name"])


def build_staff_feedback(
    person_rows: list[dict[str, Any]],
    week_ends: Iterable[date],
    config: dict[str, Any],
) -> dict[str, Any]:
    """Compare all valid subjects with other eligible same-store weekly peers.

    No input rows are mutated. Dates are the caller's complete shared weeks;
    the latest eight are used. Comparisons are metric-specific when wine is
    missing. Actual and peer averages always use identical weeks and weights.
    """
    weeks = tuple(sorted(set(week_ends))[-8:])
    selected = set(weeks)
    latest = weeks[-1] if weeks else None
    min_peers = max(1, int(config.get("management_peer_reference", {}).get("min_distinct_peers_per_week", 5)))
    normalized: list[dict[str, Any]] = []
    seen: set[tuple[Any, str, str]] = set()
    for source in person_rows:
        if source["week_end"] not in selected:
            continue
        location, name = str(source["location"]), str(source["raw_user_name"])
        key = (source["week_end"], location, name)
        if key in seen:
            raise ValueError(f"Duplicate resolved person/week: {location}, {name}, {source['week_end']}")
        seen.add(key)
        sales, guests, wine = (_number(source.get(field)) for field in ("gross_sales", "guest_count", "wine_sales"))
        active = _number(source.get("active_days")) or 0.0
        valid = sales is not None and sales > 0 and guests is not None and guests > 0 and active >= 1
        spg = _number(sales / guests) if valid else None
        valid = valid and spg is not None
        raw_names = source.get("source_raw_names") or (name,)
        if isinstance(raw_names, str):
            raw_names = (raw_names,)
        normalized.append({
            "week_end": source["week_end"], "location": location, "raw_user_name": name,
            "display_name": str(source.get("display_name") or name),
            "source_raw_names": tuple(sorted(str(value) for value in raw_names)),
            "gross_sales": sales, "guest_count": guests, "wine_sales": wine,
            "check_average": spg,
            "wine_pct": _number(wine / sales) if valid and wine is not None else None,
            "active_days": active, "valid_subject": valid,
            "feedback_peer_eligible": bool(source.get("feedback_peer_eligible")) and valid,
            "daily_records": source.get("daily_records") or (),
        })
    pools: dict[tuple[date, str], list[dict[str, Any]]] = defaultdict(list)
    roster: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in normalized:
        roster[(row["location"], row["raw_user_name"])].append(row)
        if row["feedback_peer_eligible"]:
            pools[(row["week_end"], row["location"])].append(row)
    detail_rows: list[dict[str, Any]] = []
    detail_by_person: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in sorted(normalized, key=lambda item: (item["week_end"], item["location"], item["raw_user_name"])):
        peers = [peer for peer in pools[(row["week_end"], row["location"])] if peer["raw_user_name"] != row["raw_user_name"]]
        wine_peers = [peer for peer in peers if peer["wine_pct"] is not None]
        peer_spg = median(peer["check_average"] for peer in peers) if row["valid_subject"] and len(peers) >= min_peers else None
        peer_wine = median(peer["wine_pct"] for peer in wine_peers) if row["valid_subject"] and row["wine_pct"] is not None and len(wine_peers) >= min_peers else None
        spg_gap = row["check_average"] - peer_spg if peer_spg is not None else None
        wine_gap = row["wine_pct"] - peer_wine if peer_wine is not None else None
        compared = spg_gap is not None or wine_gap is not None
        if not row["valid_subject"]:
            note = "Needs positive recorded sales and guests, and at least one active day."
        elif len(peers) < min_peers:
            note = f"{len(peers)} eligible same-store peers; needs {min_peers}."
        elif row["wine_pct"] is None:
            note = "Sales per guest compared; recorded wine sales unavailable."
        elif len(wine_peers) < min_peers:
            note = f"Sales per guest compared; {len(wine_peers)} peers have wine data, needs {min_peers}."
        else:
            note = "Compared with other eligible servers in the same store and week."
        detail = {key: value for key, value in row.items() if key != "daily_records"}
        detail.update({
            "distinct_peer_count": len(peers), "wine_peer_count": len(wine_peers),
            "peer_spg": peer_spg, "peer_wine_pct": peer_wine,
            "spg_gap": spg_gap, "wine_gap": wine_gap,
            "guest_weight": min(row["guest_count"], GUEST_WEIGHT_CAP) if compared else 0.0,
            "compared": compared, "comparison_note": note,
            "spg_topic": _topic(spg_gap, "spg"), "wine_topic": _topic(wine_gap, "wine"),
        })
        detail_rows.append(detail)
        detail_by_person[(row["location"], row["raw_user_name"])].append(detail)

    rows: list[dict[str, Any]] = []
    for key, observed in roster.items():
        current = max(observed, key=lambda row: row["week_end"])
        details = detail_by_person[key]
        compared = [row for row in details if row["compared"]]
        spg_rows = [row for row in details if row["spg_gap"] is not None]
        wine_rows = [row for row in details if row["wine_gap"] is not None]
        guests = sum(row["guest_count"] for row in compared)
        count = len(compared)
        coverage = "Established" if count >= 6 and guests >= 200 else "Building" if count >= 4 and guests >= 150 else "Early indication" if count else "No comparison"
        topics = {"spg_topic": _topic(_weighted(spg_rows, "spg_gap"), "spg"), "wine_topic": _topic(_weighted(wine_rows, "wine_gap"), "wine")}
        for topic_key, matched, field in (("spg_topic", spg_rows, "spg_gap"), ("wine_topic", wine_rows, "wine_gap")):
            topic = topics[topic_key]
            direction = 1 if topic["direction"] == "Higher" else -1 if topic["direction"] == "Lower" else 0
            threshold = 1.5 if topic["metric"] == "wine" else 5.0
            repeated = sum(direction * _topic(row[field], topic["metric"])["display_gap"] >= threshold for row in matched) if direction else 0
            topic.update(repeat_count=repeated, repeat_weeks=len(matched), repeat_rate=repeated / len(matched) if matched else 0.0)
        coach_topics = [topic["label"].lower() for topic in topics.values() if topic["action"] == "Coaching conversation"]
        recognition_topics = [topic["label"].lower() for topic in topics.values() if topic["action"] == "Recognition opportunity"]
        near = any(topic["near_cutoff"] for topic in topics.values())
        overall = "Mixed results" if coach_topics and recognition_topics else "Coaching conversation" if coach_topics else "Recognition opportunity" if recognition_topics else "Near conversation threshold" if near else "Routine check-in" if compared else "Comparison unavailable"
        active = [row for row in observed if row["active_days"] >= 1 and ((row["gross_sales"] or 0) > 0 or (row["guest_count"] or 0) > 0)]
        worked_dates = [day["report_date"] for row in active for day in row["daily_records"] if isinstance(day.get("report_date"), date) and ((_number(day.get("gross_sales")) or 0) > 0 or (_number(day.get("guest_count")) or 0) > 0)]
        present = any(row["week_end"] == latest for row in active)
        sales_total = sum(row["gross_sales"] or 0.0 for row in observed)
        guest_total = sum(row["guest_count"] or 0.0 for row in observed)
        complete_wine = all(row["wine_sales"] is not None for row in observed)
        wine_total = sum(row["wine_sales"] or 0.0 for row in observed)
        spg_sd = stdev(row["spg_gap"] for row in spg_rows) if len(spg_rows) >= 2 else None
        wine_sd = stdev(row["wine_gap"] for row in wine_rows) if len(wine_rows) >= 2 else None
        thresholds = config.get("management_score_thresholds", {})
        spg_limits = {"neutral": 11.5, "strong": 17.5, **thresholds.get("check_average", {})}
        wine_limits = {"neutral": .041, "strong": .057, **thresholds.get("wine_pct", {})}
        if not compared:
            consistency = "Unavailable"
        elif count == 1:
            consistency = "One week only"
        elif spg_sd is None or wine_sd is None:
            consistency = "Limited history"
        elif spg_sd <= spg_limits["neutral"] and wine_sd <= wine_limits["neutral"]:
            consistency = "Steady"
        elif spg_sd <= spg_limits["strong"] and wine_sd <= wine_limits["strong"]:
            consistency = "Some variation"
        else:
            consistency = "Variable"
        one_guest_weeks = sum(row["guest_count"] == 1 for row in compared)
        coverage_note = f"{count} compared {'week' if count == 1 else 'weeks'} / {guests:,.0f} guests. {coverage}."
        if coverage == "Early indication":
            coverage_note += " Use this as an initial conversation and confirm against more shifts."
        if one_guest_weeks:
            coverage_note += f" Includes {one_guest_weeks} {'week' if one_guest_weeks == 1 else 'weeks'} with only 1 guest; verify the guest count and sale."
        if len(spg_rows) != len(wine_rows):
            coverage_note += f" Sales comparison: {len(spg_rows)} weeks; wine comparison: {len(wine_rows)} weeks."
        prefix = "Early indication: " if coverage == "Early indication" else ""
        coach_prompt = prefix + "Discuss " + " and ".join(coach_topics) + ". Ask about guest preferences, section mix, and recommendation opportunities." if coach_topics else ""
        recognition_prompt = prefix + "Recognize higher recorded " + " and ".join(recognition_topics) + ". Ask which recommendations worked and can be shared." if recognition_topics else ""
        checkin_prompt = "Review results near the conversation threshold; a small change should not determine the feedback." if near else "Ask what worked and what support would help on the next shifts." if compared else "Confirm the recorded sales, guest count, and available peer comparison."
        feedback = " ".join(value for value in (coach_prompt, recognition_prompt, checkin_prompt if near or not (coach_topics or recognition_topics) else "") if value)
        if not present:
            feedback = "No recorded activity in the latest week. " + feedback
        rows.append({
            "location": key[0], "raw_user_name": key[1], "display_name": current["display_name"],
            "source_raw_names": tuple(sorted({name for row in observed for name in row["source_raw_names"]})),
            "present_latest_week": present,
            "last_worked_date": max(worked_dates) if worked_dates else None,
            "last_worked_week_end": max(row["week_end"] for row in active) if active else None,
            "observed_weeks": len(active), "observed_guests": guest_total,
            "observed_spg": sales_total / guest_total if guest_total > 0 else None,
            "observed_wine_pct": wine_total / sales_total if sales_total > 0 and complete_wine else None,
            "compared_weeks": count, "compared_guests": guests,
            "spg_compared_weeks": len(spg_rows), "wine_compared_weeks": len(wine_rows),
            "comparison_spg": _weighted(spg_rows, "check_average"), "peer_spg": _weighted(spg_rows, "peer_spg"),
            "average_spg_gap": topics["spg_topic"]["gap"],
            "comparison_wine_pct": _weighted(wine_rows, "wine_pct"), "peer_wine_pct": _weighted(wine_rows, "peer_wine_pct"),
            "average_wine_gap": topics["wine_topic"]["gap"],
            "weekly_spg_gap_sd": spg_sd, "weekly_wine_gap_sd": wine_sd,
            "data_coverage": coverage, "coverage_note": coverage_note,
            "consistency": consistency,
            "consistency_note": f"{consistency} across {count} compared {'week' if count == 1 else 'weeks'}; describes variation, not result quality.",
            "overall_read": overall, **topics,
            "coaching_prompt": coach_prompt, "recognition_prompt": recognition_prompt,
            "checkin_prompt": checkin_prompt, "feedback_text": feedback,
            "near_cutoff": near, "one_guest_weeks": one_guest_weeks,
        })
    rows.sort(key=lambda row: _priority(row, ""))
    coaching = sorted((row for row in rows if row["coaching_prompt"]), key=lambda row: _priority(row, "Coaching conversation"))
    recognition = sorted((row for row in rows if row["recognition_prompt"]), key=lambda row: _priority(row, "Recognition opportunity"))
    checkins = [row for row in rows if not row["coaching_prompt"] and not row["recognition_prompt"]]
    counts = {
        "people": len(rows), "compared": sum(row["compared_weeks"] > 0 for row in rows),
        "current": sum(row["present_latest_week"] for row in rows),
        "coaching": len(coaching), "recognition": len(recognition), "checkin": len(checkins),
        "mixed": sum(row["overall_read"] == "Mixed results" for row in rows),
        "near_cutoff": sum(row["near_cutoff"] for row in rows),
        "established": sum(row["data_coverage"] == "Established" for row in rows),
        "building": sum(row["data_coverage"] == "Building" for row in rows),
        "early_indication": sum(row["data_coverage"] == "Early indication" for row in rows),
        "unavailable": sum(row["data_coverage"] == "No comparison" for row in rows),
    }
    return {
        "methodology_version": STAFF_FEEDBACK_METHODOLOGY_VERSION,
        "week_ends": weeks, "latest_complete_week_end": latest,
        "rows": rows, "detail_rows": detail_rows,
        "coaching_rows": coaching, "recognition_rows": recognition, "checkin_rows": checkins,
        "counts": counts,
        "thresholds": {"spg": SPG_CONVERSATION_THRESHOLD, "wine": WINE_CONVERSATION_THRESHOLD,
                       "near_fraction": NEAR_THRESHOLD_FRACTION, "guest_weight_cap": GUEST_WEIGHT_CAP,
                       "minimum_peers": min_peers},
    }
