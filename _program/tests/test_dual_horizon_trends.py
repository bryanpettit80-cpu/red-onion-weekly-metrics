from __future__ import annotations

from datetime import date, timedelta

from openpyxl import Workbook

import red_onion_weekly_metrics as metrics


def server_week(
    week_end: date,
    *,
    person: str = "Alex Example",
    check_average: float,
    wine_pct: float,
    rate: float,
    ticket_seconds: float,
    guests: float = 100,
    source_days: int = metrics.OPERATING_WEEK_DAYS,
    rank: int = 5,
) -> tuple[dict, dict]:
    gross_sales = check_average * guests
    row = {
        "week_start": week_end - timedelta(days=metrics.OPERATING_WEEK_DAYS - 1),
        "week_end": week_end,
        "location": "RC Richmond",
        "raw_user_name": person,
        "display_name": person,
        "gross_sales": gross_sales,
        "guest_count": guests,
        "check_average": check_average,
        "wine_sales": gross_sales * wine_pct,
        "wine_pct": wine_pct,
        "rate_of_sale_by_guest_count": rate,
        "average_ticket_time_seconds": ticket_seconds,
        "active_days": source_days,
        "source_days": source_days,
        "source_files": "synthetic.xls",
    }
    ranked = {
        **row,
        "check_average_rank": rank,
        "wine_pct_rank": rank,
        "rate_rank": rank,
        "ticket_time_rank": rank,
    }
    return row, ranked


def score_rows(specs: list[dict]) -> tuple[list[dict], list[dict], list[dict], dict]:
    start = date(2026, 5, 3)
    server_rows: list[dict] = []
    ranked_rows: list[dict] = []
    for index, spec in enumerate(specs):
        row, ranked = server_week(start + timedelta(days=7 * index), **spec)
        server_rows.append(row)
        ranked_rows.append(ranked)
        if index < len(specs) - 1:
            peer, peer_ranked = server_week(
                start + timedelta(days=7 * index),
                person="Peer Reference",
                check_average=20,
                wine_pct=0.10,
                rate=0.20,
                ticket_seconds=4800,
            )
            server_rows.append(peer)
            ranked_rows.append(peer_ranked)
    location_rows = [dict(row) for row in server_rows]
    location_rows = [
        row for row in location_rows if row["raw_user_name"] == "Alex Example"
    ]
    config = metrics.load_config(metrics.PROGRAM_DIR / "__missing_dual_horizon_test_config__.json")
    config["management_peer_reference"].update(
        {
            "min_distinct_peers_per_week": 1,
            "min_peer_server_weeks": 3,
        }
    )
    for family in (
        "management_score_thresholds",
        "management_peer_score_thresholds",
    ):
        config[family] = {
            "check_average": {
                "neutral": 2.5,
                "strong": 5.0,
                "lower_is_better": False,
            },
            "wine_pct": {
                "neutral": 0.005,
                "strong": 0.01,
                "lower_is_better": False,
            },
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
        }
    return server_rows, location_rows, ranked_rows, config


def test_recent_rebound_can_disagree_with_full_8_week_direction() -> None:
    specs = [
        *[
            dict(check_average=30, wine_pct=0.20, rate=0.10, ticket_seconds=3600, rank=1)
            for _ in range(4)
        ],
        *[
            dict(check_average=15, wine_pct=0.05, rate=0.30, ticket_seconds=6000, rank=10)
            for _ in range(3)
        ],
        dict(check_average=26, wine_pct=0.18, rate=0.12, ticket_seconds=3900, rank=2),
    ]
    servers, locations, ranked, config = score_rows(specs)

    result = metrics.management_server_rows(servers, locations, ranked, {}, config)[0]

    assert result["momentum"] == "Upward"
    assert result["long_term_direction"] == "Downward"
    assert result["long_term_history_label"] == "Full"
    assert "8 weeks / 800 guests" in result["history_used"]


def test_six_qualified_weeks_produce_developing_direction() -> None:
    specs = [
        *[
            dict(check_average=15, wine_pct=0.05, rate=0.30, ticket_seconds=6000, guests=25, rank=10)
            for _ in range(2)
        ],
        *[
            dict(check_average=25, wine_pct=0.15, rate=0.15, ticket_seconds=4200, guests=25, rank=3)
            for _ in range(4)
        ],
    ]
    servers, locations, ranked, config = score_rows(specs)

    result = metrics.management_server_rows(servers, locations, ranked, {}, config)[0]

    assert result["long_term_history_label"] == "Developing"
    assert result["long_term_direction"] == "Upward"
    assert "recent 4w / 100g; earlier 2w / 50g" in result["history_used"]


def test_building_history_is_not_scored() -> None:
    specs = [
        dict(check_average=15 + index, wine_pct=0.05, rate=0.30, ticket_seconds=6000, guests=30)
        for index in range(5)
    ]
    servers, locations, ranked, config = score_rows(specs)

    result = metrics.management_server_rows(servers, locations, ranked, {}, config)[0]

    assert result["long_term_history_label"] == "Building History"
    assert result["long_term_direction"] == "Not Evaluated"


def test_eight_low_volume_weeks_do_not_overstate_full_history() -> None:
    specs = [
        dict(
            check_average=20 + index,
            wine_pct=0.10,
            rate=0.20,
            ticket_seconds=4800,
            guests=20,
        )
        for index in range(8)
    ]
    servers, locations, ranked, config = score_rows(specs)

    result = metrics.management_server_rows(servers, locations, ranked, {}, config)[0]

    assert result["long_term_history_label"] == "Developing"


def test_low_current_sample_suppresses_recent_action_but_keeps_qualified_history() -> None:
    specs = [
        *[
            dict(check_average=15, wine_pct=0.05, rate=0.30, ticket_seconds=6000, guests=25, rank=10)
            for _ in range(4)
        ],
        *[
            dict(check_average=25, wine_pct=0.15, rate=0.15, ticket_seconds=4200, guests=25, rank=3)
            for _ in range(3)
        ],
        dict(check_average=25, wine_pct=0.15, rate=0.15, ticket_seconds=4200, guests=20, rank=3),
    ]
    servers, locations, ranked, config = score_rows(specs)

    result = metrics.management_server_rows(servers, locations, ranked, {}, config)[0]

    assert result["momentum"] == "Not Evaluated"
    assert result["action"] == "Monitor"
    assert result["prominent"] is False
    assert result["confidence"] == "Limited Volume"
    assert result["long_term_direction"] == "Upward"


def test_incomplete_latest_week_suppresses_both_trends_and_server_actions() -> None:
    specs = [
        *[
            dict(check_average=15 + index, wine_pct=0.10, rate=0.20, ticket_seconds=4800, rank=5)
            for index in range(8)
        ],
        dict(
            check_average=30,
            wine_pct=0.20,
            rate=0.10,
            ticket_seconds=3600,
            source_days=5,
            rank=1,
        ),
    ]
    servers, locations, ranked, config = score_rows(specs)

    result = metrics.management_server_rows(servers, locations, ranked, {}, config)[0]
    signals = metrics.build_management_action_signals([result], [], [], locations)

    assert result["momentum"] == "Not Evaluated"
    assert result["long_term_direction"] == "Not Evaluated"
    assert result["action"] == "Monitor"
    assert result["confidence"] == "Incomplete Week"
    assert all(signal["Action"] == "Data Quality" for signal in signals)


def test_team_trends_uses_all_person_17_column_schema_and_exact_wow_values() -> None:
    specs = [
        dict(check_average=20 + index, wine_pct=0.10, rate=0.20, ticket_seconds=4800, rank=5)
        for index in range(8)
    ]
    servers, locations, ranked, config = score_rows(specs)
    rows = metrics.management_server_rows(servers, locations, ranked, {}, config)
    workbook = Workbook()

    metrics.write_team_trends_sheet(workbook, rows)
    trends = workbook["Team Trends"]
    headers = [cell.value for cell in trends[3]]

    assert headers == [
        "Location",
        "Server",
        "Week End",
        "Prior Week End",
        "Current Sample",
        "Sales / Guest",
        "WoW Sales / Guest Δ",
        "Wine %",
        "WoW Wine % Δ",
        "WoW Movement",
        "4-Week Movement",
        "8-Week Direction",
        "Peer Comparison",
        "Evidence Status",
        "Action Gate",
        "Trend Drivers",
        "History Used",
    ]
    assert "Score" not in headers
    assert trends.max_row == 4
    assert trends["A4"].value == "RC Richmond"
    assert trends["B4"].value == "Alex Example"
    assert trends["C4"].value == date(2026, 6, 21)
    assert trends["D4"].value == date(2026, 6, 14)
    assert trends["F4"].value == 27
    assert trends["G4"].value == 1
    assert trends["H4"].value == 0.10
    assert trends["I4"].value == 0
    assert trends["J4"].value == "Mixed"
    assert trends["N4"].value == rows[0]["confidence"]
    assert trends["O4"].value == rows[0]["action"]
