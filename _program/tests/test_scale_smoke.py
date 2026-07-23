from __future__ import annotations

from datetime import date, timedelta

import red_onion_weekly_metrics as metrics


def test_three_year_analytical_scale_smoke_is_deterministic() -> None:
    records: list[metrics.MetricRecord] = []
    first_week_end = date(2023, 8, 6)
    locations = ("RC Richmond", "RC Virginia Beach")
    servers = tuple(f"Alex {index}" for index in range(6))
    for week_offset in range(156):
        week_end = first_week_end + timedelta(days=7 * week_offset)
        for location_index, location in enumerate(locations):
            for server_index, server in enumerate(servers):
                for day_offset in range(5, -1, -1):
                    report_date = week_end - timedelta(days=day_offset)
                    gross_sales = 120.0 + location_index * 10 + server_index + week_offset / 10
                    guest_count = 8.0 + (server_index % 3)
                    records.append(
                        metrics.MetricRecord(
                            source_file=f"Daily Report {report_date}.xlsx",
                            report_date=report_date,
                            location=location,
                            raw_user_name=server,
                            display_name=server,
                            is_location_total=False,
                            gross_sales=gross_sales,
                            guest_count=guest_count,
                            check_average=gross_sales / guest_count,
                            wine_sales=gross_sales * 0.1,
                            wine_pct=0.1,
                            rate_of_sale_by_guest_count=0.2,
                            average_ticket_time_seconds=4200.0,
                        )
                    )
            for day_offset in range(5, -1, -1):
                report_date = week_end - timedelta(days=day_offset)
                base = 120.0 + location_index * 10 + week_offset / 10
                gross_sales = sum(base + server_index for server_index in range(6))
                guest_count = 54.0
                records.append(
                    metrics.MetricRecord(
                        source_file=f"Daily Report {report_date}.xlsx",
                        report_date=report_date,
                        location=location,
                        raw_user_name="",
                        display_name="",
                        is_location_total=True,
                        gross_sales=gross_sales,
                        guest_count=guest_count,
                        check_average=gross_sales / guest_count,
                        wine_sales=gross_sales * 0.1,
                        wine_pct=0.1,
                        rate_of_sale_by_guest_count=0.2,
                        average_ticket_time_seconds=4200.0,
                    )
                )

    weekly_servers, weekly_locations = metrics.weekly_rollups(records)
    ranked = metrics.weekly_server_rank_rows(weekly_servers, 1)
    config = metrics.load_config(
        metrics.PROGRAM_DIR / "__missing_scale_smoke_config__.json"
    )
    management = metrics.management_server_rows(
        weekly_servers, weekly_locations, ranked, {}, config
    )

    assert len(records) == 13_104
    assert len(weekly_servers) == 1_872
    assert len(weekly_locations) == 312
    assert len(management) == 12
    assert {row["week_end"] for row in management} == {
        first_week_end + timedelta(days=7 * 155)
    }
    assert all(
        row["evidence_week_ends"] == sorted(set(row["evidence_week_ends"]))
        for row in management
    )
