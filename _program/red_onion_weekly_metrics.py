from __future__ import annotations

import argparse
import filecmp
import hashlib
import json
import os
import re
import shutil
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo


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

METRICS = [
    ("gross_sales", "Gross Sales"),
    ("guest_count", "Guest Count"),
    ("check_average", "Check Average"),
    ("wine_sales", "Wine Sales"),
    ("rate_of_sale_by_guest_count", "Rate of Sale by Guest Count"),
    ("average_ticket_time", "Average Ticket Time"),
]

MANAGEMENT_METRICS: tuple[tuple[str, str, str], ...] = (
    ("gross_sales", "Gross Sales", "currency"),
    ("guest_count", "Guest Count", "count"),
    ("check_average", "Check Average", "currency"),
    ("wine_pct", "Wine %", "percent"),
    ("rate_of_sale_by_guest_count", "Rate of Sale", "rate"),
    ("average_ticket_time_seconds", "Ticket Time", "duration"),
)

TARGET_FIELDS: tuple[tuple[str, str], ...] = (
    ("gross_sales", "Weekly Sales Target"),
    ("guest_count", "Weekly Guest Target"),
    ("check_average", "Check Average Target"),
    ("wine_pct", "Wine % Target"),
    ("rate_of_sale_by_guest_count", "Rate Target"),
    ("average_ticket_time_seconds", "Ticket Time Target (Min)"),
)

ACTION_HEADERS = [
    "Action ID",
    "Entity Key",
    "Priority",
    "Status",
    "Owner",
    "Due Date",
    "Location",
    "Person / Area",
    "Action",
    "Signal",
    "Why It Matters",
    "Recommended Next Step",
    "Last Seen",
    "Manager Notes",
    "Performance Level",
    "Momentum",
    "First Seen",
    "Weeks Open",
    "Confidence",
    "Signal State",
]

VISIBLE_MANAGEMENT_SHEETS = [
    "Dashboard",
    "Action Board",
    "Server Scorecard",
    "Store & Group Scorecards",
    "Rising & Falling Stars",
    "Action History",
    "Data Quality",
    "Management Setup",
    "Run Notes",
]

OPERATING_WEEK_START_WEEKDAY = 1  # Tuesday; date.weekday() uses Monday=0.
OPERATING_WEEK_END_WEEKDAY = 6  # Sunday.
OPERATING_WEEK_DAYS = 6
OPERATING_WEEK_LABEL = "Tuesday-Sunday"
FILENAME_DATE_BUSINESS_DATE_OFFSET_DAYS = 1
DAILY_REPORT_PATTERN = "Daily Report*.xls"
PROGRAM_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = PROGRAM_DIR.parent
DEFAULT_INPUT_DIR = PROJECT_ROOT / "Daily Reports"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "Output"
DEFAULT_ARCHIVE_DIR = PROJECT_ROOT / "Archive - Old Files"
DEFAULT_CONFIG_PATH = PROGRAM_DIR / "red_onion_config.json"


@dataclass(frozen=True)
class MetricRecord:
    source_file: str
    report_date: date
    location: str
    raw_user_name: str
    display_name: str
    is_location_total: bool
    gross_sales: float
    guest_count: float
    check_average: float
    wine_sales: float
    wine_pct: float
    rate_of_sale_by_guest_count: float
    average_ticket_time_seconds: float


def load_config(path: Path) -> dict[str, Any]:
    config = json.loads(json.dumps(DEFAULT_CONFIG))
    if path.exists():
        user_config = json.loads(path.read_text(encoding="utf-8"))
        for key, value in user_config.items():
            if isinstance(value, dict) and isinstance(config.get(key), dict):
                config[key].update(value)
            else:
                config[key] = value
    return config


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except TypeError:
        pass
    return str(value).strip() == ""


def to_float(value: Any) -> float:
    if is_blank(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    return float(str(value).replace(",", "").replace("$", "").strip())


def ticket_time_to_seconds(value: Any) -> float:
    if is_blank(value):
        return 0.0
    if isinstance(value, time):
        return (
            value.hour * 3600
            + value.minute * 60
            + value.second
            + value.microsecond / 1_000_000
        )
    if isinstance(value, datetime):
        t = value.time()
        return t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1_000_000
    if isinstance(value, (int, float)):
        number = float(value)
        return number * 86400 if 0 <= number < 1 else number

    text = str(value).strip()
    match = re.match(r"^(\d+):(\d{2}):(\d{2})(?:\.(\d+))?$", text)
    if not match:
        return 0.0
    hours, minutes, seconds, fraction = match.groups()
    micros = float(f"0.{fraction}") if fraction else 0.0
    return int(hours) * 3600 + int(minutes) * 60 + int(seconds) + micros


def parse_date_text(text: str) -> date:
    text = text.strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Could not parse report date: {text}")


def parse_report_date(df: pd.DataFrame, path: Path) -> date:
    for value in df.to_numpy().ravel():
        if isinstance(value, str) and "Date(s):" in value:
            match = re.search(r"Date\(s\):\s*([^\n]+)", value)
            if match:
                date_text = match.group(1).split("-")[-1].strip()
                return parse_date_text(date_text)
    match = re.search(r"(\d{2})-(\d{2})-(\d{4})", path.name)
    if match:
        month, day, year = map(int, match.groups())
        return date(year, month, day) - timedelta(days=FILENAME_DATE_BUSINESS_DATE_OFFSET_DAYS)
    raise ValueError(f"Could not find report date in {path.name}")


def find_header_row(df: pd.DataFrame) -> tuple[int, int, list[int]]:
    for row_index, row in df.iterrows():
        values = ["" if is_blank(v) else str(v).strip() for v in row.tolist()]
        if "Store" in values and "Gross Sales" in values:
            location_col = values.index("Store")
            block_starts = [idx for idx, value in enumerate(values) if value == "Gross Sales"]
            if block_starts:
                return row_index, location_col, block_starts
    raise ValueError("Could not find the metric header row.")


def display_name_for(raw_name: str, config: dict[str, Any]) -> str:
    aliases = config.get("public_name_aliases", {})
    return aliases.get(raw_name, raw_name)


def is_numbered_server_placeholder(name: Any) -> bool:
    text = "" if is_blank(name) else str(name).strip()
    pattern = r"\d+\s+Server\d*(?:\s+Server\d*)*"
    return re.fullmatch(pattern, text, re.IGNORECASE) is not None


def parse_daily_report(path: Path, config: dict[str, Any]) -> list[MetricRecord]:
    df = pd.ExcelFile(path, engine="xlrd").parse("Report(All)", header=None)
    report_date = parse_report_date(df, path)
    header_row, location_col, block_starts = find_header_row(df)
    location_names = set(config["locations"])
    records: list[MetricRecord] = []

    for _, row in df.iloc[header_row + 1 :].iterrows():
        location_value = row.get(location_col)
        if is_blank(location_value):
            continue

        location = str(location_value).strip()
        if location not in location_names:
            continue

        raw_user_name = "" if is_blank(row.get(1)) else str(row.get(1)).strip()

        for start_col in block_starts:
            values = [row.get(start_col + offset) for offset in range(len(METRICS))]
            if all(is_blank(value) for value in values):
                continue

            gross_sales = to_float(values[0])
            guest_count = to_float(values[1])
            wine_sales = to_float(values[3])
            rate = to_float(values[4])
            ticket_seconds = ticket_time_to_seconds(values[5])

            if raw_user_name and gross_sales == 0 and guest_count == 0 and wine_sales == 0:
                continue

            check_average = gross_sales / guest_count if guest_count else 0.0
            wine_pct = wine_sales / gross_sales if gross_sales else 0.0
            records.append(
                MetricRecord(
                    source_file=path.name,
                    report_date=report_date,
                    location=location,
                    raw_user_name=raw_user_name,
                    display_name=display_name_for(raw_user_name, config),
                    is_location_total=raw_user_name == "",
                    gross_sales=gross_sales,
                    guest_count=guest_count,
                    check_average=check_average,
                    wine_sales=wine_sales,
                    wine_pct=wine_pct,
                    rate_of_sale_by_guest_count=rate,
                    average_ticket_time_seconds=ticket_seconds,
                )
            )
            break

    return records


def daily_report_paths(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        return []
    return sorted(path for path in input_dir.glob(DAILY_REPORT_PATTERN) if path.is_file())


def archived_daily_report_paths(archive_dir: Path) -> list[Path]:
    if not archive_dir.exists():
        return []
    return sorted(path for path in archive_dir.rglob(DAILY_REPORT_PATTERN) if path.is_file())


def read_reports_by_path(paths: Iterable[Path], config: dict[str, Any]) -> dict[Path, list[MetricRecord]]:
    records_by_path: dict[Path, list[MetricRecord]] = {}
    for path in paths:
        records = parse_daily_report(path, config)
        if not records:
            raise ValueError(f"No metric rows were found in {path}")
        records_by_path[path] = records
    return records_by_path


def flatten_report_records(records_by_path: dict[Path, list[MetricRecord]]) -> list[MetricRecord]:
    records: list[MetricRecord] = []
    for path_records in records_by_path.values():
        records.extend(path_records)
    return records


def read_all_reports(input_dir: Path, config: dict[str, Any]) -> list[MetricRecord]:
    paths = daily_report_paths(input_dir)
    if not paths:
        raise FileNotFoundError(f"No daily .xls reports found in {input_dir}")

    return flatten_report_records(read_reports_by_path(paths, config))


def is_operating_day(day: date) -> bool:
    return OPERATING_WEEK_START_WEEKDAY <= day.weekday() <= OPERATING_WEEK_END_WEEKDAY


def week_period_for(day: date) -> tuple[date, date]:
    days_since_start = (day.weekday() - OPERATING_WEEK_START_WEEKDAY) % 7
    week_start = day - timedelta(days=days_since_start)
    return week_start, week_start + timedelta(days=OPERATING_WEEK_DAYS - 1)


def active_week_for_paths(records_by_path: dict[Path, list[MetricRecord]]) -> tuple[date, date]:
    groups: dict[tuple[date, date], list[Path]] = defaultdict(list)
    for path, path_records in records_by_path.items():
        report_dates = sorted({record.report_date for record in path_records})
        if len(report_dates) != 1:
            raise ValueError(f"{path.name} contains multiple report dates: {report_dates}")
        groups[week_period_for(report_dates[0])].append(path)

    if not groups:
        raise ValueError("No active daily report files were parsed.")

    if len(groups) > 1:
        lines = [
            "Daily Reports contains files from more than one Tuesday-Sunday operating week.",
            "Move the extra files out of Daily Reports and rerun:",
        ]
        for (_, week_end), paths in sorted(groups.items()):
            lines.append(f"  week-ending-{week_end.isoformat()}:")
            for path in sorted(paths):
                lines.append(f"    {path.name}")
        raise ValueError("\n".join(lines))

    return next(iter(groups))


def selected_public_dates(
    records: Iterable[MetricRecord], week_start: str | None, week_end: str | None
) -> tuple[date, date]:
    all_dates = sorted({record.report_date for record in records})
    if not all_dates:
        raise ValueError("No report dates were found.")
    operating_dates = [report_date for report_date in all_dates if is_operating_day(report_date)]
    if not operating_dates:
        raise ValueError("No Tuesday-Sunday operating report dates were found.")

    if week_start and week_end:
        return date.fromisoformat(week_start), date.fromisoformat(week_end)
    if week_start:
        start = date.fromisoformat(week_start)
        return start, start + timedelta(days=OPERATING_WEEK_DAYS - 1)
    if week_end:
        end = date.fromisoformat(week_end)
        return end - timedelta(days=OPERATING_WEEK_DAYS - 1), end

    return week_period_for(max(operating_dates))


def archive_destination_for(source_path: Path, archive_dir: Path) -> tuple[Path, bool]:
    destination = archive_dir / source_path.name
    if not destination.exists():
        return destination, False

    if filecmp.cmp(source_path, destination, shallow=False):
        return destination, True

    counter = 1
    while True:
        candidate = archive_dir / f"{source_path.stem} ({counter}){source_path.suffix}"
        if not candidate.exists():
            return candidate, False
        if filecmp.cmp(source_path, candidate, shallow=False):
            return candidate, True
        counter += 1


def archive_processed_files(
    source_paths: Iterable[Path], archive_root: Path, week_end: date
) -> list[Path]:
    archive_dir = archive_root / "processed-daily-reports" / f"week-ending-{week_end.isoformat()}"
    archive_dir.mkdir(parents=True, exist_ok=True)

    move_plan: list[tuple[Path, Path, bool]] = []
    for source_path in source_paths:
        destination, already_archived = archive_destination_for(source_path, archive_dir)
        move_plan.append((source_path, destination, already_archived))

    archived_paths: list[Path] = []
    for source_path, destination, already_archived in move_plan:
        if already_archived:
            source_path.unlink()
        else:
            shutil.move(str(source_path), str(destination))
        archived_paths.append(destination)

    return archived_paths


def aggregate_records(
    records: Iterable[MetricRecord], key_fields: tuple[str, ...], extra: dict[str, Any] | None = None
) -> list[dict[str, Any]]:
    groups: dict[tuple[Any, ...], dict[str, Any]] = {}

    for record in records:
        key = tuple(getattr(record, field) for field in key_fields)
        if key not in groups:
            groups[key] = {field: getattr(record, field) for field in key_fields}
            if extra:
                groups[key].update(extra)
            groups[key].update(
                {
                    "gross_sales": 0.0,
                    "guest_count": 0.0,
                    "wine_sales": 0.0,
                    "rate_weighted_sum": 0.0,
                    "rate_weight": 0.0,
                    "ticket_weighted_sum": 0.0,
                    "ticket_weight": 0.0,
                    "active_dates": set(),
                    "source_files": set(),
                }
            )

        group = groups[key]
        group["gross_sales"] += record.gross_sales
        group["guest_count"] += record.guest_count
        group["wine_sales"] += record.wine_sales
        if record.guest_count > 0:
            group["rate_weighted_sum"] += (
                record.rate_of_sale_by_guest_count * record.guest_count
            )
            group["rate_weight"] += record.guest_count
            group["ticket_weighted_sum"] += (
                record.average_ticket_time_seconds * record.guest_count
            )
            group["ticket_weight"] += record.guest_count
        if record.guest_count > 0 or record.gross_sales > 0:
            group["active_dates"].add(record.report_date)
        group["source_files"].add(record.source_file)

    rollups: list[dict[str, Any]] = []
    for group in groups.values():
        gross_sales = group["gross_sales"]
        guest_count = group["guest_count"]
        wine_sales = group["wine_sales"]
        rollup = dict(group)
        rollup["check_average"] = gross_sales / guest_count if guest_count else 0.0
        rollup["wine_pct"] = wine_sales / gross_sales if gross_sales else 0.0
        rollup["rate_of_sale_by_guest_count"] = (
            group["rate_weighted_sum"] / group["rate_weight"] if group["rate_weight"] else 0.0
        )
        rollup["average_ticket_time_seconds"] = (
            group["ticket_weighted_sum"] / group["ticket_weight"]
            if group["ticket_weight"]
            else 0.0
        )
        rollup["active_days"] = len(group["active_dates"])
        rollup["source_days"] = len(group["active_dates"])
        rollup["source_files"] = ", ".join(sorted(group["source_files"]))
        for helper in (
            "rate_weighted_sum",
            "rate_weight",
            "ticket_weighted_sum",
            "ticket_weight",
            "active_dates",
        ):
            rollup.pop(helper, None)
        rollups.append(rollup)

    return rollups


def row_name_haystack(row: dict[str, Any]) -> str:
    return f"{row.get('raw_user_name', '')} {row.get('display_name', '')}".casefold()


def row_matches_name_patterns(row: dict[str, Any], patterns: Iterable[Any]) -> bool:
    haystack = row_name_haystack(row)
    return any(str(pattern).casefold() in haystack for pattern in patterns)


def dashboard_exclusion_patterns(config: dict[str, Any]) -> list[Any]:
    patterns: list[Any] = []
    for key in ("dashboard_exclude_name_contains", "public_exclude_name_contains"):
        for pattern in config.get(key, []):
            if pattern not in patterns:
                patterns.append(pattern)
    return patterns


def dashboard_trend_eligible(row: dict[str, Any], config: dict[str, Any]) -> bool:
    min_guest_count = float(config.get("dashboard_min_guest_count_for_trends", 25))
    min_active_days = int(config.get("dashboard_min_active_days_for_trends", 3))
    return row.get("guest_count", 0) >= min_guest_count and row.get("active_days", 0) >= min_active_days


def public_excluded(row: dict[str, Any], config: dict[str, Any]) -> bool:
    if row["guest_count"] < float(config.get("public_min_guest_count", 1)):
        return True
    if is_numbered_server_placeholder(row.get("raw_user_name")) or is_numbered_server_placeholder(
        row.get("display_name")
    ):
        return True
    return row_matches_name_patterns(row, config.get("public_exclude_name_contains", []))


def dashboard_excluded(row: dict[str, Any], config: dict[str, Any]) -> bool:
    if is_numbered_server_placeholder(row.get("raw_user_name")) or is_numbered_server_placeholder(
        row.get("display_name")
    ):
        return True
    exact_names = {
        str(value).strip().casefold()
        for value in config.get("dashboard_exclude_exact_names", [])
    }
    if str(row.get("display_name", "")).strip().casefold() in exact_names:
        return True
    return row_matches_name_patterns(row, dashboard_exclusion_patterns(config))


def format_date_range(start: date, end: date) -> str:
    if start == end:
        return start.strftime("%m/%d/%Y")
    return f"{start:%m/%d/%Y} - {end:%m/%d/%Y}"


def duration_fraction(seconds: float) -> float:
    return seconds / 86400 if seconds else 0.0


def style_public_sheet(ws, last_row: int) -> None:
    header_fill = PatternFill("solid", fgColor="E1E1E1")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 78
    widths = {"A": 18, "B": 28, "C": 13, "D": 13, "E": 15, "F": 12, "G": 16}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for hidden_col in ("C", "D", "F"):
        ws.column_dimensions[hidden_col].hidden = True

    ws["A1"].font = Font(name="Calibri", size=12, color="68696C")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for cell in ws[2]:
        cell.fill = header_fill
        cell.font = Font(name="Calibri", size=12, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row in ws.iter_rows(min_row=3, max_row=last_row, min_col=1, max_col=7):
        for cell in row:
            cell.font = Font(name="Calibri", size=12, bold=cell.row == 3)
            cell.border = border
        row[2].number_format = "$#,##0.00"
        row[3].number_format = "#,##0"
        row[4].number_format = "$#,##0.00"
        row[5].number_format = "$#,##0.00"
        row[6].number_format = "0.00%"

    if last_row >= 2:
        ws.auto_filter.ref = f"A2:G{last_row}"
        ws.freeze_panes = "A3"


def apply_public_metric_highlights(ws, last_row: int) -> None:
    first_server_row = 4
    if last_row < first_server_row:
        return

    light_green = PatternFill("solid", fgColor="92D050")
    light_yellow = PatternFill("solid", fgColor="FFFF00")
    light_red = PatternFill("solid", fgColor="F4CCCC")

    rows = list(range(first_server_row, last_row + 1))
    check_values = [
        (row, ws.cell(row=row, column=5).value or 0)
        for row in rows
    ]
    wine_values = []
    for row in rows:
        gross_sales = ws.cell(row=row, column=3).value or 0
        wine_sales = ws.cell(row=row, column=6).value or 0
        wine_pct = wine_sales / gross_sales if gross_sales else 0
        wine_values.append((row, wine_pct))

    top_check_row = max(check_values, key=lambda item: item[1])[0]
    top_wine_row = max(wine_values, key=lambda item: item[1])[0]
    bottom_check_rows = [
        row for row, _ in sorted(check_values, key=lambda item: (item[1], item[0]))[:3]
    ]

    for row in bottom_check_rows:
        ws.cell(row=row, column=5).fill = light_red
    ws.cell(row=top_check_row, column=5).fill = light_green
    ws.cell(row=top_wine_row, column=7).fill = light_yellow


def write_report_definition(ws, location: str, start: date, end: date) -> None:
    bullet = "\u25cf"
    filter_text = (
        "FILTERS\n"
        f"  {bullet}  Date(s):  {format_date_range(start, end)}\n"
        f"  {bullet}  Location Type:  Store\n"
        f"  {bullet}  Locations:  {location}"
    )
    rows = [
        ["Daily Report - TM (Auto-Run)"],
        [filter_text],
        [],
        ["Column", "Definition"],
        ["Gross Sales", "KPI: GrossAmount"],
        ["Guest Count", "KPI: GuestCount"],
        ["Check Average", "KPI: GrossGuestAvg"],
        ["Wine Sales", "KPI: MenuPrice\nMenuCategory: Banquet, Wine - Glass, Wine Bottle, Wine Glass"],
        ["Rate of Sale by Guest Count", "KPI: RateOfSaleGuestCount"],
        ["Average Ticket Time", "KPI: AvgTicketTime"],
    ]
    for row in rows:
        ws.append(row)

    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 76
    ws.row_dimensions[2].height = 78
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["B8"].alignment = Alignment(wrap_text=True)
    for cell in ws[4]:
        cell.fill = PatternFill("solid", fgColor="E1E1E1")
        cell.font = Font(bold=True)


def write_public_workbook(
    location: str,
    selected_records: list[MetricRecord],
    output_dir: Path,
    config: dict[str, Any],
    public_start: date,
    public_end: date,
) -> Path:
    actual_start = min(record.report_date for record in selected_records)
    actual_end = max(record.report_date for record in selected_records)
    short_code = config["locations"][location]["short_code"]

    location_records = [record for record in selected_records if record.location == location]
    totals = [record for record in location_records if record.is_location_total]
    server_records = [record for record in location_records if not record.is_location_total]

    total_rows = aggregate_records(totals, ("location", "display_name", "raw_user_name"))
    public_rows = aggregate_records(server_records, ("location", "display_name", "raw_user_name"))
    public_rows = [row for row in public_rows if not public_excluded(row, config)]
    public_rows.sort(key=lambda row: (-row["check_average"], row["display_name"]))

    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    ws = wb.active
    ws.title = "Report(All)"
    ws.merge_cells("A1:F1")

    bullet = "\u25cf"
    ws["A1"] = (
        "FILTERS\n"
        f"  {bullet}  Week:  {format_date_range(public_start, public_end)} ({OPERATING_WEEK_LABEL})\n"
        f"  {bullet}  Source Date(s):  {format_date_range(actual_start, actual_end)}\n"
        f"  {bullet}  Location Type:  Store\n"
        f"  {bullet}  Locations:  {location}"
    )
    headers = ["Store", "", "Gross Sales", "Guest Count", "Check Average", "Wine Sales", "Wine %"]
    for col_index, header in enumerate(headers, start=1):
        ws.cell(row=2, column=col_index, value=header)

    rows_to_write = []
    if total_rows:
        rows_to_write.append(total_rows[0])
    rows_to_write.extend(public_rows)

    for row_index, row in enumerate(rows_to_write, start=3):
        ws.cell(row=row_index, column=1, value=location)
        ws.cell(row=row_index, column=2, value=row["display_name"])
        ws.cell(row=row_index, column=3, value=row["gross_sales"])
        ws.cell(row=row_index, column=4, value=row["guest_count"])
        ws.cell(row=row_index, column=5, value=row["check_average"])
        ws.cell(row=row_index, column=6, value=row["wine_sales"])
        ws.cell(row=row_index, column=7, value=row["wine_pct"])

    last_row = max(2, len(rows_to_write) + 2)
    style_public_sheet(ws, last_row)
    apply_public_metric_highlights(ws, last_row)
    write_report_definition(wb.create_sheet("Report Definition"), location, public_start, public_end)

    output_path = output_dir / f"Check_Wine_{short_code}{public_end:%m%d%y}.xlsx"
    wb.save(output_path)
    return output_path


def write_table_sheet(
    wb: Workbook,
    title: str,
    headers: list[str],
    rows: list[list[Any]],
    table_name: str,
    widths: dict[str, float] | None = None,
    index: int | None = None,
):
    ws = wb.create_sheet(title, index=index)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"
    ws.cell(row=1, column=1, value=title)
    ws.cell(row=1, column=1).font = Font(size=15, bold=True, color="FFFFFF")
    ws.cell(row=1, column=1).fill = PatternFill("solid", fgColor="7A1E1E")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(headers)))

    for col_index, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_index, value=header)
        cell.fill = PatternFill("solid", fgColor="E1E1E1")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_index, row in enumerate(rows, start=4):
        for col_index, value in enumerate(row, start=1):
            cell = ws.cell(row=row_index, column=col_index, value=value)
            header = headers[col_index - 1]
            if "Date" in header or header in {"Week Start", "Week End", "Latest Week End"}:
                cell.number_format = "m/d/yyyy"
            elif "Composite Score" in header or "Rank Movement" in header:
                cell.number_format = "0.0"
            elif "Rank" in header:
                cell.number_format = "#,##0"
            elif "Ticket Time Change (Min)" in header:
                cell.number_format = "0.0"
            elif "Check Average Change" in header:
                cell.number_format = "$#,##0.00"
            elif "Wine % Change" in header:
                cell.number_format = "0.00%"
            elif "Rate Change" in header:
                cell.number_format = "0.00"
            elif "Wine %" in header:
                cell.number_format = "0.00%"
            elif "Rate of Sale" in header:
                cell.number_format = "0.00"
            elif "Ticket Time" in header:
                cell.number_format = "[h]:mm:ss"
            elif (
                "Gross Sales" in header
                or "Wine Sales" in header
                or ("Check Average" in header and "Rank" not in header)
            ):
                cell.number_format = "$#,##0.00"
            elif "Guest" in header or "Days" in header or "Weeks" in header:
                cell.number_format = "#,##0"

    if rows:
        ref = f"A3:{get_column_letter(len(headers))}{len(rows) + 3}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)

    default_widths = {
        "A": 14,
        "B": 14,
        "C": 20,
        "D": 24,
        "E": 24,
        "F": 18,
        "G": 12,
        "H": 16,
        "I": 14,
        "J": 12,
        "K": 20,
        "L": 18,
        "M": 12,
        "N": 12,
        "O": 18,
        "P": 18,
    }
    for index in range(1, len(headers) + 1):
        col = get_column_letter(index)
        default_widths.setdefault(col, min(max(len(headers[index - 1]) + 2, 12), 24))
    default_widths.update(widths or {})
    for col, width in default_widths.items():
        ws.column_dimensions[col].width = width
    return ws


def weekly_rollups(records: list[MetricRecord]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    grouped: dict[tuple[date, date], list[MetricRecord]] = defaultdict(list)
    for record in records:
        if not is_operating_day(record.report_date):
            continue
        grouped[week_period_for(record.report_date)].append(record)

    server_rows: list[dict[str, Any]] = []
    location_rows: list[dict[str, Any]] = []
    for (week_start, week_end), period_records in sorted(grouped.items()):
        servers = [record for record in period_records if not record.is_location_total]
        locations = [record for record in period_records if record.is_location_total]
        server_rows.extend(
            aggregate_records(
                servers,
                ("location", "raw_user_name", "display_name"),
                {"week_start": week_start, "week_end": week_end},
            )
        )
        location_rows.extend(
            aggregate_records(
                locations,
                ("location",),
                {"week_start": week_start, "week_end": week_end},
            )
        )
    return server_rows, location_rows


def assign_rank(
    rows: list[dict[str, Any]],
    value_field: str,
    rank_field: str,
    *,
    higher_is_better: bool,
    min_guest_count: int,
) -> None:
    eligible = [
        row
        for row in rows
        if row.get("guest_count", 0) >= min_guest_count and row.get(value_field) is not None
    ]
    eligible.sort(
        key=lambda row: row[value_field],
        reverse=higher_is_better,
    )

    previous_value: float | None = None
    previous_rank = 0
    for index, row in enumerate(eligible, start=1):
        value = row[value_field]
        rank = previous_rank if previous_value == value else index
        row[rank_field] = rank
        previous_rank = rank
        previous_value = value

    for row in rows:
        row.setdefault(rank_field, None)


def weekly_server_rank_rows(
    weekly_server_rows: list[dict[str, Any]],
    min_guest_count: int,
) -> list[dict[str, Any]]:
    groups: dict[tuple[date, date, str], list[dict[str, Any]]] = defaultdict(list)
    for row in weekly_server_rows:
        ranked = dict(row)
        groups[(ranked["week_start"], ranked["week_end"], ranked["location"])].append(ranked)

    ranked_rows: list[dict[str, Any]] = []
    for group_rows in groups.values():
        assign_rank(
            group_rows,
            "check_average",
            "check_average_rank",
            higher_is_better=True,
            min_guest_count=min_guest_count,
        )
        assign_rank(
            group_rows,
            "wine_pct",
            "wine_pct_rank",
            higher_is_better=True,
            min_guest_count=min_guest_count,
        )
        assign_rank(
            group_rows,
            "rate_of_sale_by_guest_count",
            "rate_rank",
            higher_is_better=False,
            min_guest_count=min_guest_count,
        )
        assign_rank(
            group_rows,
            "average_ticket_time_seconds",
            "ticket_time_rank",
            higher_is_better=False,
            min_guest_count=min_guest_count,
        )
        ranked_rows.extend(group_rows)

    ranked_rows.sort(
        key=lambda row: (
            row["week_end"],
            row["location"],
            row.get("check_average_rank") or 9999,
            row["display_name"],
        )
    )
    return ranked_rows


def trend_note(
    check_change: float | None,
    wine_change: float | None,
    rate_change: float | None,
    ticket_change_minutes: float | None,
) -> str:
    if check_change is None:
        return "No prior week"

    improved = 0
    declined = 0
    for value, higher_is_better in (
        (check_change, True),
        (wine_change, True),
        (rate_change, False),
        (ticket_change_minutes, False),
    ):
        if value is None or abs(value) < 0.000001:
            continue
        if (higher_is_better and value > 0) or ((not higher_is_better) and value < 0):
            improved += 1
        else:
            declined += 1

    if improved >= 2 and improved > declined:
        return "Improving"
    if declined >= 2 and declined > improved:
        return "Watch"
    return "Mixed"


def metric_delta(current: dict[str, Any], previous: dict[str, Any] | None, field: str) -> float | None:
    if previous is None:
        return None
    return current[field] - previous[field]


def rank_movement(current_rank: Any, previous_rank: Any) -> float | None:
    if current_rank is None or previous_rank is None:
        return None
    return float(previous_rank) - float(current_rank)


def star_score_component(
    change: float | None, *, lower_is_better: bool, strong_threshold: float
) -> int:
    if change is None or abs(change) < 0.000001:
        return 0
    directional_change = -change if lower_is_better else change
    if directional_change >= strong_threshold:
        return 2
    if directional_change > 0:
        return 1
    if directional_change <= -strong_threshold:
        return -2
    return -1


def format_change(change: float | None, kind: str) -> str:
    if change is None:
        return ""
    sign = "+" if change > 0 else "-" if change < 0 else ""
    value = abs(change)
    if kind == "currency":
        return f"{sign}${value:,.2f}"
    if kind == "pct_points":
        return f"{sign}{value * 100:.1f} pts"
    if kind == "minutes":
        return f"{sign}{value:.1f} min"
    if kind == "rank":
        return f"{sign}{value:.1f}"
    return f"{sign}{value:.2f}"


def average_rank_movement(row: dict[str, Any]) -> float | None:
    movements = [
        row.get("check_average_rank_movement"),
        row.get("wine_pct_rank_movement"),
        row.get("rate_rank_movement"),
        row.get("ticket_time_rank_movement"),
    ]
    available = [movement for movement in movements if movement is not None]
    if not available:
        return None
    return sum(available) / len(available)


def server_week_trend_rows(ranked_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in ranked_rows:
        groups[(row["location"], row["raw_user_name"], row["display_name"])].append(row)

    trend_rows: list[dict[str, Any]] = []
    for (_, _, _), rows in groups.items():
        rows.sort(key=lambda row: row["week_end"])
        previous: dict[str, Any] | None = None
        for row in rows:
            check_change = metric_delta(row, previous, "check_average")
            wine_change = metric_delta(row, previous, "wine_pct")
            rate_change = metric_delta(row, previous, "rate_of_sale_by_guest_count")
            ticket_change_minutes = (
                (row["average_ticket_time_seconds"] - previous["average_ticket_time_seconds"]) / 60
                if previous
                else None
            )
            rank_fields = (
                ("check_average_rank", "check_average_rank_movement"),
                ("wine_pct_rank", "wine_pct_rank_movement"),
                ("rate_rank", "rate_rank_movement"),
                ("ticket_time_rank", "ticket_time_rank_movement"),
            )
            rank_values: dict[str, Any] = {}
            for rank_field, movement_field in rank_fields:
                prior_field = f"prior_{rank_field}"
                prior_rank = previous.get(rank_field) if previous else None
                rank_values[prior_field] = prior_rank
                rank_values[movement_field] = rank_movement(row.get(rank_field), prior_rank)
            trend_rows.append(
                {
                    **row,
                    "check_average_change": check_change,
                    "wine_pct_change": wine_change,
                    "rate_change": rate_change,
                    "ticket_time_change_minutes": ticket_change_minutes,
                    **rank_values,
                    "trend_note": trend_note(
                        check_change,
                        wine_change,
                        rate_change,
                        ticket_change_minutes,
                    ),
                }
            )
            previous = row

    trend_rows.sort(
        key=lambda row: (
            row["week_end"],
            row["location"],
            row.get("check_average_rank") or 9999,
            row["display_name"],
        )
    )
    return trend_rows


def star_classification(score: float) -> str:
    if score >= 2:
        return "Rising Star"
    if score <= -2:
        return "Falling Star"
    return "Stable"


def star_why(row: dict[str, Any], components: list[tuple[str, int, str]]) -> str:
    drivers = [
        label
        for label, score, _ in sorted(
            components,
            key=lambda item: (abs(item[1]), item[2]),
            reverse=True,
        )
        if score
    ]
    if not drivers:
        return "Minimal week-over-week movement"
    return "; ".join(drivers[:4])


def server_star_rows(
    server_trend_detail_rows: list[dict[str, Any]], config: dict[str, Any]
) -> list[dict[str, Any]]:
    rows_with_prior = [
        row
        for row in server_trend_detail_rows
        if row.get("check_average_change") is not None
    ]
    if not rows_with_prior:
        return []

    latest_week_end = max(row["week_end"] for row in rows_with_prior)
    star_rows: list[dict[str, Any]] = []
    for row in rows_with_prior:
        if row["week_end"] != latest_week_end:
            continue
        if dashboard_excluded(row, config) or not dashboard_trend_eligible(row, config):
            continue

        metric_components = [
            (
                "Check average",
                star_score_component(
                    row.get("check_average_change"),
                    lower_is_better=False,
                    strong_threshold=5.0,
                ),
                f"Check avg {format_change(row.get('check_average_change'), 'currency')}",
                "check",
            ),
            (
                "Wine %",
                star_score_component(
                    row.get("wine_pct_change"),
                    lower_is_better=False,
                    strong_threshold=0.01,
                ),
                f"Wine {format_change(row.get('wine_pct_change'), 'pct_points')}",
                "wine",
            ),
            (
                "Rate",
                star_score_component(
                    row.get("rate_change"),
                    lower_is_better=True,
                    strong_threshold=0.25,
                ),
                f"Rate {format_change(row.get('rate_change'), 'number')}",
                "rate",
            ),
            (
                "Ticket time",
                star_score_component(
                    row.get("ticket_time_change_minutes"),
                    lower_is_better=True,
                    strong_threshold=5.0,
                ),
                f"Ticket {format_change(row.get('ticket_time_change_minutes'), 'minutes')}",
                "ticket",
            ),
        ]
        avg_rank_move = average_rank_movement(row)
        rank_score = 0
        if avg_rank_move is not None:
            if avg_rank_move >= 5:
                rank_score = 2
            elif avg_rank_move >= 2:
                rank_score = 1
            elif avg_rank_move <= -5:
                rank_score = -2
            elif avg_rank_move <= -2:
                rank_score = -1
        score = sum(component[1] for component in metric_components) + rank_score
        components = [(label, component_score, sort_key) for _, component_score, label, sort_key in metric_components]
        components.append(
            (
                f"Rank {format_change(avg_rank_move, 'rank')}",
                rank_score,
                "rank",
            )
        )
        star_rows.append(
            {
                "category": star_classification(score),
                "composite_score": score,
                "week_start": row["week_start"],
                "week_end": row["week_end"],
                "location": row["location"],
                "raw_user_name": row["raw_user_name"],
                "display_name": row["display_name"],
                "guest_count": row["guest_count"],
                "active_days": row["active_days"],
                "check_average": row["check_average"],
                "check_average_change": row["check_average_change"],
                "wine_pct": row["wine_pct"],
                "wine_pct_change": row["wine_pct_change"],
                "rate_of_sale_by_guest_count": row["rate_of_sale_by_guest_count"],
                "rate_change": row["rate_change"],
                "average_ticket_time_seconds": row["average_ticket_time_seconds"],
                "ticket_time_change_minutes": row["ticket_time_change_minutes"],
                "average_rank_movement": avg_rank_move,
                "why": star_why(row, components),
            }
        )

    category_order = {"Rising Star": 0, "Falling Star": 1, "Stable": 2}
    star_rows.sort(
        key=lambda row: (
            category_order[row["category"]],
            -row["composite_score"] if row["category"] == "Rising Star" else row["composite_score"],
            row["location"],
            row["display_name"],
        )
    )
    return star_rows


def compact_money(value: float | None, *, signed: bool = False) -> str:
    if value is None:
        return ""
    sign = ""
    if signed:
        sign = "+" if value > 0 else "-" if value < 0 else ""
    return f"{sign}${abs(value):,.0f}"


def compact_count(value: float | None, *, signed: bool = False) -> str:
    if value is None:
        return ""
    sign = "+" if signed and value > 0 else ""
    return f"{sign}{value:,.0f}"


def compact_pct_points(value: float | None) -> str:
    if value is None:
        return ""
    sign = "+" if value > 0 else "-" if value < 0 else ""
    return f"{sign}{abs(value) * 100:.1f} pts"


def compact_minutes(value: float | None) -> str:
    if value is None:
        return ""
    sign = "+" if value > 0 else "-" if value < 0 else ""
    return f"{sign}{abs(value):.1f} min"


def star_action(row: dict[str, Any]) -> tuple[str, str, str]:
    if row["category"] == "Falling Star":
        priority = "High" if row["composite_score"] <= -5 else "Medium"
        return priority, "Coach", "Review recent shifts and coach the drivers shown in Evidence."
    if row["category"] == "Rising Star":
        priority = "Recognize" if row["composite_score"] >= 5 else "Share"
        return priority, "Recognize", "Recognize the improvement and ask what changed so the tactic can be shared."
    return "Monitor", "Monitor", "Keep watching next week before acting."


def store_status(row: dict[str, Any]) -> tuple[str, str, str]:
    gross_down = (row.get("gross_sales_change") or 0) < 0
    guests_down = (row.get("guest_count_change") or 0) < 0
    check_down = (row.get("check_average_change") or 0) < 0
    wine_down = (row.get("wine_pct_change") or 0) < 0
    ticket_slower = (row.get("ticket_time_change_minutes") or 0) > 0
    negative_count = sum([gross_down, guests_down, check_down, wine_down, ticket_slower])

    if gross_down and guests_down:
        return "High", "Traffic Watch", "Sales and guest count both fell; review traffic, staffing, and event calendar."
    if check_down and wine_down:
        return "Medium", "Upsell Watch", "Guests held better than spend; coach check average and wine attachment."
    if ticket_slower and negative_count >= 2:
        return "Medium", "Service Watch", "Ticket time worsened with other declines; review service flow."
    if negative_count >= 2:
        return "Medium", "Mixed Watch", "Multiple metrics moved the wrong way; review manager notes."
    return "Monitor", "Stable / Mixed", "No urgent store action; monitor next run."


def trend_evidence(row: dict[str, Any]) -> str:
    return (
        f"Sales {compact_money(row.get('gross_sales_change'), signed=True)}; "
        f"guests {compact_count(row.get('guest_count_change'), signed=True)}; "
        f"check {compact_money(row.get('check_average_change'), signed=True)}; "
        f"wine {compact_pct_points(row.get('wine_pct_change'))}; "
        f"ticket {compact_minutes(row.get('ticket_time_change_minutes'))}"
    )


def build_action_board_rows(
    star_rows: list[dict[str, Any]],
    store_trend_summary_rows: list[dict[str, Any]],
    group_trend_summary_rows: list[dict[str, Any]],
    weekly_location_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sort_order = {
        "High": 0,
        "Medium": 1,
        "Recognize": 2,
        "Share": 3,
        "Monitor": 4,
        "Review": 5,
    }

    for star in sorted(
        (row for row in star_rows if row["category"] == "Falling Star"),
        key=lambda row: (row["composite_score"], row["location"], row["display_name"]),
    )[:6]:
        priority, action, follow_up = star_action(star)
        rows.append(
            {
                "priority": priority,
                "action": action,
                "location": star["location"],
                "subject": star["display_name"],
                "signal": "Falling Star",
                "impact": f"Score {star['composite_score']:+.0f}",
                "evidence": star["why"],
                "recommended_follow_up": follow_up,
                "week_end": star["week_end"],
                "guest_count": star["guest_count"],
                "active_days": star["active_days"],
                "_sort": sort_order[priority],
            }
        )

    for star in sorted(
        (row for row in star_rows if row["category"] == "Rising Star"),
        key=lambda row: (-row["composite_score"], row["location"], row["display_name"]),
    )[:6]:
        priority, action, follow_up = star_action(star)
        rows.append(
            {
                "priority": priority,
                "action": action,
                "location": star["location"],
                "subject": star["display_name"],
                "signal": "Rising Star",
                "impact": f"Score {star['composite_score']:+.0f}",
                "evidence": star["why"],
                "recommended_follow_up": follow_up,
                "week_end": star["week_end"],
                "guest_count": star["guest_count"],
                "active_days": star["active_days"],
                "_sort": sort_order[priority],
            }
        )

    for store in store_trend_summary_rows:
        priority, signal, follow_up = store_status(store)
        rows.append(
            {
                "priority": priority,
                "action": "Store Review",
                "location": store["location"],
                "subject": store["location"],
                "signal": signal,
                "impact": trend_evidence(store),
                "evidence": f"Latest week {store['latest_week_end']:%m/%d/%Y}",
                "recommended_follow_up": follow_up,
                "week_end": store["latest_week_end"],
                "guest_count": store["latest_guest_count"],
                "active_days": None,
                "_sort": sort_order[priority] + 0.5,
            }
        )

    for group in group_trend_summary_rows:
        priority, signal, follow_up = store_status(group)
        rows.append(
            {
                "priority": priority,
                "action": "Group Review",
                "location": group.get("group", "All Stores"),
                "subject": group.get("group", "All Stores"),
                "signal": signal,
                "impact": trend_evidence(group),
                "evidence": f"Latest week {group['latest_week_end']:%m/%d/%Y}",
                "recommended_follow_up": follow_up,
                "week_end": group["latest_week_end"],
                "guest_count": group["latest_guest_count"],
                "active_days": None,
                "_sort": sort_order[priority] + 0.25,
            }
        )

    for status, location, week_end, detail in data_quality_warning_rows(weekly_location_rows):
        if status == "OK":
            continue
        rows.append(
            {
                "priority": "Review",
                "action": "Data Quality",
                "location": location,
                "subject": location,
                "signal": status,
                "impact": detail,
                "evidence": "Short weeks can distort trends and star movement.",
                "recommended_follow_up": "Confirm whether missing source days are expected before coaching from that week.",
                "week_end": week_end,
                "guest_count": None,
                "active_days": None,
                "_sort": sort_order["Review"],
            }
        )

    rows.sort(key=lambda row: (row["_sort"], row["action"], row["location"], row["subject"]))
    return rows


def weekly_metric_trend_rows(
    weekly_rows: list[dict[str, Any]], identity_fields: tuple[str, ...]
) -> list[dict[str, Any]]:
    groups: dict[tuple[Any, ...], list[dict[str, Any]]] = defaultdict(list)
    for row in weekly_rows:
        groups[tuple(row[field] for field in identity_fields)].append(row)

    trend_rows: list[dict[str, Any]] = []
    for rows in groups.values():
        rows.sort(key=lambda row: row["week_end"])
        previous: dict[str, Any] | None = None
        for row in rows:
            check_change = metric_delta(row, previous, "check_average")
            wine_change = metric_delta(row, previous, "wine_pct")
            rate_change = metric_delta(row, previous, "rate_of_sale_by_guest_count")
            ticket_change_minutes = (
                (row["average_ticket_time_seconds"] - previous["average_ticket_time_seconds"]) / 60
                if previous
                else None
            )
            trend_rows.append(
                {
                    **row,
                    "prior_week_end": previous["week_end"] if previous else None,
                    "gross_sales_change": metric_delta(row, previous, "gross_sales"),
                    "guest_count_change": metric_delta(row, previous, "guest_count"),
                    "check_average_change": check_change,
                    "wine_sales_change": metric_delta(row, previous, "wine_sales"),
                    "wine_pct_change": wine_change,
                    "rate_change": rate_change,
                    "ticket_time_change_minutes": ticket_change_minutes,
                    "trend_note": trend_note(
                        check_change,
                        wine_change,
                        rate_change,
                        ticket_change_minutes,
                    ),
                }
            )
            previous = row

    trend_rows.sort(key=lambda row: (row["week_end"],) + tuple(row[field] for field in identity_fields))
    return trend_rows


def group_weekly_rows(records: list[MetricRecord]) -> list[dict[str, Any]]:
    grouped: dict[tuple[date, date], list[MetricRecord]] = defaultdict(list)
    for record in records:
        if record.is_location_total and is_operating_day(record.report_date):
            grouped[week_period_for(record.report_date)].append(record)

    rows: list[dict[str, Any]] = []
    for (week_start, week_end), period_records in sorted(grouped.items()):
        rows.extend(
            aggregate_records(
                period_records,
                (),
                {"week_start": week_start, "week_end": week_end, "group": "All Stores"},
            )
        )
    return rows


def weekly_metric_summary_rows(
    weekly_rows: list[dict[str, Any]], identity_fields: tuple[str, ...]
) -> list[dict[str, Any]]:
    groups: dict[tuple[Any, ...], list[dict[str, Any]]] = defaultdict(list)
    for row in weekly_rows:
        groups[tuple(row[field] for field in identity_fields)].append(row)

    summary: list[dict[str, Any]] = []
    for key, rows in groups.items():
        rows.sort(key=lambda row: row["week_end"])
        total_gross = sum(row["gross_sales"] for row in rows)
        total_guests = sum(row["guest_count"] for row in rows)
        total_wine = sum(row["wine_sales"] for row in rows)
        rate_weighted = sum(row["rate_of_sale_by_guest_count"] * row["guest_count"] for row in rows)
        ticket_weighted = sum(row["average_ticket_time_seconds"] * row["guest_count"] for row in rows)
        latest = rows[-1]
        prior = rows[-2] if len(rows) > 1 else None
        row = {field: value for field, value in zip(identity_fields, key)}
        row.update(
            {
                "weeks_tracked": len(rows),
                "active_days_tracked": sum(item["active_days"] for item in rows),
                "source_days_tracked": sum(item["source_days"] for item in rows),
                "total_gross_sales": total_gross,
                "total_guest_count": total_guests,
                "weighted_check_average": total_gross / total_guests if total_guests else 0.0,
                "total_wine_sales": total_wine,
                "overall_wine_pct": total_wine / total_gross if total_gross else 0.0,
                "weighted_rate": rate_weighted / total_guests if total_guests else 0.0,
                "weighted_ticket_time_seconds": ticket_weighted / total_guests if total_guests else 0.0,
                "first_week_end": rows[0]["week_end"],
                "latest_week_end": latest["week_end"],
                "prior_week_end": prior["week_end"] if prior else None,
                "latest_gross_sales": latest["gross_sales"],
                "prior_gross_sales": prior["gross_sales"] if prior else None,
                "gross_sales_change": metric_delta(latest, prior, "gross_sales"),
                "latest_guest_count": latest["guest_count"],
                "prior_guest_count": prior["guest_count"] if prior else None,
                "guest_count_change": metric_delta(latest, prior, "guest_count"),
                "latest_check_average": latest["check_average"],
                "prior_check_average": prior["check_average"] if prior else None,
                "check_average_change": metric_delta(latest, prior, "check_average"),
                "latest_wine_pct": latest["wine_pct"],
                "prior_wine_pct": prior["wine_pct"] if prior else None,
                "wine_pct_change": metric_delta(latest, prior, "wine_pct"),
                "latest_rate": latest["rate_of_sale_by_guest_count"],
                "prior_rate": prior["rate_of_sale_by_guest_count"] if prior else None,
                "rate_change": metric_delta(latest, prior, "rate_of_sale_by_guest_count"),
                "latest_ticket_time_seconds": latest["average_ticket_time_seconds"],
                "prior_ticket_time_seconds": (
                    prior["average_ticket_time_seconds"] if prior else None
                ),
                "ticket_time_change_minutes": (
                    (latest["average_ticket_time_seconds"] - prior["average_ticket_time_seconds"]) / 60
                    if prior
                    else None
                ),
            }
        )
        summary.append(row)

    summary.sort(key=lambda row: tuple(row[field] for field in identity_fields))
    return summary


def trend_summary_rows(weekly_server_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in weekly_server_rows:
        groups[(row["location"], row["raw_user_name"], row["display_name"])].append(row)

    summary: list[dict[str, Any]] = []
    for (location, raw_name, display_name), rows in groups.items():
        rows.sort(key=lambda row: row["week_end"])
        total_gross = sum(row["gross_sales"] for row in rows)
        total_guests = sum(row["guest_count"] for row in rows)
        total_wine = sum(row["wine_sales"] for row in rows)
        rate_weighted = sum(row["rate_of_sale_by_guest_count"] * row["guest_count"] for row in rows)
        ticket_weighted = sum(row["average_ticket_time_seconds"] * row["guest_count"] for row in rows)
        latest = rows[-1]
        prior = rows[-2] if len(rows) > 1 else None
        summary.append(
            {
                "location": location,
                "raw_user_name": raw_name,
                "display_name": display_name,
                "weeks_tracked": len(rows),
                "active_days_tracked": sum(row["active_days"] for row in rows),
                "total_gross_sales": total_gross,
                "total_guest_count": total_guests,
                "weighted_check_average": total_gross / total_guests if total_guests else 0.0,
                "total_wine_sales": total_wine,
                "overall_wine_pct": total_wine / total_gross if total_gross else 0.0,
                "weighted_rate": rate_weighted / total_guests if total_guests else 0.0,
                "weighted_ticket_time_seconds": ticket_weighted / total_guests if total_guests else 0.0,
                "first_week_end": rows[0]["week_end"],
                "latest_week_end": latest["week_end"],
                "prior_week_end": prior["week_end"] if prior else None,
                "latest_check_average": latest["check_average"],
                "prior_check_average": prior["check_average"] if prior else None,
                "check_average_change": (
                    latest["check_average"] - prior["check_average"] if prior else None
                ),
                "latest_wine_pct": latest["wine_pct"],
                "prior_wine_pct": prior["wine_pct"] if prior else None,
                "wine_pct_change": latest["wine_pct"] - prior["wine_pct"] if prior else None,
                "latest_rate": latest["rate_of_sale_by_guest_count"],
                "prior_rate": prior["rate_of_sale_by_guest_count"] if prior else None,
                "rate_change": (
                    latest["rate_of_sale_by_guest_count"] - prior["rate_of_sale_by_guest_count"]
                    if prior
                    else None
                ),
                "latest_ticket_time_seconds": latest["average_ticket_time_seconds"],
                "prior_ticket_time_seconds": (
                    prior["average_ticket_time_seconds"] if prior else None
                ),
                "ticket_time_change_minutes": (
                    (latest["average_ticket_time_seconds"] - prior["average_ticket_time_seconds"]) / 60
                    if prior
                    else None
                ),
            }
        )

    summary.sort(key=lambda row: (row["location"], -row["weighted_check_average"], row["display_name"]))
    return summary


def style_title(ws, title: str, end_col: int) -> None:
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value=title)
    ws.cell(row=1, column=1).font = Font(size=16, bold=True, color="FFFFFF")
    ws.cell(row=1, column=1).fill = PatternFill("solid", fgColor="7A1E1E")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)


def style_section_header(ws, row: int, start_col: int, end_col: int, label: str) -> None:
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = PatternFill("solid", fgColor="E1E1E1")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.cell(row=row, column=start_col, value=label)


def soft_fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def priority_fill(value: Any) -> PatternFill | None:
    text = str(value or "").casefold()
    if text in {"high", "coach", "falling star", "traffic watch"}:
        return soft_fill("F4CCCC")
    if text in {"medium", "store review", "group review", "upsell watch", "service watch", "mixed watch"}:
        return soft_fill("FFF2CC")
    if text in {"recognize", "share", "rising star"}:
        return soft_fill("D9EAD3")
    if text in {"review", "data quality", "short week"}:
        return soft_fill("D9EAF7")
    if text in {"monitor", "stable / mixed"}:
        return soft_fill("EDEDED")
    return None


def apply_dashboard_number_formats(ws, row_start: int, row_end: int) -> None:
    for row in range(row_start, row_end + 1):
        for col in range(1, ws.max_column + 1):
            header = str(ws.cell(row=row_start - 1, column=col).value or "")
            cell = ws.cell(row=row, column=col)
            if "Date" in header or "Week End" in header:
                cell.number_format = "m/d/yyyy"
            elif "Composite Score" in header or "Rank Movement" in header:
                cell.number_format = "0.0"
            elif "Wine %" in header:
                cell.number_format = "0.00%"
            elif "Rate" in header:
                cell.number_format = "0.00"
            elif "Ticket Time Change (Min)" in header:
                cell.number_format = "0.0"
            elif "Ticket Time" in header:
                cell.number_format = "[h]:mm:ss"
            elif "Gross Sales" in header or "Check Average" in header or "Wine Sales" in header:
                cell.number_format = "$#,##0.00"
            elif "Guest" in header or "Rank" in header:
                cell.number_format = "#,##0"


def data_quality_warning_rows(weekly_location_rows: list[dict[str, Any]]) -> list[list[Any]]:
    short_rows = [
        row
        for row in sorted(weekly_location_rows, key=lambda item: (item["week_end"], item["location"]))
        if row["source_days"] < OPERATING_WEEK_DAYS
    ]
    if not short_rows:
        return [["OK", "All tracked store weeks include a full Tuesday-Sunday source set.", None, None]]
    return [
        [
            "Short Week",
            row["location"],
            row["week_end"],
            f"{row['source_days']} of {OPERATING_WEEK_DAYS} source days",
        ]
        for row in short_rows[:6]
    ]


def write_dashboard_table(
    ws,
    title: str,
    row: int,
    col: int,
    headers: list[str],
    rows: list[list[Any]],
) -> int:
    end_col = col + len(headers) - 1
    style_section_header(ws, row, col, end_col, title)
    for offset, header in enumerate(headers):
        cell = ws.cell(row=row + 1, column=col + offset, value=header)
        cell.fill = PatternFill("solid", fgColor="F3F4F6")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_offset, values in enumerate(rows, start=2):
        for col_offset, value in enumerate(values):
            cell = ws.cell(row=row + row_offset, column=col + col_offset, value=value)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    if rows:
        apply_dashboard_number_formats(ws, row + 2, row + len(rows) + 1)
    return row + len(rows) + 3


def write_dashboard_sheet(
    wb: Workbook,
    records: list[MetricRecord],
    weekly_location_rows: list[dict[str, Any]],
    ranked_rows: list[dict[str, Any]],
    star_rows: list[dict[str, Any]],
    action_rows: list[dict[str, Any]],
    store_trend_summary_rows: list[dict[str, Any]],
    group_trend_summary_rows: list[dict[str, Any]],
    config: dict[str, Any],
    source_dir: Path,
    public_start: date,
    public_end: date,
) -> None:
    ws = wb.create_sheet("Dashboard", 0)
    style_title(ws, "Red Onion Weekly Performance Dashboard", 13)
    ws.freeze_panes = "A4"
    for col, width in {
        "A": 15,
        "B": 18,
        "C": 20,
        "D": 12,
        "E": 30,
        "F": 34,
        "G": 16,
        "H": 15,
        "I": 18,
        "J": 18,
        "K": 18,
        "L": 34,
        "M": 30,
    }.items():
        ws.column_dimensions[col].width = width

    dates = sorted({record.report_date for record in records})
    latest_week_end = max((row["week_end"] for row in weekly_location_rows), default=None)
    latest_location_rows = [
        row for row in weekly_location_rows if latest_week_end and row["week_end"] == latest_week_end
    ]

    summary_rows = [
        ["Latest Week", public_end],
        ["Reports Read", len({record.source_file for record in records})],
        ["Date Coverage", format_date_range(min(dates), max(dates))],
        ["Weeks Tracked", len({row["week_end"] for row in weekly_location_rows})],
    ]
    snapshot_end = write_dashboard_table(ws, "Snapshot", 3, 1, ["Item", "Value"], summary_rows)

    group_rows = []
    for row in group_trend_summary_rows:
        _, status, focus = store_status(row)
        group_rows.append(
            [
                row.get("group", "All Stores"),
                status,
                compact_money(row["latest_gross_sales"]),
                compact_money(row["gross_sales_change"], signed=True),
                compact_count(row["guest_count_change"], signed=True),
                compact_money(row["check_average_change"], signed=True),
                compact_pct_points(row["wine_pct_change"]),
                compact_minutes(row["ticket_time_change_minutes"]),
                focus,
            ]
        )
    group_end = write_dashboard_table(
        ws,
        "All-Stores Group Pulse",
        3,
        4,
        [
            "Group",
            "Status",
            "Sales",
            "Sales Change",
            "Guest Change",
            "Check Change",
            "Wine Change",
            "Ticket Change",
            "Focus",
        ],
        group_rows,
    )

    write_dashboard_table(
        ws,
        "Data Quality",
        max(snapshot_end, group_end) + 1,
        1,
        ["Status", "Location / Note", "Week End", "Detail"],
        data_quality_warning_rows(weekly_location_rows),
    )

    coach_rows = [
        [
            row["priority"],
            row["location"],
            row["subject"],
            row["impact"],
            row["evidence"],
            row["recommended_follow_up"],
        ]
        for row in action_rows
        if row["action"] == "Coach"
    ][:5]
    recognize_rows = [
        [
            row["priority"],
            row["location"],
            row["subject"],
            row["impact"],
            row["evidence"],
            row["recommended_follow_up"],
        ]
        for row in action_rows
        if row["action"] == "Recognize"
    ][:5]
    if not coach_rows:
        coach_rows = [["", "", "No coach-now items", "", "", ""]]
    if not recognize_rows:
        recognize_rows = [["", "", "No recognition items", "", "", ""]]
    coach_end = write_dashboard_table(
        ws,
        "Coach First",
        max(snapshot_end, group_end) + 7,
        1,
        ["Priority", "Location", "Server", "Impact", "Evidence", "Recommended Follow-Up"],
        coach_rows,
    )
    recognize_end = write_dashboard_table(
        ws,
        "Recognize / Replicate",
        max(snapshot_end, group_end) + 7,
        8,
        ["Priority", "Location", "Server", "Impact", "Evidence", "Recommended Follow-Up"],
        recognize_rows,
    )

    store_rows = []
    for row in store_trend_summary_rows:
        _, status, focus = store_status(row)
        store_rows.append(
            [
                row["location"],
                status,
                compact_money(row["gross_sales_change"], signed=True),
                compact_count(row["guest_count_change"], signed=True),
                compact_money(row["check_average_change"], signed=True),
                compact_pct_points(row["wine_pct_change"]),
                compact_minutes(row["ticket_time_change_minutes"]),
                focus,
            ]
        )
    store_end = write_dashboard_table(
        ws,
        "Store Action Pulse",
        max(coach_end, recognize_end) + 1,
        1,
        [
            "Location",
            "Status",
            "Sales Change",
            "Guest Count",
            "Check Change",
            "Wine Change",
            "Ticket Change",
            "Recommended Focus",
        ],
        store_rows,
    )

    latest_ranked = [
        row
        for row in ranked_rows
        if latest_week_end
        and row["week_end"] == latest_week_end
        and not dashboard_excluded(row, config)
        and dashboard_trend_eligible(row, config)
    ]
    leaderboard_sections = [
        (
            "Highest Check Average",
            lambda rows: sorted(
                rows,
                key=lambda row: (row.get("check_average_rank") or 9999, row["display_name"]),
            )[:3],
            "Check Average",
            "check_average",
            "$#,##0.00",
        ),
        (
            "Highest Wine %",
            lambda rows: sorted(
                rows,
                key=lambda row: (row.get("wine_pct_rank") or 9999, row["display_name"]),
            )[:3],
            "Wine %",
            "wine_pct",
            "0.00%",
        ),
        (
            "Best Rate of Sale by Guest Count",
            lambda rows: sorted(
                rows,
                key=lambda row: (row.get("rate_rank") or 9999, row["display_name"]),
            )[:3],
            "Rate of Sale by Guest Count",
            "rate_of_sale_by_guest_count",
            "0.00",
        ),
        (
            "Fastest Ticket Time",
            lambda rows: sorted(
                rows,
                key=lambda row: (row.get("ticket_time_rank") or 9999, row["display_name"]),
            )[:3],
            "Ticket Time",
            "average_ticket_time_seconds",
            "[h]:mm:ss",
        ),
    ]

    table_row = store_end + 1
    for title, selector, value_label, value_field, number_format in leaderboard_sections:
        style_section_header(ws, table_row, 1, 6, title)
        for col, header in enumerate(["Location", "Rank", "Server", "Guest Count", value_label, "Active Days"], start=1):
            cell = ws.cell(row=table_row + 1, column=col, value=header)
            cell.fill = PatternFill("solid", fgColor="F3F4F6")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        write_row = table_row + 2
        for location in sorted({row["location"] for row in latest_ranked}):
            rows = [
                row
                for row in latest_ranked
                if row["location"] == location and row.get("guest_count", 0) > 0
            ]
            for visible_rank, selected in enumerate(selector(rows), start=1):
                if value_field == "average_ticket_time_seconds":
                    value = duration_fraction(selected[value_field])
                else:
                    value = selected[value_field]
                ws.cell(row=write_row, column=1, value=location)
                ws.cell(row=write_row, column=2, value=visible_rank)
                ws.cell(row=write_row, column=3, value=selected["display_name"])
                ws.cell(row=write_row, column=4, value=selected["guest_count"])
                ws.cell(row=write_row, column=5, value=value)
                ws.cell(row=write_row, column=6, value=selected["active_days"])
                ws.cell(row=write_row, column=5).number_format = number_format
                write_row += 1
        table_row = write_row + 2

    for row_index in range(4, ws.max_row + 1):
        for col_index in range(1, min(ws.max_column, 13) + 1):
            cell = ws.cell(row=row_index, column=col_index)
            fill = priority_fill(cell.value)
            if fill:
                cell.fill = fill
        if row_index <= store_end:
            ws.row_dimensions[row_index].height = 42 if row_index >= max(snapshot_end, group_end) + 9 else 32
        else:
            ws.row_dimensions[row_index].height = 18

    if latest_location_rows:
        chart_ws = wb.create_sheet("_Dashboard Chart Data")
        chart_ws.sheet_state = "hidden"
        chart_anchor_row = 10
        helper_row = chart_anchor_row
        helper_location_col = 1
        helper_value_col = 2
        chart_ws.cell(row=helper_row, column=helper_location_col, value="Location")
        chart_ws.cell(row=helper_row, column=helper_value_col, value="Check Average")
        for index, row in enumerate(sorted(latest_location_rows, key=lambda item: item["location"]), start=1):
            chart_ws.cell(row=helper_row + index, column=helper_location_col, value=row["location"])
            chart_ws.cell(row=helper_row + index, column=helper_value_col, value=row["check_average"])
            chart_ws.cell(row=helper_row + index, column=helper_value_col).number_format = "$#,##0.00"
        chart = BarChart()
        chart.type = "col"
        chart.title = "Latest Check Average By Location"
        chart.y_axis.title = "Check Average"
        chart.x_axis.title = "Location"
        data = Reference(chart_ws, min_col=helper_value_col, min_row=helper_row, max_row=helper_row + len(latest_location_rows))
        cats = Reference(chart_ws, min_col=helper_location_col, min_row=helper_row + 1, max_row=helper_row + len(latest_location_rows))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.legend = None
        chart.height = 7
        chart.width = 13
        ws.add_chart(chart, f"H{max(store_end + 2, table_row - 12)}")



def write_data_quality_sheet(
    wb: Workbook,
    records: list[MetricRecord],
    weekly_location_rows: list[dict[str, Any]],
    public_start: date,
    public_end: date,
) -> None:
    source_groups: dict[str, list[MetricRecord]] = defaultdict(list)
    for record in records:
        source_groups[record.source_file].append(record)

    source_headers = [
        "Source File",
        "Report Date",
        "Locations Found",
        "Server Rows",
        "Location Total Rows",
        "Status",
    ]
    source_rows: list[list[Any]] = []
    for source_file, source_records in sorted(source_groups.items()):
        locations = sorted({record.location for record in source_records})
        source_rows.append(
            [
                source_file,
                min(record.report_date for record in source_records),
                ", ".join(locations),
                sum(1 for record in source_records if not record.is_location_total),
                sum(1 for record in source_records if record.is_location_total),
                "OK" if len(locations) >= 2 else "Review",
            ]
        )

    ws = wb.create_sheet("Data Quality")
    style_title(ws, "Data Quality Checks", len(source_headers))
    ws.freeze_panes = "A4"
    for col_index, header in enumerate(source_headers, start=1):
        cell = ws.cell(row=3, column=col_index, value=header)
        cell.fill = PatternFill("solid", fgColor="E1E1E1")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_index, row in enumerate(source_rows, start=4):
        for col_index, value in enumerate(row, start=1):
            cell = ws.cell(row=row_index, column=col_index, value=value)
            if col_index == 2:
                cell.number_format = "m/d/yyyy"
    if source_rows:
        ref = f"A3:{get_column_letter(len(source_headers))}{len(source_rows) + 3}"
        table = Table(displayName="DataQualitySources", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)

    for col, width in {"A": 56, "B": 14, "C": 30, "D": 14, "E": 18, "F": 14}.items():
        ws.column_dimensions[col].width = width

    all_dates = sorted({record.report_date for record in records})
    check_row = len(source_rows) + 6
    style_section_header(ws, check_row, 1, 4, "Date Coverage")
    date_headers = ["Date", "Status", "Locations With Totals", "Source Files"]
    for col_index, header in enumerate(date_headers, start=1):
        cell = ws.cell(row=check_row + 1, column=col_index, value=header)
        cell.fill = PatternFill("solid", fgColor="F3F4F6")
        cell.font = Font(bold=True)
    expected_dates = []
    if all_dates:
        current = min(min(all_dates), public_start)
        coverage_end = max(max(all_dates), public_end)
        while current <= coverage_end:
            if is_operating_day(current):
                expected_dates.append(current)
            current += timedelta(days=1)

    location_totals_by_date: dict[date, set[str]] = defaultdict(set)
    source_files_by_date: dict[date, set[str]] = defaultdict(set)
    for record in records:
        if record.is_location_total:
            location_totals_by_date[record.report_date].add(record.location)
        source_files_by_date[record.report_date].add(record.source_file)

    for offset, report_date in enumerate(expected_dates, start=2):
        row_index = check_row + offset
        status = "OK" if report_date in all_dates else "Missing"
        ws.cell(row=row_index, column=1, value=report_date).number_format = "m/d/yyyy"
        ws.cell(row=row_index, column=2, value=status)
        ws.cell(
            row=row_index,
            column=3,
            value=", ".join(sorted(location_totals_by_date.get(report_date, set()))),
        )
        ws.cell(
            row=row_index,
            column=4,
            value=", ".join(sorted(source_files_by_date.get(report_date, set()))),
        )
        if status == "Missing":
            ws.cell(row=row_index, column=2).fill = PatternFill("solid", fgColor="F4CCCC")

    location_check_row = check_row + len(expected_dates) + 5
    style_section_header(ws, location_check_row, 1, 5, "Weekly Location Coverage")
    loc_headers = ["Week End", "Location", "Active Days", "Source Days", "Status"]
    for col_index, header in enumerate(loc_headers, start=1):
        cell = ws.cell(row=location_check_row + 1, column=col_index, value=header)
        cell.fill = PatternFill("solid", fgColor="F3F4F6")
        cell.font = Font(bold=True)
    for offset, row in enumerate(sorted(weekly_location_rows, key=lambda item: (item["week_end"], item["location"])), start=2):
        row_index = location_check_row + offset
        status = "OK" if row["source_days"] >= OPERATING_WEEK_DAYS else "Short Week"
        ws.cell(row=row_index, column=1, value=row["week_end"]).number_format = "m/d/yyyy"
        ws.cell(row=row_index, column=2, value=row["location"])
        ws.cell(row=row_index, column=3, value=row["active_days"])
        ws.cell(row=row_index, column=4, value=row["source_days"])
        ws.cell(row=row_index, column=5, value=status)
        if status != "OK":
            ws.cell(row=row_index, column=5).fill = PatternFill("solid", fgColor="FFF2CC")


def metric_week_trend_headers(entity_label: str) -> list[str]:
    return [
        "Week Start",
        "Week End",
        entity_label,
        "Gross Sales",
        "Gross Sales Change",
        "Guest Count",
        "Guest Count Change",
        "Check Average",
        "Check Average Change",
        "Wine Sales",
        "Wine Sales Change",
        "Wine %",
        "Wine % Change",
        "Rate of Sale by Guest Count",
        "Rate Change",
        "Average Ticket Time",
        "Ticket Time Change (Min)",
        "Active Days",
        "Source Days",
        "Trend Note",
    ]


def metric_week_trend_data(rows: list[dict[str, Any]], entity_field: str) -> list[list[Any]]:
    return [
        [
            row["week_start"],
            row["week_end"],
            row[entity_field],
            row["gross_sales"],
            row["gross_sales_change"],
            row["guest_count"],
            row["guest_count_change"],
            row["check_average"],
            row["check_average_change"],
            row["wine_sales"],
            row["wine_sales_change"],
            row["wine_pct"],
            row["wine_pct_change"],
            row["rate_of_sale_by_guest_count"],
            row["rate_change"],
            duration_fraction(row["average_ticket_time_seconds"]),
            row["ticket_time_change_minutes"],
            row["active_days"],
            row["source_days"],
            row["trend_note"],
        ]
        for row in rows
    ]


def metric_summary_headers(entity_label: str) -> list[str]:
    return [
        entity_label,
        "Weeks Tracked",
        "Active Days Tracked",
        "Source Days Tracked",
        "Total Gross Sales",
        "Total Guest Count",
        "Weighted Check Average",
        "Total Wine Sales",
        "Wine %",
        "Rate of Sale by Guest Count",
        "Average Ticket Time",
        "First Week End",
        "Latest Week End",
        "Prior Week End",
        "Latest Gross Sales",
        "Prior Gross Sales",
        "Gross Sales Change",
        "Latest Guest Count",
        "Prior Guest Count",
        "Guest Count Change",
        "Latest Check Average",
        "Prior Check Average",
        "Check Average Change",
        "Latest Wine %",
        "Prior Wine %",
        "Wine % Change",
        "Latest Rate of Sale by Guest Count",
        "Prior Rate of Sale by Guest Count",
        "Rate Change",
        "Latest Ticket Time",
        "Prior Ticket Time",
        "Ticket Time Change (Min)",
    ]


def metric_summary_data(rows: list[dict[str, Any]], entity_field: str) -> list[list[Any]]:
    return [
        [
            row[entity_field],
            row["weeks_tracked"],
            row["active_days_tracked"],
            row["source_days_tracked"],
            row["total_gross_sales"],
            row["total_guest_count"],
            row["weighted_check_average"],
            row["total_wine_sales"],
            row["overall_wine_pct"],
            row["weighted_rate"],
            duration_fraction(row["weighted_ticket_time_seconds"]),
            row["first_week_end"],
            row["latest_week_end"],
            row["prior_week_end"],
            row["latest_gross_sales"],
            row["prior_gross_sales"],
            row["gross_sales_change"],
            row["latest_guest_count"],
            row["prior_guest_count"],
            row["guest_count_change"],
            row["latest_check_average"],
            row["prior_check_average"],
            row["check_average_change"],
            row["latest_wine_pct"],
            row["prior_wine_pct"],
            row["wine_pct_change"],
            row["latest_rate"],
            row["prior_rate"],
            row["rate_change"],
            duration_fraction(row["latest_ticket_time_seconds"]),
            duration_fraction(row["prior_ticket_time_seconds"])
            if row["prior_ticket_time_seconds"] is not None
            else None,
            row["ticket_time_change_minutes"],
        ]
        for row in rows
    ]


def write_action_board_sheet(wb: Workbook, action_rows: list[dict[str, Any]]) -> None:
    headers = [
        "Priority",
        "Action",
        "Location",
        "Person / Area",
        "Signal",
        "Score / Impact",
        "Evidence",
        "Recommended Follow-Up",
        "Week End",
        "Guest Count",
        "Active Days",
    ]
    data = [
        [
            row["priority"],
            row["action"],
            row["location"],
            row["subject"],
            row["signal"],
            row["impact"],
            row["evidence"],
            row["recommended_follow_up"],
            row["week_end"],
            row["guest_count"],
            row["active_days"],
        ]
        for row in action_rows
    ]
    ws = write_table_sheet(
        wb,
        "Action Board",
        headers,
        data,
        "ActionBoard",
        widths={
            "A": 13,
            "B": 16,
            "C": 20,
            "D": 24,
            "E": 18,
            "F": 32,
            "G": 44,
            "H": 56,
            "I": 14,
            "J": 12,
            "K": 12,
        },
        index=1,
    )
    ws.sheet_properties.tabColor = "7A1E1E"
    for row in range(4, len(data) + 4):
        priority = ws.cell(row=row, column=1).value
        action = ws.cell(row=row, column=2).value
        signal = ws.cell(row=row, column=5).value
        fill = priority_fill(priority) or priority_fill(action) or priority_fill(signal)
        if fill:
            for col in range(1, len(headers) + 1):
                ws.cell(row=row, column=col).fill = fill
        for col in (6, 7, 8):
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 42


def _write_master_workbook_base(
    records: list[MetricRecord],
    output_path: Path,
    config: dict[str, Any],
    source_dir: Path,
    public_start: date,
    public_end: date,
) -> Path:
    weekly_server_rows, weekly_location_rows = weekly_rollups(records)
    rank_min_guest_count = int(
        config.get("master_min_guest_count_for_rankings", config.get("public_min_guest_count", 1))
    )
    ranked_rows = weekly_server_rank_rows(weekly_server_rows, rank_min_guest_count)
    server_trend_detail_rows = server_week_trend_rows(ranked_rows)
    server_star_detail_rows = server_star_rows(server_trend_detail_rows, config)
    trend_rows = trend_summary_rows(weekly_server_rows)
    store_week_trend_detail_rows = weekly_metric_trend_rows(weekly_location_rows, ("location",))
    store_trend_summary = weekly_metric_summary_rows(weekly_location_rows, ("location",))
    weekly_group_rows = group_weekly_rows(records)
    group_week_trend_detail_rows = weekly_metric_trend_rows(weekly_group_rows, ("group",))
    group_trend_summary = weekly_metric_summary_rows(weekly_group_rows, ("group",))
    action_board_rows = build_action_board_rows(
        server_star_detail_rows,
        store_trend_summary,
        group_trend_summary,
        weekly_location_rows,
    )

    wb = Workbook()
    wb.remove(wb.active)
    write_dashboard_sheet(
        wb,
        records,
        weekly_location_rows,
        ranked_rows,
        server_star_detail_rows,
        action_board_rows,
        store_trend_summary,
        group_trend_summary,
        config,
        source_dir,
        public_start,
        public_end,
    )
    write_action_board_sheet(wb, action_board_rows)

    star_headers = [
        "Category",
        "Suggested Action",
        "Priority",
        "Composite Score",
        "Week Start",
        "Week End",
        "Location",
        "Raw Server",
        "Display Name",
        "Guest Count",
        "Active Days",
        "Check Average",
        "Check Average Change",
        "Wine %",
        "Wine % Change",
        "Rate of Sale by Guest Count",
        "Rate Change",
        "Average Ticket Time",
        "Ticket Time Change (Min)",
        "Average Rank Movement",
        "Why",
    ]
    star_data = [
        [
            row["category"],
            star_action(row)[1],
            star_action(row)[0],
            row["composite_score"],
            row["week_start"],
            row["week_end"],
            row["location"],
            row["raw_user_name"],
            row["display_name"],
            row["guest_count"],
            row["active_days"],
            row["check_average"],
            row["check_average_change"],
            row["wine_pct"],
            row["wine_pct_change"],
            row["rate_of_sale_by_guest_count"],
            row["rate_change"],
            duration_fraction(row["average_ticket_time_seconds"]),
            row["ticket_time_change_minutes"],
            row["average_rank_movement"],
            row["why"],
        ]
        for row in server_star_detail_rows
    ]
    star_ws = write_table_sheet(
        wb,
        "Rising & Falling Stars",
        star_headers,
        star_data,
        "RisingFallingStars",
        widths={
            "A": 16,
            "B": 16,
            "C": 14,
            "D": 16,
            "E": 14,
            "F": 14,
            "G": 22,
            "H": 28,
            "I": 28,
            "P": 24,
            "U": 54,
        },
    )
    for row_index in range(4, len(star_data) + 4):
        fill = (
            priority_fill(star_ws.cell(row=row_index, column=1).value)
            or priority_fill(star_ws.cell(row=row_index, column=2).value)
            or priority_fill(star_ws.cell(row=row_index, column=3).value)
        )
        if fill:
            for col_index in range(1, len(star_headers) + 1):
                star_ws.cell(row=row_index, column=col_index).fill = fill
        star_ws.cell(row=row_index, column=len(star_headers)).alignment = Alignment(
            wrap_text=True, vertical="top"
        )
        star_ws.row_dimensions[row_index].height = 32

    server_headers = [
        "Week Start",
        "Week End",
        "Location",
        "Raw Server",
        "Display Name",
        "Gross Sales",
        "Guest Count",
        "Check Average",
        "Wine Sales",
        "Wine %",
        "Rate of Sale by Guest Count",
        "Average Ticket Time",
        "Active Days",
        "Source Days",
    ]
    server_data = [
        [
            row["week_start"],
            row["week_end"],
            row["location"],
            row["raw_user_name"],
            row["display_name"],
            row["gross_sales"],
            row["guest_count"],
            row["check_average"],
            row["wine_sales"],
            row["wine_pct"],
            row["rate_of_sale_by_guest_count"],
            duration_fraction(row["average_ticket_time_seconds"]),
            row["active_days"],
            row["source_days"],
        ]
        for row in sorted(
            weekly_server_rows,
            key=lambda item: (item["week_end"], item["location"], item["display_name"]),
        )
    ]
    write_table_sheet(wb, "Weekly Server Metrics", server_headers, server_data, "WeeklyServerMetrics")

    ranking_headers = [
        "Week Start",
        "Week End",
        "Location",
        "Raw Server",
        "Display Name",
        "Guest Count",
        "Check Average",
        "Check Average Rank",
        "Wine %",
        "Wine % Rank",
        "Rate of Sale by Guest Count",
        "Rate Rank",
        "Average Ticket Time",
        "Ticket Time Rank",
        "Active Days",
    ]
    ranking_data = [
        [
            row["week_start"],
            row["week_end"],
            row["location"],
            row["raw_user_name"],
            row["display_name"],
            row["guest_count"],
            row["check_average"],
            row["check_average_rank"],
            row["wine_pct"],
            row["wine_pct_rank"],
            row["rate_of_sale_by_guest_count"],
            row["rate_rank"],
            duration_fraction(row["average_ticket_time_seconds"]),
            row["ticket_time_rank"],
            row["active_days"],
        ]
        for row in ranked_rows
    ]
    write_table_sheet(
        wb,
        "Weekly Server Rankings",
        ranking_headers,
        ranking_data,
        "WeeklyServerRankings",
        widths={"A": 14, "B": 14, "C": 22, "D": 28, "E": 28, "K": 22, "M": 18},
    )

    trend_detail_headers = [
        "Week Start",
        "Week End",
        "Location",
        "Raw Server",
        "Display Name",
        "Guest Count",
        "Check Average",
        "Check Average Rank",
        "Prior Check Average Rank",
        "Check Average Rank Movement",
        "Check Average Change",
        "Wine %",
        "Wine % Rank",
        "Prior Wine % Rank",
        "Wine % Rank Movement",
        "Wine % Change",
        "Rate of Sale by Guest Count",
        "Rate Rank",
        "Prior Rate Rank",
        "Rate Rank Movement",
        "Rate Change",
        "Average Ticket Time",
        "Ticket Time Rank",
        "Prior Ticket Time Rank",
        "Ticket Time Rank Movement",
        "Ticket Time Change (Min)",
        "Active Days",
        "Trend Note",
    ]
    trend_detail_data = [
        [
            row["week_start"],
            row["week_end"],
            row["location"],
            row["raw_user_name"],
            row["display_name"],
            row["guest_count"],
            row["check_average"],
            row["check_average_rank"],
            row["prior_check_average_rank"],
            row["check_average_rank_movement"],
            row["check_average_change"],
            row["wine_pct"],
            row["wine_pct_rank"],
            row["prior_wine_pct_rank"],
            row["wine_pct_rank_movement"],
            row["wine_pct_change"],
            row["rate_of_sale_by_guest_count"],
            row["rate_rank"],
            row["prior_rate_rank"],
            row["rate_rank_movement"],
            row["rate_change"],
            duration_fraction(row["average_ticket_time_seconds"]),
            row["ticket_time_rank"],
            row["prior_ticket_time_rank"],
            row["ticket_time_rank_movement"],
            row["ticket_time_change_minutes"],
            row["active_days"],
            row["trend_note"],
        ]
        for row in server_trend_detail_rows
    ]
    write_table_sheet(
        wb,
        "Server Week Trends",
        trend_detail_headers,
        trend_detail_data,
        "ServerWeekTrends",
        widths={
            "A": 14,
            "B": 14,
            "C": 22,
            "D": 28,
            "E": 28,
            "I": 22,
            "J": 24,
            "M": 18,
            "N": 20,
            "O": 22,
            "R": 20,
            "S": 20,
            "T": 22,
            "W": 22,
            "X": 24,
            "Y": 22,
        },
    )

    location_headers = [
        "Week Start",
        "Week End",
        "Location",
        "Gross Sales",
        "Guest Count",
        "Check Average",
        "Wine Sales",
        "Wine %",
        "Rate of Sale by Guest Count",
        "Average Ticket Time",
        "Active Days",
        "Source Days",
    ]
    location_data = [
        [
            row["week_start"],
            row["week_end"],
            row["location"],
            row["gross_sales"],
            row["guest_count"],
            row["check_average"],
            row["wine_sales"],
            row["wine_pct"],
            row["rate_of_sale_by_guest_count"],
            duration_fraction(row["average_ticket_time_seconds"]),
            row["active_days"],
            row["source_days"],
        ]
        for row in sorted(weekly_location_rows, key=lambda item: (item["week_end"], item["location"]))
    ]
    write_table_sheet(wb, "Weekly Location Metrics", location_headers, location_data, "WeeklyLocationMetrics")

    trend_tab_widths = {
        "A": 14,
        "B": 14,
        "C": 24,
        "E": 20,
        "G": 20,
        "I": 22,
        "K": 20,
        "M": 16,
        "N": 24,
        "Q": 22,
        "T": 16,
    }
    write_table_sheet(
        wb,
        "Store Week Trends",
        metric_week_trend_headers("Location"),
        metric_week_trend_data(store_week_trend_detail_rows, "location"),
        "StoreWeekTrends",
        widths=trend_tab_widths,
    )
    write_table_sheet(
        wb,
        "Store Trend Summary",
        metric_summary_headers("Location"),
        metric_summary_data(store_trend_summary, "location"),
        "StoreTrendSummary",
        widths={
            "A": 24,
            "J": 24,
            "K": 18,
            "Q": 20,
            "W": 22,
            "AA": 24,
            "AC": 16,
            "AF": 22,
        },
    )
    write_table_sheet(
        wb,
        "Group Week Trends",
        metric_week_trend_headers("Group"),
        metric_week_trend_data(group_week_trend_detail_rows, "group"),
        "GroupWeekTrends",
        widths=trend_tab_widths,
    )
    write_table_sheet(
        wb,
        "Group Trend Summary",
        metric_summary_headers("Group"),
        metric_summary_data(group_trend_summary, "group"),
        "GroupTrendSummary",
        widths={
            "A": 18,
            "J": 24,
            "K": 18,
            "Q": 20,
            "W": 22,
            "AA": 24,
            "AC": 16,
            "AF": 22,
        },
    )

    daily_server_headers = [
        "Report Date",
        "Location",
        "Raw Server",
        "Display Name",
        "Source File",
        "Gross Sales",
        "Guest Count",
        "Check Average",
        "Wine Sales",
        "Wine %",
        "Rate of Sale by Guest Count",
        "Average Ticket Time",
    ]
    daily_server_data = [
        [
            record.report_date,
            record.location,
            record.raw_user_name,
            record.display_name,
            record.source_file,
            record.gross_sales,
            record.guest_count,
            record.check_average,
            record.wine_sales,
            record.wine_pct,
            record.rate_of_sale_by_guest_count,
            duration_fraction(record.average_ticket_time_seconds),
        ]
        for record in sorted(
            (item for item in records if not item.is_location_total),
            key=lambda item: (item.report_date, item.location, item.display_name),
        )
    ]
    write_table_sheet(wb, "Daily Server Detail", daily_server_headers, daily_server_data, "DailyServerDetail")

    daily_location_data = [
        [
            record.report_date,
            record.location,
            record.source_file,
            record.gross_sales,
            record.guest_count,
            record.check_average,
            record.wine_sales,
            record.wine_pct,
            record.rate_of_sale_by_guest_count,
            duration_fraction(record.average_ticket_time_seconds),
        ]
        for record in sorted(
            (item for item in records if item.is_location_total),
            key=lambda item: (item.report_date, item.location),
        )
    ]
    write_table_sheet(
        wb,
        "Daily Location Detail",
        [
            "Report Date",
            "Location",
            "Source File",
            "Gross Sales",
            "Guest Count",
            "Check Average",
            "Wine Sales",
            "Wine %",
            "Rate of Sale by Guest Count",
            "Average Ticket Time",
        ],
        daily_location_data,
        "DailyLocationDetail",
    )

    trend_headers = [
        "Location",
        "Raw Server",
        "Display Name",
        "Weeks Tracked",
        "Active Days Tracked",
        "Total Gross Sales",
        "Total Guest Count",
        "Weighted Check Average",
        "Total Wine Sales",
        "Wine %",
        "Rate of Sale by Guest Count",
        "Average Ticket Time",
        "First Week End",
        "Latest Week End",
        "Prior Week End",
        "Latest Check Average",
        "Prior Check Average",
        "Check Average Change",
        "Latest Wine %",
        "Prior Wine %",
        "Wine % Change",
        "Latest Rate of Sale by Guest Count",
        "Prior Rate of Sale by Guest Count",
        "Rate Change",
        "Latest Ticket Time",
        "Prior Ticket Time",
        "Ticket Time Change (Min)",
    ]
    trend_data = [
        [
            row["location"],
            row["raw_user_name"],
            row["display_name"],
            row["weeks_tracked"],
            row["active_days_tracked"],
            row["total_gross_sales"],
            row["total_guest_count"],
            row["weighted_check_average"],
            row["total_wine_sales"],
            row["overall_wine_pct"],
            row["weighted_rate"],
            duration_fraction(row["weighted_ticket_time_seconds"]),
            row["first_week_end"],
            row["latest_week_end"],
            row["prior_week_end"],
            row["latest_check_average"],
            row["prior_check_average"],
            row["check_average_change"],
            row["latest_wine_pct"],
            row["prior_wine_pct"],
            row["wine_pct_change"],
            row["latest_rate"],
            row["prior_rate"],
            row["rate_change"],
            duration_fraction(row["latest_ticket_time_seconds"]),
            duration_fraction(row["prior_ticket_time_seconds"])
            if row["prior_ticket_time_seconds"] is not None
            else None,
            row["ticket_time_change_minutes"],
        ]
        for row in trend_rows
    ]
    write_table_sheet(
        wb,
        "Server Trend Summary",
        trend_headers,
        trend_data,
        "ServerTrendSummary",
        widths={
            "A": 22,
            "B": 28,
            "C": 28,
            "D": 14,
            "E": 18,
            "K": 22,
            "L": 18,
            "R": 20,
            "U": 16,
            "V": 24,
            "X": 14,
            "AA": 22,
        },
    )

    write_data_quality_sheet(wb, records, weekly_location_rows, public_start, public_end)

    notes = wb.create_sheet("Run Notes", 2)
    notes.sheet_view.showGridLines = False
    notes.column_dimensions["A"].width = 28
    notes.column_dimensions["B"].width = 90
    notes["A1"] = "Red Onion Server Master"
    notes["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    notes["A1"].fill = PatternFill("solid", fgColor="7A1E1E")
    notes.merge_cells("A1:B1")
    note_rows = [
        ("Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Source Folder", str(source_dir)),
        ("Operating Week", f"{OPERATING_WEEK_LABEL}; Mondays are closed and excluded from weekly rollups."),
        (
            "Report Date Source",
            "Uses each raw workbook's Date(s) value; filename date fallback is one day after the business date.",
        ),
        ("Raw Reports Read", len({record.source_file for record in records})),
        ("Date Coverage", format_date_range(min(r.report_date for r in records), max(r.report_date for r in records))),
        ("Public Snapshot Dates", format_date_range(public_start, public_end)),
        ("Public Exclude Patterns", ", ".join(config.get("public_exclude_name_contains", []))),
        ("Dashboard Exclude Patterns", ", ".join(config.get("dashboard_exclude_name_contains", []))),
        ("Ranking Minimum Guest Count", rank_min_guest_count),
        (
            "Dashboard Trend Eligibility",
            "Server trend lists include rows with guest count >= "
            f"{config.get('dashboard_min_guest_count_for_trends', 25)} or active days >= "
            f"{config.get('dashboard_min_active_days_for_trends', 3)}.",
        ),
        (
            "Dashboard",
            "Dashboard highlights coach-first items, recognition opportunities, store action pulses, group pulse, and current metric leaders.",
        ),
        (
            "Action Board",
            "Prioritized follow-up list for coaching, recognition, store review, group review, and data-quality checks.",
        ),
        (
            "Rising/Falling Stars",
            "Composite score balances check average, wine percent, rate of sale, ticket time, and rank movement.",
        ),
        (
            "Trend Tabs",
            "Server, store, and group trend tabs show week-over-week changes plus latest/prior/total-period summaries.",
        ),
        ("Metric Rule", "Check average and wine percent are recalculated from rolled-up sales, guests, and wine sales."),
        ("Metric Rule", "Rate of sale and ticket time are guest-weighted averages."),
    ]
    for row_index, (label, value) in enumerate(note_rows, start=3):
        notes.cell(row=row_index, column=1, value=label).font = Font(bold=True)
        notes.cell(row=row_index, column=2, value=value)
        notes.cell(row=row_index, column=2).alignment = Alignment(wrap_text=True)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and value.strip():
        try:
            return datetime.fromisoformat(value.strip()).date()
        except ValueError:
            return None
    return None


def safe_pct_delta(current: float | None, comparison: float | None) -> float | None:
    if current is None or comparison in (None, 0):
        return None
    return (current - comparison) / comparison


def management_threshold(config: dict[str, Any], field: str) -> dict[str, Any]:
    defaults = DEFAULT_CONFIG["management_score_thresholds"][field]
    configured = config.get("management_score_thresholds", {}).get(field, {})
    return {**defaults, **configured}


def directional_score(change: float | None, threshold: dict[str, Any]) -> int:
    if change is None:
        return 0
    directional = -change if threshold.get("lower_is_better") else change
    neutral = float(threshold["neutral"])
    strong = float(threshold["strong"])
    if directional >= strong:
        return 2
    if directional >= neutral:
        return 1
    if directional <= -strong:
        return -2
    if directional <= -neutral:
        return -1
    return 0


def directional_variance(current: float, benchmark: float, threshold: dict[str, Any]) -> float:
    variance = current - benchmark
    return -variance if threshold.get("lower_is_better") else variance


def aggregate_weekly_rows(rows: Iterable[dict[str, Any]]) -> dict[str, float] | None:
    selected = list(rows)
    if not selected:
        return None
    gross_sales = sum(float(row.get("gross_sales", 0) or 0) for row in selected)
    guest_count = sum(float(row.get("guest_count", 0) or 0) for row in selected)
    wine_sales = sum(float(row.get("wine_sales", 0) or 0) for row in selected)
    rate_weight = sum(float(row.get("guest_count", 0) or 0) for row in selected)
    ticket_weight = rate_weight
    rate_weighted = sum(
        float(row.get("rate_of_sale_by_guest_count", 0) or 0)
        * float(row.get("guest_count", 0) or 0)
        for row in selected
    )
    ticket_weighted = sum(
        float(row.get("average_ticket_time_seconds", 0) or 0)
        * float(row.get("guest_count", 0) or 0)
        for row in selected
    )
    return {
        "gross_sales": gross_sales,
        "guest_count": guest_count,
        "check_average": gross_sales / guest_count if guest_count else 0.0,
        "wine_sales": wine_sales,
        "wine_pct": wine_sales / gross_sales if gross_sales else 0.0,
        "rate_of_sale_by_guest_count": rate_weighted / rate_weight if rate_weight else 0.0,
        "average_ticket_time_seconds": (
            ticket_weighted / ticket_weight if ticket_weight else 0.0
        ),
    }


def full_week_ends_by_location(
    weekly_location_rows: list[dict[str, Any]],
) -> tuple[dict[str, set[date]], set[date]]:
    locations = sorted({row["location"] for row in weekly_location_rows})
    by_location = {
        location: {
            row["week_end"]
            for row in weekly_location_rows
            if row["location"] == location and row.get("source_days", 0) >= OPERATING_WEEK_DAYS
        }
        for location in locations
    }
    if not locations:
        return by_location, set()
    global_full = set.intersection(*(by_location[location] for location in locations))
    return by_location, global_full


def metric_driver(field: str, change: float) -> str:
    if field == "check_average":
        return f"Check avg {format_change(change, 'currency')}"
    if field == "wine_pct":
        return f"Wine {format_change(change, 'pct_points')}"
    if field == "rate_of_sale_by_guest_count":
        direction = "improved" if change < 0 else "worsened"
        return f"Rate {direction} {abs(change):.3f}"
    minutes = abs(change) / 60
    direction = "faster" if change < 0 else "slower"
    return f"Ticket {minutes:.1f} min {direction}"


def recommended_server_follow_up(row: dict[str, Any]) -> str:
    negative_fields = [field for field, score in row["metric_scores"].items() if score < 0]
    positive_fields = [field for field, score in row["metric_scores"].items() if score > 0]
    field = negative_fields[0] if negative_fields else positive_fields[0] if positive_fields else None
    if row["action"] == "Recognize & Replicate":
        if negative_fields:
            return "Recognize the improvement, then address the watch item before sharing the successful practice."
        return "Recognize the improvement and ask what changed so the practice can be shared."
    if row["action"] == "Reinforce Improvement":
        return "Acknowledge the progress and reinforce the strongest improving behavior."
    if row["action"] == "Protect Performance":
        return "Check in on the decline while reinforcing the server's current strengths."
    if row["action"] == "Coach Fundamentals":
        return "Review recent shifts and agree on one measurable improvement for next week."
    if row["action"] != "Coach Now":
        return "Monitor another full week before taking action."
    if field == "check_average":
        return "Review suggestive selling, add-ons, and check-building opportunities on recent shifts."
    if field == "wine_pct":
        return "Coach wine pairing prompts and wine attachment opportunities on recent shifts."
    if field == "rate_of_sale_by_guest_count":
        return "Review table pacing and section load to identify the source of the slower rate."
    if field == "average_ticket_time_seconds":
        return "Review service flow and ticket-time bottlenecks on the affected shifts."
    return "Review recent shifts and coach the largest declining driver."


def management_server_rows(
    weekly_server_rows: list[dict[str, Any]],
    weekly_location_rows: list[dict[str, Any]],
    ranked_rows: list[dict[str, Any]],
    targets: dict[str, dict[str, float | None]],
    config: dict[str, Any],
) -> list[dict[str, Any]]:
    if not weekly_server_rows:
        return []
    latest_week_end = max(row["week_end"] for row in weekly_server_rows)
    baseline_limit = int(config.get("dashboard_baseline_full_weeks", 4))
    min_prior_weeks = int(config.get("dashboard_min_prior_full_weeks", 2))
    min_prior_guests = float(config.get("dashboard_min_prior_guest_count", 50))
    full_by_location, _ = full_week_ends_by_location(weekly_location_rows)
    rank_lookup = {
        (row["week_end"], row["location"], row["raw_user_name"]): row for row in ranked_rows
    }
    location_rows: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in weekly_location_rows:
        location_rows[row["location"]].append(row)

    output: list[dict[str, Any]] = []
    for current in weekly_server_rows:
        if current["week_end"] != latest_week_end or dashboard_excluded(current, config):
            continue
        location = current["location"]
        eligible_week_ends = sorted(
            week_end
            for week_end in full_by_location.get(location, set())
            if week_end < latest_week_end
        )[-baseline_limit:]
        prior_rows = [
            row
            for row in weekly_server_rows
            if row["location"] == location
            and row["raw_user_name"] == current["raw_user_name"]
            and row["week_end"] in eligible_week_ends
        ]
        baseline = aggregate_weekly_rows(prior_rows)
        location_baseline = aggregate_weekly_rows(
            row
            for row in location_rows.get(location, [])
            if row["week_end"] in eligible_week_ends
        )
        prior_guest_count = sum(float(row.get("guest_count", 0) or 0) for row in prior_rows)
        full_latest = latest_week_end in full_by_location.get(location, set())
        prominent = (
            full_latest
            and dashboard_trend_eligible(current, config)
            and len(prior_rows) >= min_prior_weeks
            and prior_guest_count >= min_prior_guests
        )

        metric_scores: dict[str, int] = {}
        changes: dict[str, float | None] = {}
        positive_drivers: list[str] = []
        negative_drivers: list[str] = []
        for field in (
            "check_average",
            "wine_pct",
            "rate_of_sale_by_guest_count",
            "average_ticket_time_seconds",
        ):
            change = current[field] - baseline[field] if baseline else None
            score = directional_score(change, management_threshold(config, field))
            changes[field] = change
            metric_scores[field] = score
            if change is not None and score > 0:
                positive_drivers.append(metric_driver(field, change))
            elif change is not None and score < 0:
                negative_drivers.append(metric_driver(field, change))

        latest_rank = rank_lookup.get((latest_week_end, location, current["raw_user_name"]))
        rank_movements: list[float] = []
        for rank_field in (
            "check_average_rank",
            "wine_pct_rank",
            "rate_rank",
            "ticket_time_rank",
        ):
            prior_ranks = [
                rank_lookup[(row["week_end"], location, current["raw_user_name"])].get(rank_field)
                for row in prior_rows
                if (row["week_end"], location, current["raw_user_name"]) in rank_lookup
                and rank_lookup[(row["week_end"], location, current["raw_user_name"])].get(rank_field)
                is not None
            ]
            latest_value = latest_rank.get(rank_field) if latest_rank else None
            if prior_ranks and latest_value is not None:
                rank_movements.append(sum(prior_ranks) / len(prior_ranks) - latest_value)
        average_rank_move = sum(rank_movements) / len(rank_movements) if rank_movements else None
        rank_score = 1 if average_rank_move is not None and average_rank_move >= 3 else -1 if average_rank_move is not None and average_rank_move <= -3 else 0
        core_score = sum(metric_scores.values())
        composite_score = core_score + rank_score
        positive_count = sum(score > 0 for score in metric_scores.values())
        negative_count = sum(score < 0 for score in metric_scores.values())
        if composite_score >= 3 and positive_count >= 2:
            momentum = "Rising"
        elif composite_score <= -3 and negative_count >= 2:
            momentum = "Falling"
        else:
            momentum = "Stable"

        benchmark_values: dict[str, float | None] = {}
        benchmark_sources: dict[str, str] = {}
        level_statuses: dict[str, str] = {}
        for field in (
            "check_average",
            "wine_pct",
            "rate_of_sale_by_guest_count",
            "average_ticket_time_seconds",
        ):
            target_value = targets.get(location, {}).get(field)
            benchmark_value = target_value if target_value is not None else (
                location_baseline.get(field) if location_baseline else None
            )
            benchmark_values[field] = benchmark_value
            benchmark_sources[field] = (
                "Target" if target_value is not None else f"{len(eligible_week_ends)}-week baseline"
            )
            if benchmark_value is None:
                level_statuses[field] = "Unavailable"
                continue
            threshold = management_threshold(config, field)
            variance = directional_variance(current[field], benchmark_value, threshold)
            neutral = float(threshold["neutral"])
            level_statuses[field] = (
                "Above" if variance >= neutral else "Below" if variance <= -neutral else "On Track"
            )
        below_count = sum(status == "Below" for status in level_statuses.values())
        above_count = sum(status == "Above" for status in level_statuses.values())
        available_levels = sum(status != "Unavailable" for status in level_statuses.values())
        if not available_levels:
            performance_level = "Insufficient History"
        elif below_count >= 2:
            performance_level = "Below Benchmark"
        elif above_count >= 2 and below_count <= 1:
            performance_level = "Above Benchmark"
        else:
            performance_level = "On Track"

        if not prominent:
            priority, action, confidence = "Monitor", "Monitor", "Low Sample"
        elif momentum == "Falling" and performance_level == "Below Benchmark":
            priority, action, confidence = "High", "Coach Now", "High"
        elif momentum == "Falling":
            priority, action, confidence = "Medium", "Protect Performance", "High"
        elif momentum == "Rising" and performance_level == "Below Benchmark":
            priority, action, confidence = "Medium", "Reinforce Improvement", "High"
        elif momentum == "Rising":
            priority, action, confidence = "Recognize", "Recognize & Replicate", "High"
        elif performance_level == "Below Benchmark":
            priority, action, confidence = "Medium", "Coach Fundamentals", "High"
        else:
            priority, action, confidence = "Monitor", "Monitor", "High"

        why_parts = []
        if positive_drivers:
            why_parts.append("Improving: " + "; ".join(positive_drivers[:3]))
        if negative_drivers:
            why_parts.append("Watch: " + "; ".join(negative_drivers[:3]))
        if average_rank_move is not None and rank_score:
            why_parts.append(f"Rank movement {average_rank_move:+.1f}")
        row = {
            **current,
            "prior_weeks": len(prior_rows),
            "prior_guest_count": prior_guest_count,
            "baseline": baseline,
            "changes": changes,
            "metric_scores": metric_scores,
            "rank_modifier": rank_score,
            "average_rank_movement": average_rank_move,
            "composite_score": composite_score,
            "momentum": momentum,
            "performance_level": performance_level,
            "benchmark_values": benchmark_values,
            "benchmark_sources": benchmark_sources,
            "level_statuses": level_statuses,
            "positive_drivers": positive_drivers,
            "negative_drivers": negative_drivers,
            "priority": priority,
            "action": action,
            "confidence": confidence,
            "prominent": prominent,
            "why": " | ".join(why_parts) if why_parts else "No material movement",
        }
        row["recommended_next_step"] = recommended_server_follow_up(row)
        output.append(row)

    priority_order = {"High": 0, "Medium": 1, "Recognize": 2, "Monitor": 3}
    output.sort(
        key=lambda row: (
            priority_order.get(row["priority"], 9),
            row["composite_score"] if row["priority"] in {"High", "Medium"} else -row["composite_score"],
            row["location"],
            row["display_name"],
        )
    )
    return output


def management_entity_rows(
    weekly_rows: list[dict[str, Any]],
    entity_field: str,
    targets: dict[str, dict[str, float | None]],
    config: dict[str, Any],
    allowed_baseline_weeks: dict[str, set[date]] | set[date],
) -> list[dict[str, Any]]:
    if not weekly_rows:
        return []
    baseline_limit = int(config.get("dashboard_baseline_full_weeks", 4))
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in weekly_rows:
        grouped[row[entity_field]].append(row)
    output: list[dict[str, Any]] = []
    for entity, rows in grouped.items():
        rows.sort(key=lambda row: row["week_end"])
        latest = rows[-1]
        prior = rows[-2] if len(rows) > 1 else None
        allowed = allowed_baseline_weeks.get(entity, set()) if isinstance(allowed_baseline_weeks, dict) else allowed_baseline_weeks
        prior_full_ends = sorted(
            week_end for week_end in allowed if week_end < latest["week_end"]
        )[-baseline_limit:]
        baseline_rows = [row for row in rows if row["week_end"] in prior_full_ends]
        baseline = aggregate_weekly_rows(baseline_rows)
        if baseline and baseline_rows:
            baseline["gross_sales"] /= len(baseline_rows)
            baseline["guest_count"] /= len(baseline_rows)
            baseline["wine_sales"] /= len(baseline_rows)
        benchmark_values: dict[str, float | None] = {}
        benchmark_sources: dict[str, str] = {}
        for field, _, _ in MANAGEMENT_METRICS:
            target_value = targets.get(entity, {}).get(field)
            benchmark_values[field] = target_value if target_value is not None else (
                baseline.get(field) if baseline else None
            )
            benchmark_sources[field] = (
                "Target" if target_value is not None else f"{len(baseline_rows)}-week baseline"
            )

        prior_changes = {
            field: latest[field] - prior[field] if prior else None
            for field, _, _ in MANAGEMENT_METRICS
        }
        benchmark_changes = {
            field: latest[field] - benchmark_values[field]
            if benchmark_values[field] is not None
            else None
            for field, _, _ in MANAGEMENT_METRICS
        }
        materiality = config.get("management_materiality", DEFAULT_CONFIG["management_materiality"])
        sales_pct = safe_pct_delta(latest["gross_sales"], benchmark_values["gross_sales"])
        guest_pct = safe_pct_delta(latest["guest_count"], benchmark_values["guest_count"])
        check_change = benchmark_changes["check_average"]
        wine_change = benchmark_changes["wine_pct"]
        rate_change = benchmark_changes["rate_of_sale_by_guest_count"]
        ticket_change_minutes = (
            benchmark_changes["average_ticket_time_seconds"] / 60
            if benchmark_changes["average_ticket_time_seconds"] is not None
            else None
        )
        negative_count = sum(
            [
                sales_pct is not None and sales_pct <= -float(materiality["sales_pct"]),
                guest_pct is not None and guest_pct <= -float(materiality["guest_pct"]),
                check_change is not None and check_change <= -float(materiality["check_average"]),
                wine_change is not None and wine_change <= -float(materiality["wine_pct"]),
                rate_change is not None and rate_change >= float(materiality["rate"]),
                ticket_change_minutes is not None
                and ticket_change_minutes >= float(materiality["ticket_minutes"]),
            ]
        )
        if sales_pct is not None and guest_pct is not None and sales_pct <= -float(materiality["sales_pct"]) and guest_pct <= -float(materiality["guest_pct"]):
            priority = "High"
            status = "Traffic Watch"
            focus = "Review traffic drivers, staffing, holidays, and the event calendar."
        elif check_change is not None and wine_change is not None and check_change <= -float(materiality["check_average"]) and wine_change <= -float(materiality["wine_pct"]):
            priority = "Medium"
            status = "Upsell Watch"
            focus = "Coach check-building and wine attachment opportunities."
        elif ticket_change_minutes is not None and ticket_change_minutes >= float(materiality["ticket_minutes"]) and negative_count >= 2:
            priority = "Medium"
            status = "Service Watch"
            focus = "Review service flow, table pacing, and section load."
        elif negative_count >= 2:
            priority = "Medium"
            status = "Mixed Watch"
            focus = "Review the material declines and agree on one management response."
        else:
            priority = "Monitor"
            status = "On Track / Mixed"
            focus = "No material alert; continue monitoring."
        output.append(
            {
                "entity": entity,
                "latest": latest,
                "prior": prior,
                "baseline": baseline,
                "baseline_weeks": len(baseline_rows),
                "benchmark_values": benchmark_values,
                "benchmark_sources": benchmark_sources,
                "prior_changes": prior_changes,
                "benchmark_changes": benchmark_changes,
                "priority": priority,
                "status": status,
                "recommended_focus": focus,
            }
        )
    output.sort(key=lambda row: (0 if row["entity"] == "All Stores" else 1, row["entity"]))
    return output


def find_sheet_header_row(ws, required_header: str) -> int | None:
    for row in range(1, min(ws.max_row, 12) + 1):
        if any(ws.cell(row=row, column=col).value == required_header for col in range(1, ws.max_column + 1)):
            return row
    return None


def records_from_sheet(ws, required_header: str) -> list[dict[str, Any]]:
    header_row = find_sheet_header_row(ws, required_header)
    if header_row is None:
        return []
    headers = [ws.cell(row=header_row, column=col).value for col in range(1, ws.max_column + 1)]
    records: list[dict[str, Any]] = []
    for row in range(header_row + 1, ws.max_row + 1):
        record = {
            str(header): ws.cell(row=row, column=col).value
            for col, header in enumerate(headers, start=1)
            if header
        }
        if record.get(required_header):
            records.append(record)
    return records


def read_management_state(output_path: Path) -> dict[str, Any]:
    state: dict[str, Any] = {
        "targets": {},
        "owners": [],
        "active_actions": [],
        "action_history": [],
    }
    if not output_path.exists():
        return state
    try:
        wb = load_workbook(output_path, data_only=False)
    except Exception as exc:
        raise RuntimeError(
            f"Could not read the existing master workbook at {output_path}. "
            "Close the workbook in Excel and rerun; no source files were moved."
        ) from exc
    try:
        if "Management Setup" in wb.sheetnames:
            ws = wb["Management Setup"]
            target_header_row = find_sheet_header_row(ws, "Entity")
            if target_header_row:
                headers = {
                    ws.cell(row=target_header_row, column=col).value: col
                    for col in range(1, ws.max_column + 1)
                    if ws.cell(row=target_header_row, column=col).value
                }
                row = target_header_row + 1
                while row <= ws.max_row:
                    entity = ws.cell(row=row, column=headers["Entity"]).value
                    if not entity:
                        break
                    values: dict[str, float | None] = {}
                    for field, label in TARGET_FIELDS:
                        value = ws.cell(row=row, column=headers[label]).value if label in headers else None
                        if isinstance(value, (int, float)):
                            values[field] = float(value) * 60 if field == "average_ticket_time_seconds" else float(value)
                        else:
                            values[field] = None
                    state["targets"][str(entity)] = values
                    row += 1
            state["owners"] = [
                str(ws.cell(row=row, column=10).value).strip()
                for row in range(6, min(ws.max_row, 25) + 1)
                if ws.cell(row=row, column=10).value
            ]
        if "Action Board" in wb.sheetnames:
            state["active_actions"] = records_from_sheet(wb["Action Board"], "Action ID")
        if "Action History" in wb.sheetnames:
            state["action_history"] = records_from_sheet(wb["Action History"], "Action ID")
    finally:
        wb.close()
    return state


def action_episode_id(entity_key: str, first_seen: date) -> str:
    digest = hashlib.sha1(f"{entity_key}|{first_seen.isoformat()}".encode("utf-8")).hexdigest()
    return digest[:12].upper()


def compact_server_evidence(row: dict[str, Any]) -> str:
    parts = []
    if row["positive_drivers"]:
        parts.append("Improving: " + "; ".join(row["positive_drivers"][:2]))
    if row["negative_drivers"]:
        parts.append("Watch: " + "; ".join(row["negative_drivers"][:2]))
    parts.append(f"{row['guest_count']:,.0f} guests / {row['active_days']} days")
    return " | ".join(parts)


def build_management_action_signals(
    server_rows: list[dict[str, Any]],
    store_rows: list[dict[str, Any]],
    group_rows: list[dict[str, Any]],
    weekly_location_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    signals: list[dict[str, Any]] = []
    for row in server_rows:
        if not row["prominent"] or row["action"] == "Monitor":
            continue
        family = "recognition" if row["action"] == "Recognize & Replicate" else "coaching"
        entity_key = "|".join(
            ["server", row["location"], str(row["raw_user_name"]), family]
        ).casefold()
        signals.append(
            {
                "Entity Key": entity_key,
                "Priority": row["priority"],
                "Location": row["location"],
                "Person / Area": row["display_name"],
                "Action": row["action"],
                "Signal": f"{row['momentum']} / {row['performance_level']}",
                "Why It Matters": compact_server_evidence(row),
                "Recommended Next Step": row["recommended_next_step"],
                "Performance Level": row["performance_level"],
                "Momentum": row["momentum"],
                "Confidence": row["confidence"],
                "Last Seen": row["week_end"],
            }
        )
    for action_type, rows in (("store", store_rows), ("group", group_rows)):
        for row in rows:
            if row["priority"] == "Monitor":
                continue
            entity = row["entity"]
            latest = row["latest"]
            sales_pct = safe_pct_delta(latest["gross_sales"], row["benchmark_values"]["gross_sales"])
            guest_pct = safe_pct_delta(latest["guest_count"], row["benchmark_values"]["guest_count"])
            entity_key = f"{action_type}|{entity}|{row['status']}".casefold()
            signals.append(
                {
                    "Entity Key": entity_key,
                    "Priority": row["priority"],
                    "Location": entity,
                    "Person / Area": entity,
                    "Action": "Store Review" if action_type == "store" else "Group Review",
                    "Signal": row["status"],
                    "Why It Matters": (
                        f"Sales {sales_pct:+.1%}; guests {guest_pct:+.1%} vs benchmark"
                        if sales_pct is not None and guest_pct is not None
                        else "Material movement requires management review"
                    ),
                    "Recommended Next Step": row["recommended_focus"],
                    "Performance Level": row["status"],
                    "Momentum": "Watch",
                    "Confidence": "High" if row["baseline_weeks"] >= 2 else "Low Sample",
                    "Last Seen": latest["week_end"],
                }
            )
    if weekly_location_rows:
        latest_week_end = max(row["week_end"] for row in weekly_location_rows)
        for row in weekly_location_rows:
            if row["week_end"] != latest_week_end or row.get("source_days", 0) >= OPERATING_WEEK_DAYS:
                continue
            entity_key = f"data-quality|{row['location']}|short-week".casefold()
            signals.append(
                {
                    "Entity Key": entity_key,
                    "Priority": "Review",
                    "Location": row["location"],
                    "Person / Area": row["location"],
                    "Action": "Data Quality",
                    "Signal": "Incomplete Latest Week",
                    "Why It Matters": f"{row['source_days']} of {OPERATING_WEEK_DAYS} source days",
                    "Recommended Next Step": "Confirm the missing reports before using trends for coaching.",
                    "Performance Level": "Preliminary",
                    "Momentum": "Not Scored",
                    "Confidence": "Low Sample",
                    "Last Seen": latest_week_end,
                }
            )
    return signals


def merge_management_actions(
    signals: list[dict[str, Any]], state: dict[str, Any]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    prior_active = {
        str(row.get("Entity Key", "")).casefold(): row
        for row in state.get("active_actions", [])
        if row.get("Entity Key")
    }
    history = list(state.get("action_history", []))
    current: list[dict[str, Any]] = []
    matched_keys: set[str] = set()
    completed_statuses = {"complete", "dismissed"}
    for signal in signals:
        key = str(signal["Entity Key"]).casefold()
        matched_keys.add(key)
        prior = prior_active.get(key)
        prior_status = str(prior.get("Status", "")).casefold() if prior else ""
        last_seen = as_date(signal.get("Last Seen")) or date.today()
        if prior and prior_status not in completed_statuses:
            first_seen = as_date(prior.get("First Seen")) or last_seen
            action_id = str(prior.get("Action ID") or action_episode_id(key, first_seen))
            status = prior.get("Status") or "Open"
            owner = prior.get("Owner") or ""
            due_date = as_date(prior.get("Due Date"))
            notes = prior.get("Manager Notes") or ""
        else:
            if prior:
                completed = dict(prior)
                completed["Signal State"] = completed.get("Signal State") or "Completed"
                history.append(completed)
            first_seen = last_seen
            action_id = action_episode_id(key, first_seen)
            status, owner, due_date, notes = "Open", "", None, ""
        weeks_open = max(1, ((last_seen - first_seen).days // 7) + 1)
        current.append(
            {
                "Action ID": action_id,
                **signal,
                "Status": status,
                "Owner": owner,
                "Due Date": due_date,
                "First Seen": first_seen,
                "Weeks Open": weeks_open,
                "Manager Notes": notes,
                "Signal State": "Current",
            }
        )
    for key, prior in prior_active.items():
        if key in matched_keys:
            continue
        cleared = dict(prior)
        cleared["Signal State"] = "Cleared"
        history.append(cleared)

    history_by_id: dict[str, dict[str, Any]] = {}
    current_ids = {str(row["Action ID"]) for row in current}
    for row in history:
        action_id = str(row.get("Action ID") or "")
        if action_id and action_id not in current_ids:
            history_by_id[action_id] = row
    priority_order = {"High": 0, "Medium": 1, "Recognize": 2, "Share": 3, "Review": 4, "Monitor": 5}
    current.sort(
        key=lambda row: (
            priority_order.get(str(row.get("Priority")), 9),
            str(row.get("Location", "")),
            str(row.get("Person / Area", "")),
        )
    )
    history_rows = sorted(
        history_by_id.values(),
        key=lambda row: (
            as_date(row.get("Last Seen")) or date.min,
            str(row.get("Location", "")),
        ),
        reverse=True,
    )
    return current, history_rows


def remove_sheet_if_present(wb: Workbook, name: str) -> None:
    if name in wb.sheetnames:
        wb.remove(wb[name])


def add_management_navigation(ws) -> None:
    links = [
        ("Dashboard", "Dashboard"),
        ("Actions", "Action Board"),
        ("Servers", "Server Scorecard"),
        ("Stores", "Store & Group Scorecards"),
        ("Stars", "Rising & Falling Stars"),
        ("Quality", "Data Quality"),
        ("Setup", "Management Setup"),
    ]
    for col, (label, target) in enumerate(links, start=1):
        cell = ws.cell(row=2, column=col, value=label)
        cell.hyperlink = f"#'{target}'!A1"
        cell.font = Font(color="7A1E1E", bold=True, underline="single", size=9)
        cell.alignment = Alignment(horizontal="center")


def style_management_title(ws, title: str, end_col: int) -> None:
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    cell = ws.cell(row=1, column=1, value=title)
    cell.fill = PatternFill("solid", fgColor="7A1E1E")
    cell.font = Font(color="FFFFFF", bold=True, size=16)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 26


def write_management_setup_sheet(
    wb: Workbook,
    targets: dict[str, dict[str, float | None]],
    owners: list[str],
    config: dict[str, Any],
) -> None:
    remove_sheet_if_present(wb, "Management Setup")
    ws = wb.create_sheet("Management Setup")
    style_management_title(ws, "Management Setup", 10)
    ws.merge_cells("A3:G3")
    ws["A3"] = "Blue cells are management inputs. Changes are preserved and applied on the next weekly run. Blank targets use the rolling four-full-week baseline."
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["A3"].fill = PatternFill("solid", fgColor="D9EAF7")
    ws.row_dimensions[3].height = 34

    headers = ["Entity", *[label for _, label in TARGET_FIELDS]]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=col, value=header)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    entities = ["All Stores", *config.get("locations", {}).keys()]
    for row_index, entity in enumerate(entities, start=6):
        ws.cell(row=row_index, column=1, value=entity).font = Font(bold=True)
        for col, (field, _) in enumerate(TARGET_FIELDS, start=2):
            value = targets.get(entity, {}).get(field)
            if field == "average_ticket_time_seconds" and value is not None:
                value = value / 60
            cell = ws.cell(row=row_index, column=col, value=value)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="right")
    for row in range(6, 6 + len(entities)):
        ws.cell(row=row, column=2).number_format = "$#,##0"
        ws.cell(row=row, column=3).number_format = "#,##0"
        ws.cell(row=row, column=4).number_format = "$0.00"
        ws.cell(row=row, column=5).number_format = "0.0%"
        ws.cell(row=row, column=6).number_format = "0.000"
        ws.cell(row=row, column=7).number_format = "0.0"
    table = Table(displayName="ManagementTargets", ref=f"A5:G{5 + len(entities)}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(table)

    ws["J5"] = "Owner List"
    ws["J5"].fill = PatternFill("solid", fgColor="D9E1F2")
    ws["J5"].font = Font(bold=True)
    for row in range(6, 26):
        ws.cell(row=row, column=10, value=owners[row - 6] if row - 6 < len(owners) else None)
        ws.cell(row=row, column=10).fill = PatternFill("solid", fgColor="D9EAF7")

    threshold_headers = ["Metric", "Neutral Band", "Strong Band", "Better Direction", "Use"]
    for col, header in enumerate(threshold_headers, start=1):
        cell = ws.cell(row=12, column=col, value=header)
        cell.fill = PatternFill("solid", fgColor="E7E6E6")
        cell.font = Font(bold=True)
    metric_labels = {
        "check_average": "Check Average",
        "wine_pct": "Wine %",
        "rate_of_sale_by_guest_count": "Rate of Sale",
        "average_ticket_time_seconds": "Ticket Time",
    }
    for row, field in enumerate(metric_labels, start=13):
        threshold = management_threshold(config, field)
        neutral = float(threshold["neutral"])
        strong = float(threshold["strong"])
        if field == "average_ticket_time_seconds":
            neutral, strong = neutral / 60, strong / 60
        ws.cell(row=row, column=1, value=metric_labels[field])
        ws.cell(row=row, column=2, value=neutral)
        ws.cell(row=row, column=3, value=strong)
        ws.cell(row=row, column=4, value="Lower" if threshold.get("lower_is_better") else "Higher")
        ws.cell(row=row, column=5, value="Momentum scoring and benchmark status")
        if field == "wine_pct":
            ws.cell(row=row, column=2).number_format = "0.0%"
            ws.cell(row=row, column=3).number_format = "0.0%"
        elif field == "check_average":
            ws.cell(row=row, column=2).number_format = "$0.00"
            ws.cell(row=row, column=3).number_format = "$0.00"
        else:
            ws.cell(row=row, column=2).number_format = "0.000" if field != "average_ticket_time_seconds" else "0.0"
            ws.cell(row=row, column=3).number_format = "0.000" if field != "average_ticket_time_seconds" else "0.0"

    widths = {"A": 22, "B": 18, "C": 18, "D": 22, "E": 18, "F": 16, "G": 22, "J": 24}
    for column, width in widths.items():
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A5"


def action_row_values(row: dict[str, Any]) -> list[Any]:
    return [row.get(header) for header in ACTION_HEADERS]


def write_action_tracking_sheet(
    wb: Workbook,
    name: str,
    rows: list[dict[str, Any]],
    *,
    editable: bool,
) -> None:
    remove_sheet_if_present(wb, name)
    ws = wb.create_sheet(name)
    style_management_title(ws, name, len(ACTION_HEADERS))
    header_row = 4
    for col, header in enumerate(ACTION_HEADERS, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_index, row in enumerate(rows, start=header_row + 1):
        for col, value in enumerate(action_row_values(row), start=1):
            cell = ws.cell(row=row_index, column=col, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=col in {10, 11, 13})
        priority_style = priority_fill(row.get("Priority"))
        if priority_style:
            ws.cell(row=row_index, column=3).fill = priority_style
        ws.cell(row=row_index, column=6).number_format = "m/d/yyyy"
        ws.cell(row=row_index, column=12).number_format = "m/d/yyyy"
        ws.cell(row=row_index, column=16).number_format = "m/d/yyyy"
        ws.row_dimensions[row_index].height = 36
    if rows:
        table = Table(
            displayName="ActionBoardTable" if editable else "ActionHistoryTable",
            ref=f"A{header_row}:T{header_row + len(rows)}",
        )
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleLight1", showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False
        )
        ws.add_table(table)
    ws.column_dimensions["A"].hidden = True
    ws.column_dimensions["B"].hidden = True
    ws.column_dimensions["O"].hidden = True
    ws.column_dimensions["P"].hidden = True
    ws.column_dimensions["S"].hidden = True
    ws.column_dimensions["T"].hidden = True
    widths = {
        "C": 12, "D": 14, "E": 18, "F": 13, "G": 20, "H": 24, "I": 22,
        "J": 22, "K": 46, "L": 48, "M": 14, "N": 34, "O": 20, "P": 14,
        "Q": 13, "R": 12, "S": 14,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "G5"
    ws.sheet_view.zoomScale = 80
    if rows:
        first, last = header_row + 1, header_row + len(rows)
        red_fill = PatternFill("solid", fgColor="F4CCCC")
        amber_fill = PatternFill("solid", fgColor="FFF2CC")
        green_fill = PatternFill("solid", fgColor="D9EAD3")
        blue_fill = PatternFill("solid", fgColor="D9EAF7")
        ws.conditional_formatting.add(f"C{first}:C{last}", FormulaRule(formula=[f'$C{first}="High"'], fill=red_fill))
        ws.conditional_formatting.add(f"C{first}:C{last}", FormulaRule(formula=[f'$C{first}="Medium"'], fill=amber_fill))
        ws.conditional_formatting.add(f"C{first}:C{last}", FormulaRule(formula=[f'$C{first}="Recognize"'], fill=green_fill))
        ws.conditional_formatting.add(f"D{first}:D{last}", FormulaRule(formula=[f'$D{first}="Complete"'], fill=green_fill))
        ws.conditional_formatting.add(f"D{first}:D{last}", FormulaRule(formula=[f'$D{first}="Blocked"'], fill=red_fill))
        ws.conditional_formatting.add(
            f"F{first}:F{last}",
            FormulaRule(
                formula=[f'AND($F{first}<TODAY(),$F{first}<>"",$D{first}<>"Complete",$D{first}<>"Dismissed")'],
                fill=red_fill,
            ),
        )
        if editable:
            for row in range(first, last + 1):
                for col in (4, 5, 6, 14):
                    ws.cell(row=row, column=col).fill = blue_fill
            status_validation = DataValidation(
                type="list", formula1='"Open,In Progress,Blocked,Complete,Dismissed"', allow_blank=False
            )
            owner_validation = DataValidation(
                type="list", formula1="'Management Setup'!$J$6:$J$25", allow_blank=True
            )
            ws.add_data_validation(status_validation)
            ws.add_data_validation(owner_validation)
            status_validation.add(f"D{first}:D{last}")
            owner_validation.add(f"E{first}:E{last}")


def write_server_scorecard_sheet(wb: Workbook, rows: list[dict[str, Any]]) -> None:
    remove_sheet_if_present(wb, "Server Scorecard")
    headers = [
        "Priority", "Location", "Server", "Guest Count", "Active Days", "Confidence",
        "Performance Level", "Momentum", "Action", "Score", "Check Average", "Check vs Benchmark",
        "Wine %", "Wine vs Benchmark", "Rate of Sale", "Rate vs Benchmark", "Ticket Time",
        "Ticket vs Benchmark (Min)", "Positive Drivers", "Watch Drivers", "Recommended Next Step",
    ]
    data = []
    for row in rows:
        benchmark = row["benchmark_values"]
        data.append(
            [
                row["priority"], row["location"], row["display_name"], row["guest_count"],
                row["active_days"], row["confidence"], row["performance_level"], row["momentum"],
                row["action"], row["composite_score"], row["check_average"],
                row["check_average"] - benchmark["check_average"] if benchmark["check_average"] is not None else None,
                row["wine_pct"], row["wine_pct"] - benchmark["wine_pct"] if benchmark["wine_pct"] is not None else None,
                row["rate_of_sale_by_guest_count"],
                row["rate_of_sale_by_guest_count"] - benchmark["rate_of_sale_by_guest_count"] if benchmark["rate_of_sale_by_guest_count"] is not None else None,
                duration_fraction(row["average_ticket_time_seconds"]),
                (row["average_ticket_time_seconds"] - benchmark["average_ticket_time_seconds"]) / 60 if benchmark["average_ticket_time_seconds"] is not None else None,
                "; ".join(row["positive_drivers"]), "; ".join(row["negative_drivers"]), row["recommended_next_step"],
            ]
        )
    ws = write_table_sheet(
        wb, "Server Scorecard", headers, data, "ServerScorecard",
        widths={"A": 12, "B": 20, "C": 24, "F": 14, "G": 20, "H": 12, "I": 22,
                "S": 40, "T": 40, "U": 48}
    )
    ws.freeze_panes = "D4"
    ws.sheet_view.zoomScale = 75
    for row in range(4, 4 + len(data)):
        for col in (1, 6, 7, 8, 9):
            fill = priority_fill(ws.cell(row=row, column=col).value)
            if fill:
                ws.cell(row=row, column=col).fill = fill
        for col in (19, 20, 21):
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32


def write_rising_falling_sheet(wb: Workbook, rows: list[dict[str, Any]]) -> None:
    remove_sheet_if_present(wb, "Rising & Falling Stars")
    star_rows = [row for row in rows if row["prominent"] and row["momentum"] in {"Rising", "Falling"}]
    headers = [
        "Category", "Action", "Priority", "Location", "Server", "Performance Level", "Score",
        "Guest Count", "Active Days", "Check vs Baseline", "Wine vs Baseline", "Rate vs Baseline",
        "Ticket vs Baseline (Min)", "Positive Drivers", "Watch Drivers", "Recommended Next Step",
    ]
    data = [
        [
            f"{row['momentum']} Star", row["action"], row["priority"], row["location"], row["display_name"],
            row["performance_level"], row["composite_score"], row["guest_count"], row["active_days"],
            row["changes"]["check_average"], row["changes"]["wine_pct"], row["changes"]["rate_of_sale_by_guest_count"],
            row["changes"]["average_ticket_time_seconds"] / 60 if row["changes"]["average_ticket_time_seconds"] is not None else None,
            "; ".join(row["positive_drivers"]), "; ".join(row["negative_drivers"]), row["recommended_next_step"],
        ]
        for row in star_rows
    ]
    ws = write_table_sheet(
        wb, "Rising & Falling Stars", headers, data, "RisingFallingStarsV2",
        widths={"A": 15, "B": 24, "C": 12, "D": 20, "E": 24, "F": 20,
                "N": 40, "O": 40, "P": 48}
    )
    ws.freeze_panes = "F4"
    ws.sheet_view.zoomScale = 80
    for row in range(4, 4 + len(data)):
        for col in (1, 2, 3, 6):
            fill = priority_fill(ws.cell(row=row, column=col).value)
            if fill:
                ws.cell(row=row, column=col).fill = fill
        for col in (14, 15, 16):
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 34


def format_management_value(field: str, value: float | None) -> tuple[Any, str]:
    if value is None:
        return None, "General"
    if field == "gross_sales":
        return value, "$#,##0"
    if field == "guest_count":
        return value, "#,##0"
    if field == "check_average":
        return value, "$0.00"
    if field == "wine_pct":
        return value, "0.0%"
    if field == "rate_of_sale_by_guest_count":
        return value, "0.000"
    return duration_fraction(value), "[h]:mm"


def management_metric_status(field: str, item: dict[str, Any], config: dict[str, Any] | None = None) -> str:
    config = config or DEFAULT_CONFIG
    current = item["latest"][field]
    benchmark = item["benchmark_values"].get(field)
    if benchmark is None:
        return "No Baseline"
    if field in {"gross_sales", "guest_count"}:
        delta = safe_pct_delta(current, benchmark)
        threshold_key = "sales_pct" if field == "gross_sales" else "guest_pct"
        threshold = float(config.get("management_materiality", DEFAULT_CONFIG["management_materiality"])[threshold_key])
        return "Above" if delta is not None and delta >= threshold else "Watch" if delta is not None and delta <= -threshold else "On Track"
    threshold = management_threshold(config, field)
    variance = directional_variance(current, benchmark, threshold)
    neutral = float(threshold["neutral"])
    return "Above" if variance >= neutral else "Watch" if variance <= -neutral else "On Track"


def write_store_group_scorecards_sheet(
    wb: Workbook, rows: list[dict[str, Any]], config: dict[str, Any]
) -> None:
    remove_sheet_if_present(wb, "Store & Group Scorecards")
    ws = wb.create_sheet("Store & Group Scorecards")
    style_management_title(ws, "Store & Group Scorecards", 7)
    ws.freeze_panes = "A4"
    current_row = 4
    for item in rows:
        entity = item["entity"]
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        title = ws.cell(row=current_row, column=1, value=f"{entity} | {item['status']} | {item['recommended_focus']}")
        title.fill = priority_fill(item["priority"]) or PatternFill("solid", fgColor="E7E6E6")
        title.font = Font(bold=True)
        title.alignment = Alignment(wrap_text=True)
        for col, header in enumerate(["Metric", "Current", "vs Prior", "vs Benchmark", "Benchmark", "Source", "Status"], start=1):
            cell = ws.cell(row=current_row + 1, column=col, value=header)
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        for offset, (field, label, _) in enumerate(MANAGEMENT_METRICS, start=2):
            row = current_row + offset
            current_value, number_format = format_management_value(field, item["latest"][field])
            prior_change = item["prior_changes"][field]
            benchmark_change = item["benchmark_changes"][field]
            benchmark_value, benchmark_format = format_management_value(field, item["benchmark_values"][field])
            ws.cell(row=row, column=1, value=label)
            ws.cell(row=row, column=2, value=current_value).number_format = number_format
            if field in {"gross_sales", "guest_count"}:
                prior_value = safe_pct_delta(item["latest"][field], item["prior"][field] if item["prior"] else None)
                benchmark_delta = safe_pct_delta(item["latest"][field], item["benchmark_values"][field])
                ws.cell(row=row, column=3, value=prior_value).number_format = "0.0%"
                ws.cell(row=row, column=4, value=benchmark_delta).number_format = "0.0%"
            elif field == "average_ticket_time_seconds":
                ws.cell(row=row, column=3, value=prior_change / 60 if prior_change is not None else None).number_format = "0.0"
                ws.cell(row=row, column=4, value=benchmark_change / 60 if benchmark_change is not None else None).number_format = "0.0"
            else:
                ws.cell(row=row, column=3, value=prior_change).number_format = number_format
                ws.cell(row=row, column=4, value=benchmark_change).number_format = number_format
            ws.cell(row=row, column=5, value=benchmark_value).number_format = benchmark_format
            ws.cell(row=row, column=6, value=item["benchmark_sources"][field])
            status = management_metric_status(field, item, config)
            ws.cell(row=row, column=7, value=status)
            fill = priority_fill(
                "Medium" if status == "Watch" else "Recognize" if status == "Above" else "Monitor"
            )
            if fill:
                ws.cell(row=row, column=7).fill = fill
        current_row += 10
    for column, width in {"A": 24, "B": 16, "C": 16, "D": 18, "E": 16, "F": 20, "G": 14}.items():
        ws.column_dimensions[column].width = width


def write_management_data_quality_sheet(
    wb: Workbook,
    weekly_location_rows: list[dict[str, Any]],
) -> None:
    remove_sheet_if_present(wb, "Data Quality")
    ws = wb.create_sheet("Data Quality")
    style_management_title(ws, "Data Quality", 6)
    latest_week_end = max((row["week_end"] for row in weekly_location_rows), default=None)
    latest_rows = [row for row in weekly_location_rows if row["week_end"] == latest_week_end]
    latest_complete = bool(latest_rows) and all(
        row.get("source_days", 0) >= OPERATING_WEEK_DAYS for row in latest_rows
    )
    ws.merge_cells("A3:F3")
    ws["A3"] = (
        f"Latest week ending {latest_week_end:%m/%d/%Y} is complete and suitable for management trends."
        if latest_complete and latest_week_end
        else "Latest week is incomplete. Management trends are preliminary and server actions are suppressed."
    )
    ws["A3"].fill = PatternFill("solid", fgColor="D9EAD3" if latest_complete else "F4CCCC")
    ws["A3"].font = Font(bold=True)
    ws["A3"].alignment = Alignment(wrap_text=True)
    for col, header in enumerate(["Latest Week", "Location", "Active Days", "Source Days", "Status", "Management Use"], start=1):
        cell = ws.cell(row=5, column=col, value=header)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.font = Font(bold=True)
    for row_index, row in enumerate(sorted(latest_rows, key=lambda item: item["location"]), start=6):
        status = "Complete" if row.get("source_days", 0) >= OPERATING_WEEK_DAYS else "Short Week"
        values = [
            row["week_end"], row["location"], row["active_days"], row["source_days"], status,
            "Use" if status == "Complete" else "Preliminary only",
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_index, column=col, value=value)
        ws.cell(row=row_index, column=1).number_format = "m/d/yyyy"
        fill = priority_fill("Recognize" if status == "Complete" else "High")
        if fill:
            ws.cell(row=row_index, column=5).fill = fill

    ws["A10"] = "Historical Exceptions"
    ws["A10"].font = Font(bold=True)
    ws["A10"].fill = PatternFill("solid", fgColor="E7E6E6")
    historical = [
        row for row in weekly_location_rows if row.get("source_days", 0) < OPERATING_WEEK_DAYS
    ]
    for col, header in enumerate(["Week End", "Location", "Active Days", "Source Days", "Status", "Benchmark Treatment"], start=1):
        cell = ws.cell(row=11, column=col, value=header)
        cell.fill = PatternFill("solid", fgColor="F3F4F6")
        cell.font = Font(bold=True)
    for row_index, row in enumerate(sorted(historical, key=lambda item: (item["week_end"], item["location"])), start=12):
        values = [row["week_end"], row["location"], row["active_days"], row["source_days"], "Short Week", "Excluded"]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_index, column=col, value=value)
        ws.cell(row=row_index, column=1).number_format = "m/d/yyyy"
    for column, width in {"A": 16, "B": 22, "C": 14, "D": 14, "E": 16, "F": 24}.items():
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A5"


def write_management_chart_data(
    wb: Workbook,
    weekly_location_rows: list[dict[str, Any]],
    weekly_group_rows: list[dict[str, Any]],
) -> tuple[int, int]:
    remove_sheet_if_present(wb, "_Dashboard Chart Data")
    ws = wb.create_sheet("_Dashboard Chart Data")
    locations = sorted({row["location"] for row in weekly_location_rows})
    week_ends = sorted({row["week_end"] for row in weekly_location_rows})[-5:]
    ws.cell(row=1, column=1, value="Week End")
    for col, location in enumerate(locations, start=2):
        ws.cell(row=1, column=col, value=location.replace("RC ", ""))
    location_lookup = {
        (row["week_end"], row["location"]): row for row in weekly_location_rows
    }
    for row_index, week_end in enumerate(week_ends, start=2):
        ws.cell(row=row_index, column=1, value=week_end.strftime("%m/%d"))
        for col, location in enumerate(locations, start=2):
            row = location_lookup.get((week_end, location))
            ws.cell(row=row_index, column=col, value=row["gross_sales"] if row else None)

    group_rows = sorted(weekly_group_rows, key=lambda row: row["week_end"])[-5:]
    ws.cell(row=1, column=6, value="Week End")
    ws.cell(row=1, column=7, value="Guest Count")
    for row_index, row in enumerate(group_rows, start=2):
        ws.cell(row=row_index, column=6, value=row["week_end"].strftime("%m/%d"))
        ws.cell(row=row_index, column=7, value=row["guest_count"])
    ws.sheet_state = "hidden"
    return len(week_ends), len(group_rows)


def signed_management_delta(field: str, value: float | None) -> str:
    if value is None:
        return "n/a"
    if field == "gross_sales":
        return compact_money(value, signed=True)
    if field == "guest_count":
        return compact_count(value, signed=True)
    if field == "check_average":
        return compact_money(value, signed=True)
    if field == "wine_pct":
        return compact_pct_points(value)
    if field == "rate_of_sale_by_guest_count":
        return f"{value:+.3f}"
    return compact_minutes(value / 60)


def write_management_dashboard_sheet(
    wb: Workbook,
    records: list[MetricRecord],
    weekly_location_rows: list[dict[str, Any]],
    weekly_group_rows: list[dict[str, Any]],
    server_rows: list[dict[str, Any]],
    store_rows: list[dict[str, Any]],
    group_rows: list[dict[str, Any]],
    action_rows: list[dict[str, Any]],
) -> None:
    remove_sheet_if_present(wb, "Dashboard")
    ws = wb.create_sheet("Dashboard")
    style_management_title(ws, "Red Onion Management Dashboard", 12)
    for column in range(1, 13):
        ws.column_dimensions[get_column_letter(column)].width = 15
    ws.freeze_panes = "A5"
    latest_week_end = max((row["week_end"] for row in weekly_location_rows), default=None)
    latest_location_rows = [row for row in weekly_location_rows if row["week_end"] == latest_week_end]
    latest_complete = bool(latest_location_rows) and all(
        row.get("source_days", 0) >= OPERATING_WEEK_DAYS for row in latest_location_rows
    )
    ws.merge_cells("A3:L3")
    ws["A3"] = (
        f"Week ending {latest_week_end:%m/%d/%Y} | {len({record.source_file for record in records})} reports | "
        f"{len({row['week_end'] for row in weekly_location_rows})} weeks tracked"
        if latest_week_end
        else "No management data available"
    )
    ws["A3"].font = Font(bold=True, color="595959")
    ws.merge_cells("A4:L4")
    ws["A4"] = (
        "LATEST WEEK COMPLETE - trends and actions are management-ready"
        if latest_complete
        else "PRELIMINARY - latest week is incomplete; server actions are suppressed"
    )
    ws["A4"].fill = PatternFill("solid", fgColor="D9EAD3" if latest_complete else "F4CCCC")
    ws["A4"].font = Font(bold=True)
    ws["A4"].alignment = Alignment(horizontal="center")

    group = group_rows[0] if group_rows else None
    kpi_fields = [
        ("gross_sales", "Sales"), ("guest_count", "Guests"), ("check_average", "Check Avg"),
        ("wine_pct", "Wine %"), ("rate_of_sale_by_guest_count", "Rate of Sale"),
        ("average_ticket_time_seconds", "Ticket Time"),
    ]
    for index, (field, label) in enumerate(kpi_fields):
        start_col = index * 2 + 1
        ws.merge_cells(start_row=6, start_column=start_col, end_row=6, end_column=start_col + 1)
        ws.merge_cells(start_row=7, start_column=start_col, end_row=7, end_column=start_col + 1)
        label_cell = ws.cell(row=6, column=start_col, value=label)
        label_cell.fill = PatternFill("solid", fgColor="E7E6E6")
        label_cell.font = Font(bold=True)
        label_cell.alignment = Alignment(horizontal="center")
        if group:
            current_value, number_format = format_management_value(field, group["latest"][field])
            value_cell = ws.cell(row=7, column=start_col, value=current_value)
            value_cell.number_format = number_format
            value_cell.font = Font(size=16, bold=True, color="7A1E1E")
            value_cell.alignment = Alignment(horizontal="center")
            ws.cell(row=8, column=start_col, value="vs prior")
            prior_text = signed_management_delta(field, group["prior_changes"][field])
            benchmark_text = signed_management_delta(field, group["benchmark_changes"][field])
            if field in {"gross_sales", "guest_count"}:
                prior_pct = safe_pct_delta(
                    group["latest"][field], group["prior"][field] if group["prior"] else None
                )
                benchmark_pct = safe_pct_delta(
                    group["latest"][field], group["benchmark_values"][field]
                )
                if prior_pct is not None:
                    prior_text += f" ({prior_pct:+.1%})"
                if benchmark_pct is not None:
                    benchmark_text += f" ({benchmark_pct:+.1%})"
            ws.cell(row=8, column=start_col + 1, value=prior_text)
            ws.cell(row=9, column=start_col, value="vs benchmark")
            ws.cell(row=9, column=start_col + 1, value=benchmark_text)
        for row in (8, 9):
            ws.cell(row=row, column=start_col).font = Font(size=9, color="666666")
            ws.cell(row=row, column=start_col + 1).alignment = Alignment(horizontal="right")

    ws.merge_cells("A11:L11")
    ws["A11"] = "Store Pulse"
    ws["A11"].fill = PatternFill("solid", fgColor="E7E6E6")
    ws["A11"].font = Font(bold=True)
    store_headers = ["Location", "Status", "Sales vs Benchmark", "Guests vs Benchmark", "Check", "Wine", "Ticket"]
    for col, header in enumerate(store_headers, start=1):
        cell = ws.cell(row=12, column=col, value=header)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.merge_cells("H12:L12")
    ws["H12"] = "Management Focus"
    ws["H12"].fill = PatternFill("solid", fgColor="D9E1F2")
    ws["H12"].font = Font(bold=True)
    ws["H12"].alignment = Alignment(horizontal="center")
    for row_index, item in enumerate(store_rows, start=13):
        latest = item["latest"]
        values = [
            item["entity"], item["status"],
            safe_pct_delta(latest["gross_sales"], item["benchmark_values"]["gross_sales"]),
            safe_pct_delta(latest["guest_count"], item["benchmark_values"]["guest_count"]),
            item["benchmark_changes"]["check_average"], item["benchmark_changes"]["wine_pct"],
            item["benchmark_changes"]["average_ticket_time_seconds"] / 60 if item["benchmark_changes"]["average_ticket_time_seconds"] is not None else None,
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_index, column=col, value=value)
        ws.merge_cells(start_row=row_index, start_column=8, end_row=row_index, end_column=12)
        ws.cell(row=row_index, column=8, value=item["recommended_focus"])
        ws.cell(row=row_index, column=3).number_format = "0.0%"
        ws.cell(row=row_index, column=4).number_format = "0.0%"
        ws.cell(row=row_index, column=5).number_format = "$0.00"
        ws.cell(row=row_index, column=6).number_format = "0.0%"
        ws.cell(row=row_index, column=7).number_format = "0.0"
        ws.cell(row=row_index, column=8).alignment = Alignment(wrap_text=True)
        fill = priority_fill(item["priority"])
        if fill:
            ws.cell(row=row_index, column=2).fill = fill
        ws.row_dimensions[row_index].height = 38

    for start_col, title, selected in (
        (1, "Act First", [row for row in action_rows if row.get("Priority") in {"High", "Medium", "Review"}][:3]),
        (7, "Recognize / Replicate", [row for row in action_rows if row.get("Priority") in {"Recognize", "Share"}][:3]),
    ):
        ws.merge_cells(start_row=17, start_column=start_col, end_row=17, end_column=start_col + 5)
        ws.cell(row=17, column=start_col, value=title).fill = PatternFill("solid", fgColor="E7E6E6")
        ws.cell(row=17, column=start_col).font = Font(bold=True)
        ws.cell(row=18, column=start_col, value="Priority")
        ws.merge_cells(start_row=18, start_column=start_col + 1, end_row=18, end_column=start_col + 2)
        ws.cell(row=18, column=start_col + 1, value="Person / Location")
        ws.merge_cells(start_row=18, start_column=start_col + 3, end_row=18, end_column=start_col + 5)
        ws.cell(row=18, column=start_col + 3, value="Management Move")
        for col in range(start_col, start_col + 6):
            cell = ws.cell(row=18, column=col)
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        if not selected:
            selected = [{"Priority": "Monitor", "Location": "", "Person / Area": "No current items", "Signal": "", "Why It Matters": "", "Recommended Next Step": ""}]
        for row_offset, item in enumerate(selected, start=19):
            ws.cell(row=row_offset, column=start_col, value=item.get("Priority"))
            ws.merge_cells(start_row=row_offset, start_column=start_col + 1, end_row=row_offset, end_column=start_col + 2)
            ws.cell(
                row=row_offset,
                column=start_col + 1,
                value=f"{item.get('Person / Area', '')}\n{item.get('Location', '')}".strip(),
            )
            ws.merge_cells(start_row=row_offset, start_column=start_col + 3, end_row=row_offset, end_column=start_col + 5)
            evidence = str(item.get("Why It Matters") or "").split(" | ")[0]
            move = f"{item.get('Signal', '')}: {evidence}. Next: {item.get('Recommended Next Step', '')}".strip()
            ws.cell(row=row_offset, column=start_col + 3, value=move)
            for col in range(start_col, start_col + 6):
                ws.cell(row=row_offset, column=col).alignment = Alignment(wrap_text=True, vertical="top")
            fill = priority_fill(item.get("Priority"))
            if fill:
                ws.cell(row=row_offset, column=start_col).fill = fill
            ws.row_dimensions[row_offset].height = 68

    ws.merge_cells("A24:L24")
    full_by_location, global_full = full_week_ends_by_location(weekly_location_rows)
    ws["A24"] = (
        f"Data quality: latest week {'complete' if latest_complete else 'incomplete'}; "
        f"{len(global_full)} full all-store weeks available; partial weeks are excluded from benchmarks."
    )
    ws["A24"].fill = PatternFill("solid", fgColor="D9EAF7")
    ws["A24"].alignment = Alignment(horizontal="center")

    week_labels = " | ".join(
        week_end.strftime("%m/%d")
        for week_end in sorted({row["week_end"] for row in weekly_location_rows})[-5:]
    )
    for start_col in (1, 7):
        ws.merge_cells(start_row=25, start_column=start_col, end_row=25, end_column=start_col + 5)
        label_cell = ws.cell(row=25, column=start_col, value=f"Weeks left to right: {week_labels}")
        label_cell.font = Font(size=8, color="666666", italic=True)
        label_cell.alignment = Alignment(horizontal="center")

    location_points, group_points = write_management_chart_data(wb, weekly_location_rows, weekly_group_rows)
    chart_ws = wb["_Dashboard Chart Data"]
    if location_points:
        chart = LineChart()
        chart.title = "Five-Week Sales Trend by Store"
        chart.y_axis.title = "Gross Sales"
        chart.y_axis.numFmt = "$#,##0"
        chart.x_axis.tickLblPos = "low"
        chart.height = 7
        chart.width = 11.5
        data = Reference(chart_ws, min_col=2, max_col=1 + len({row['location'] for row in weekly_location_rows}), min_row=1, max_row=1 + location_points)
        cats = Reference(chart_ws, min_col=1, min_row=2, max_row=1 + location_points)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        category_formula = "'_Dashboard Chart Data'!$A$2:$A$" + str(1 + location_points)
        for series in chart.series:
            series.cat = AxDataSource(strRef=StrRef(f=category_formula))
        for series, color in zip(chart.series, ("4472C4", "C0504D")):
            series.graphicalProperties.line.solidFill = color
            series.graphicalProperties.line.width = 28575
        chart.legend.position = "r"
        ws.add_chart(chart, "A26")
    if group_points:
        chart = LineChart()
        chart.title = "Five-Week All-Stores Guest Trend"
        chart.y_axis.title = "Guests"
        chart.y_axis.numFmt = "#,##0"
        chart.x_axis.tickLblPos = "low"
        chart.height = 7
        chart.width = 11.5
        data = Reference(chart_ws, min_col=7, min_row=1, max_row=1 + group_points)
        cats = Reference(chart_ws, min_col=6, min_row=2, max_row=1 + group_points)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        category_formula = "'_Dashboard Chart Data'!$F$2:$F$" + str(1 + group_points)
        for series in chart.series:
            series.cat = AxDataSource(strRef=StrRef(f=category_formula))
        if chart.series:
            chart.series[0].graphicalProperties.line.solidFill = "7A1E1E"
            chart.series[0].graphicalProperties.line.width = 28575
        chart.legend = None
        ws.add_chart(chart, "G26")
    ws.sheet_view.zoomScale = 85


def write_management_run_notes(
    wb: Workbook,
    records: list[MetricRecord],
    source_dir: Path,
    public_start: date,
    public_end: date,
    config: dict[str, Any],
) -> None:
    remove_sheet_if_present(wb, "Run Notes")
    ws = wb.create_sheet("Run Notes")
    style_management_title(ws, "Red Onion Server Master", 2)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 96
    note_rows = [
        ("Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Source Folder", str(source_dir)),
        ("Operating Week", f"{OPERATING_WEEK_LABEL}; Mondays are closed."),
        ("Raw Reports Read", len({record.source_file for record in records})),
        ("Date Coverage", format_date_range(min(r.report_date for r in records), max(r.report_date for r in records))),
        ("Public Snapshot Dates", format_date_range(public_start, public_end)),
        ("Management Baseline", f"Up to {config.get('dashboard_baseline_full_weeks', 4)} prior full Tuesday-Sunday weeks; partial weeks excluded."),
        ("Prominent Server Confidence", f"Latest guests >= {config.get('dashboard_min_guest_count_for_trends', 25)} AND active days >= {config.get('dashboard_min_active_days_for_trends', 3)}, plus at least {config.get('dashboard_min_prior_full_weeks', 2)} prior full weeks and {config.get('dashboard_min_prior_guest_count', 50)} prior guests."),
        ("Targets", "Management Setup targets take precedence; blank targets use the rolling baseline. Edits apply on the next run."),
        ("Momentum", "Four metric families score from -2 to +2 using materiality bands; average rank movement contributes only -1, 0, or +1."),
        ("Performance Level", "Latest metrics are assessed separately as Above Benchmark, On Track, or Below Benchmark."),
        ("Action Tracking", "Owner, due date, status, and manager notes carry forward between weekly runs. Cleared signals move to Action History."),
        ("Metric Rule", "Check average and wine percent are recalculated from rolled-up sales, guests, and wine sales."),
        ("Metric Rule", "Rate of sale and ticket time are guest-weighted averages."),
    ]
    for row, (label, value) in enumerate(note_rows, start=4):
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=2, value=value).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30


def finalize_management_workbook(wb: Workbook) -> None:
    for sheet in wb.worksheets:
        if sheet.title in VISIBLE_MANAGEMENT_SHEETS:
            sheet.sheet_state = "visible"
            add_management_navigation(sheet)
        else:
            sheet.sheet_state = "hidden"
    ordered = [wb[name] for name in VISIBLE_MANAGEMENT_SHEETS if name in wb.sheetnames]
    ordered.extend(sheet for sheet in wb.worksheets if sheet.title not in VISIBLE_MANAGEMENT_SHEETS)
    wb._sheets = ordered
    wb.active = 0
    tab_colors = {
        "Dashboard": "7A1E1E", "Action Board": "C00000", "Server Scorecard": "5B9BD5",
        "Store & Group Scorecards": "70AD47", "Rising & Falling Stars": "FFC000",
        "Action History": "A5A5A5", "Data Quality": "5B9BD5", "Management Setup": "4472C4",
        "Run Notes": "7F7F7F",
    }
    for name, color in tab_colors.items():
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = color


def write_master_workbook(
    records: list[MetricRecord],
    output_path: Path,
    config: dict[str, Any],
    source_dir: Path,
    public_start: date,
    public_end: date,
) -> Path:
    state = read_management_state(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_name(f".{output_path.stem}.{os.getpid()}.tmp.xlsx")
    if temp_path.exists():
        temp_path.unlink()
    wb: Workbook | None = None
    try:
        _write_master_workbook_base(records, temp_path, config, source_dir, public_start, public_end)
        wb = load_workbook(temp_path)
        weekly_server_rows, weekly_location_rows = weekly_rollups(records)
        rank_min_guest_count = int(
            config.get("master_min_guest_count_for_rankings", config.get("public_min_guest_count", 1))
        )
        ranked_rows = weekly_server_rank_rows(weekly_server_rows, rank_min_guest_count)
        weekly_group_rows = group_weekly_rows(records)
        full_by_location, global_full = full_week_ends_by_location(weekly_location_rows)
        server_rows = management_server_rows(
            weekly_server_rows, weekly_location_rows, ranked_rows, state["targets"], config
        )
        store_rows = management_entity_rows(
            weekly_location_rows, "location", state["targets"], config, full_by_location
        )
        group_rows = management_entity_rows(
            weekly_group_rows, "group", state["targets"], config, global_full
        )
        signals = build_management_action_signals(
            server_rows, store_rows, group_rows, weekly_location_rows
        )
        current_actions, action_history = merge_management_actions(signals, state)

        if "Data Quality" in wb.sheetnames:
            remove_sheet_if_present(wb, "_Data Quality Detail")
            wb["Data Quality"].title = "_Data Quality Detail"
        for name in (
            "Dashboard", "Action Board", "Rising & Falling Stars", "Run Notes",
            "Server Scorecard", "Store & Group Scorecards", "Action History", "Management Setup",
        ):
            remove_sheet_if_present(wb, name)
        write_management_setup_sheet(wb, state["targets"], state["owners"], config)
        write_action_tracking_sheet(wb, "Action Board", current_actions, editable=True)
        write_server_scorecard_sheet(wb, server_rows)
        write_store_group_scorecards_sheet(wb, [*group_rows, *store_rows], config)
        write_rising_falling_sheet(wb, server_rows)
        write_action_tracking_sheet(wb, "Action History", action_history, editable=False)
        write_management_data_quality_sheet(wb, weekly_location_rows)
        write_management_run_notes(wb, records, source_dir, public_start, public_end, config)
        write_management_dashboard_sheet(
            wb, records, weekly_location_rows, weekly_group_rows, server_rows,
            store_rows, group_rows, current_actions,
        )
        finalize_management_workbook(wb)
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
        wb.save(temp_path)
        wb.close()
        wb = None
        validation = load_workbook(temp_path, read_only=True, data_only=False)
        missing = [name for name in VISIBLE_MANAGEMENT_SHEETS if name not in validation.sheetnames]
        validation.close()
        if missing:
            raise RuntimeError(f"Generated master workbook is missing required sheets: {', '.join(missing)}")
        os.replace(temp_path, output_path)
    except PermissionError as exc:
        raise RuntimeError(
            f"Could not replace {output_path.name}. Close the master workbook in Excel and rerun; "
            "no source files were moved."
        ) from exc
    finally:
        if wb is not None:
            wb.close()
        if temp_path.exists():
            temp_path.unlink()
    return output_path


def run(args: argparse.Namespace) -> list[Path]:
    input_dir = Path(args.input_dir).resolve()
    output_dir = Path(args.output_dir).resolve()
    archive_dir = Path(args.archive_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    config_path = Path(args.config).resolve()
    config = load_config(config_path)
    active_paths = daily_report_paths(input_dir)
    if not active_paths:
        raise FileNotFoundError(
            f"No active daily .xls reports found in {input_dir}. "
            "Drop current Toast reports into Daily Reports and rerun."
        )

    active_records_by_path = read_reports_by_path(active_paths, config)
    _, active_week_end = active_week_for_paths(active_records_by_path)
    active_records = flatten_report_records(active_records_by_path)
    active_path_set = set(active_paths)
    archived_paths = [
        path for path in archived_daily_report_paths(archive_dir) if path not in active_path_set
    ]
    archived_records_by_path = read_reports_by_path(archived_paths, config) if archived_paths else {}
    records = flatten_report_records(archived_records_by_path) + active_records

    public_start, public_end = selected_public_dates(active_records, args.week_start, args.week_end)
    selected_records = [
        record
        for record in records
        if public_start <= record.report_date <= public_end and is_operating_day(record.report_date)
    ]
    if not selected_records:
        raise ValueError(f"No records found between {public_start} and {public_end}.")

    generated: list[Path] = []
    for location in config["locations"]:
        location_records = [record for record in selected_records if record.location == location]
        if not location_records:
            continue
        generated.append(
            write_public_workbook(
                location, selected_records, output_dir, config, public_start, public_end
            )
        )

    master_path = output_dir / "Red_Onion_Server_Master.xlsx"
    generated.append(
        write_master_workbook(
            records,
            master_path,
            config,
            input_dir,
            public_start,
            public_end,
        )
    )
    archive_processed_files(active_paths, archive_dir, active_week_end)
    return generated


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build Red Onion weekly metric workbooks.")
    parser.add_argument(
        "--input-dir",
        default=str(DEFAULT_INPUT_DIR),
        help="Folder containing active raw daily .xls reports.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help="Folder for generated workbooks.",
    )
    parser.add_argument(
        "--archive-dir",
        default=str(DEFAULT_ARCHIVE_DIR),
        help="Folder for archived source files.",
    )
    parser.add_argument("--config", default=str(DEFAULT_CONFIG_PATH), help="Path to config JSON.")
    parser.add_argument("--week-start", help="Optional public snapshot start date, YYYY-MM-DD.")
    parser.add_argument("--week-end", help="Optional public snapshot end date, YYYY-MM-DD.")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    try:
        generated = run(args)
    except Exception as exc:
        raise SystemExit(f"ERROR: {exc}") from exc
    print("Generated:")
    for path in generated:
        print(f"  {path}")


if __name__ == "__main__":
    main()
