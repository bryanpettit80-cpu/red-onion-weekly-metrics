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
from openpyxl.chart.series import SeriesLabel
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

SERVER_TREND_FIELDS: tuple[str, ...] = (
    "check_average",
    "wine_pct",
    "rate_of_sale_by_guest_count",
    "average_ticket_time_seconds",
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
DAILY_REPORT_PREFIX = "daily report"
DAILY_REPORT_EXTENSIONS = frozenset({".xls", ".xlsx"})
DAILY_REPORT_FORMAT_LABEL = ".xls or .xlsx"
CANONICAL_DAILY_ARCHIVE_FOLDER = "processed-daily-reports"
PROGRAM_DIR = Path(__file__).resolve().parent
REPOSITORY_ROOT = PROGRAM_DIR.parent
PROJECT_ROOT = (
    REPOSITORY_ROOT.parent
    if REPOSITORY_ROOT.name == "Red Onion Weekly Metrics Automation"
    else REPOSITORY_ROOT
)
DEFAULT_INPUT_DIR = PROJECT_ROOT / "01 Daily Reports - Drop Here"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "02 Finished Reports"
DEFAULT_ARCHIVE_DIR = PROJECT_ROOT / "03 Archive"
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


@dataclass(frozen=True)
class ReportResolution:
    records_by_path: dict[Path, list[MetricRecord]]
    duplicate_paths: tuple[Path, ...]
    business_dates: tuple[date, ...]


@dataclass(frozen=True)
class HistoryMigrationPlan:
    copy_pairs: tuple[tuple[Path, Path], ...]
    effective_records_by_path: dict[Path, list[MetricRecord]]
    duplicate_paths: tuple[Path, ...]
    business_dates: tuple[date, ...]


@dataclass(frozen=True)
class HistoryMigrationResult:
    copied_paths: tuple[Path, ...]
    duplicate_files_ignored: int
    business_dates_considered: int


@dataclass(frozen=True)
class LatestWeekReadiness:
    latest_week_end: date | None
    latest_location_rows: tuple[dict[str, Any], ...]
    configured_locations: tuple[str, ...]
    location_gaps: tuple[str, ...]
    expected_dates: tuple[date, ...]
    received_dates: frozenset[date]
    missing_dates: tuple[date, ...]
    ready: bool

    @property
    def missing_parts(self) -> tuple[str, ...]:
        return (
            *(f"{report_date:%b} {report_date.day}" for report_date in self.missing_dates),
            *(f"{location} data" for location in self.location_gaps),
        )

    @property
    def missing_text(self) -> str:
        return ", ".join(self.missing_parts)


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


def daily_report_excel_engine(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        return "xlrd"
    if suffix == ".xlsx":
        return "openpyxl"
    raise ValueError(
        f"Unsupported daily report format for {path.name}. Use {DAILY_REPORT_FORMAT_LABEL}."
    )


def parse_daily_report(path: Path, config: dict[str, Any]) -> list[MetricRecord]:
    with pd.ExcelFile(path, engine=daily_report_excel_engine(path)) as workbook:
        if "Report(All)" not in workbook.sheet_names:
            if "No Data Available" in workbook.sheet_names:
                raise ValueError(
                    "the Toast export contains a 'No Data Available' worksheet instead of report "
                    "data. Replace it with a complete Daily Report export, then rerun. No "
                    "workbooks were created and no source files were moved."
                )
            available_sheets = ", ".join(workbook.sheet_names) or "none"
            raise ValueError(
                "the required 'Report(All)' worksheet is missing "
                f"(worksheets found: {available_sheets}). Replace the file with a complete Toast "
                "Daily Report export, then rerun."
            )
        df = workbook.parse("Report(All)", header=None)
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


def is_daily_report_path(path: Path) -> bool:
    return (
        path.is_file()
        and not path.name.startswith("~$")
        and path.name.casefold().startswith(DAILY_REPORT_PREFIX)
        and path.suffix.lower() in DAILY_REPORT_EXTENSIONS
    )


def find_daily_report_paths(root: Path, *, recursive: bool) -> list[Path]:
    if not root.exists() or not root.is_dir():
        return []
    candidates = root.rglob("*") if recursive else root.iterdir()
    return sorted(
        (path for path in candidates if is_daily_report_path(path)),
        key=lambda path: str(path).casefold(),
    )


def daily_report_paths(input_dir: Path) -> list[Path]:
    return find_daily_report_paths(input_dir, recursive=False)


def canonical_daily_archive_dir(archive_dir: Path) -> Path:
    return archive_dir / CANONICAL_DAILY_ARCHIVE_FOLDER


def archived_daily_report_paths(archive_dir: Path) -> list[Path]:
    return find_daily_report_paths(canonical_daily_archive_dir(archive_dir), recursive=True)


def read_reports_by_path(paths: Iterable[Path], config: dict[str, Any]) -> dict[Path, list[MetricRecord]]:
    records_by_path: dict[Path, list[MetricRecord]] = {}
    for path in paths:
        try:
            records = parse_daily_report(path, config)
        except Exception as exc:
            raise ValueError(f"Could not process {path.name}: {exc}") from exc
        if not records:
            raise ValueError(f"No metric rows were found in {path}")
        records_by_path[path] = records
    return records_by_path


def report_date_for_records(path: Path, records: list[MetricRecord]) -> date:
    report_dates = sorted({record.report_date for record in records})
    if len(report_dates) != 1:
        raise ValueError(f"{path.name} contains multiple report dates: {report_dates}")
    return report_dates[0]


def semantic_report_signature(records: list[MetricRecord]) -> tuple[tuple[Any, ...], ...]:
    def normalized_number(value: float) -> float:
        number = round(float(value), 6)
        return 0.0 if number == 0 else number

    return tuple(
        sorted(
            (
                record.location.strip(),
                record.raw_user_name.strip(),
                bool(record.is_location_total),
                normalized_number(record.gross_sales),
                normalized_number(record.guest_count),
                normalized_number(record.wine_sales),
                normalized_number(record.rate_of_sale_by_guest_count),
                normalized_number(record.average_ticket_time_seconds),
            )
            for record in records
        )
    )


def resolve_report_duplicates(
    records_by_path: dict[Path, list[MetricRecord]],
) -> ReportResolution:
    reports_by_date: dict[date, list[tuple[Path, list[MetricRecord]]]] = defaultdict(list)
    for path, records in records_by_path.items():
        reports_by_date[report_date_for_records(path, records)].append((path, records))

    resolved: dict[Path, list[MetricRecord]] = {}
    duplicates: list[Path] = []
    for report_date, reports in sorted(reports_by_date.items()):
        reports.sort(key=lambda item: str(item[0]).casefold())
        selected_path, selected_records = reports[0]
        selected_signature = semantic_report_signature(selected_records)
        conflicting_paths = [
            path
            for path, records in reports[1:]
            if semantic_report_signature(records) != selected_signature
        ]
        if conflicting_paths:
            paths = [selected_path, *conflicting_paths]
            path_lines = "\n".join(f"  - {path}" for path in paths)
            raise ValueError(
                "Conflicting daily reports were found for business date "
                f"{report_date.isoformat()}. These files contain different metric data:\n"
                f"{path_lines}\n"
                "Keep only the correct source or reconcile the reports before rerunning. "
                "No history files were copied, no workbooks were created, and no active "
                "source files were moved."
            )
        resolved[selected_path] = selected_records
        duplicates.extend(path for path, _ in reports[1:])

    return ReportResolution(
        records_by_path=resolved,
        duplicate_paths=tuple(duplicates),
        business_dates=tuple(sorted(reports_by_date)),
    )


def flatten_report_records(records_by_path: dict[Path, list[MetricRecord]]) -> list[MetricRecord]:
    records: list[MetricRecord] = []
    for path_records in records_by_path.values():
        records.extend(path_records)
    return records


def read_all_reports(input_dir: Path, config: dict[str, Any]) -> list[MetricRecord]:
    paths = daily_report_paths(input_dir)
    if not paths:
        raise FileNotFoundError(
            f"No daily reports ({DAILY_REPORT_FORMAT_LABEL}) found in {input_dir}"
        )

    resolution = resolve_report_duplicates(read_reports_by_path(paths, config))
    return flatten_report_records(resolution.records_by_path)


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
            "The daily-report drop folder contains files from more than one Tuesday-Sunday operating week.",
            "Move the extra files out of 01 Daily Reports - Drop Here and rerun:",
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


def archive_destination_for(
    source_path: Path,
    archive_dir: Path,
    reserved_destinations: set[Path] | None = None,
) -> tuple[Path, bool]:
    reserved_destinations = reserved_destinations or set()
    destination = archive_dir / source_path.name
    if not destination.exists() and destination not in reserved_destinations:
        return destination, False

    if destination.exists() and filecmp.cmp(source_path, destination, shallow=False):
        return destination, True

    counter = 1
    while True:
        candidate = archive_dir / f"{source_path.stem} ({counter}){source_path.suffix}"
        if not candidate.exists() and candidate not in reserved_destinations:
            return candidate, False
        if candidate.exists() and filecmp.cmp(source_path, candidate, shallow=False):
            return candidate, True
        counter += 1


def archive_processed_files(
    source_paths: Iterable[Path], archive_root: Path, week_end: date
) -> list[Path]:
    archive_dir = canonical_daily_archive_dir(archive_root) / f"week-ending-{week_end.isoformat()}"
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


def migration_daily_report_paths(source_dirs: Iterable[Path]) -> list[Path]:
    paths: dict[Path, None] = {}
    for source_dir in source_dirs:
        source_dir = source_dir.resolve()
        if not source_dir.exists() or not source_dir.is_dir():
            raise FileNotFoundError(f"History migration source folder was not found: {source_dir}")
        source_paths = find_daily_report_paths(source_dir, recursive=True)
        if not source_paths:
            raise FileNotFoundError(
                f"No daily reports ({DAILY_REPORT_FORMAT_LABEL}) were found under history "
                f"migration source: {source_dir}"
            )
        for path in source_paths:
            paths[path.resolve()] = None
    return sorted(paths, key=lambda path: str(path).casefold())


def build_history_migration_plan(
    source_dirs: Iterable[Path],
    archive_root: Path,
    config: dict[str, Any],
) -> HistoryMigrationPlan:
    source_paths = migration_daily_report_paths(source_dirs)
    source_records_by_path = read_reports_by_path(source_paths, config)
    canonical_paths = archived_daily_report_paths(archive_root)
    canonical_records_by_path = (
        read_reports_by_path(canonical_paths, config) if canonical_paths else {}
    )

    combined_records_by_path = dict(canonical_records_by_path)
    combined_records_by_path.update(source_records_by_path)
    resolution = resolve_report_duplicates(combined_records_by_path)

    canonical_dates = {
        report_date_for_records(path, records)
        for path, records in canonical_records_by_path.items()
    }
    source_paths_by_date: dict[date, list[Path]] = defaultdict(list)
    for path, records in source_records_by_path.items():
        source_paths_by_date[report_date_for_records(path, records)].append(path)

    copy_pairs: list[tuple[Path, Path]] = []
    reserved_destinations: set[Path] = set()
    for report_date, paths in sorted(source_paths_by_date.items()):
        if report_date in canonical_dates:
            continue
        source_path = min(
            paths,
            key=lambda path: (path.suffix.lower() != ".xlsx", str(path).casefold()),
        )
        _, week_end = week_period_for(report_date)
        destination_dir = canonical_daily_archive_dir(archive_root) / (
            f"week-ending-{week_end.isoformat()}"
        )
        destination, already_present = archive_destination_for(
            source_path, destination_dir, reserved_destinations
        )
        if not already_present:
            copy_pairs.append((source_path, destination))
            reserved_destinations.add(destination)

    return HistoryMigrationPlan(
        copy_pairs=tuple(copy_pairs),
        effective_records_by_path=resolution.records_by_path,
        duplicate_paths=resolution.duplicate_paths,
        business_dates=resolution.business_dates,
    )


def apply_history_migration_plan(plan: HistoryMigrationPlan) -> tuple[Path, ...]:
    destinations = [destination for _, destination in plan.copy_pairs]
    if len(set(destinations)) != len(destinations):
        raise RuntimeError(
            "History migration preflight selected the same destination more than once. "
            "No files were copied."
        )
    conflicts = [destination for destination in destinations if destination.exists()]
    if conflicts:
        raise RuntimeError(
            f"History migration destination appeared after preflight: {conflicts[0]}. "
            "No files were copied; rerun the migration."
        )

    copied_paths: list[Path] = []
    for source_path, destination in plan.copy_pairs:
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, destination)
        copied_paths.append(destination)
    return tuple(copied_paths)


def migrate_history_files(
    source_dirs: Iterable[Path],
    archive_root: Path,
    config: dict[str, Any],
) -> HistoryMigrationResult:
    plan = build_history_migration_plan(source_dirs, archive_root, config)
    copied_paths = apply_history_migration_plan(plan)
    return HistoryMigrationResult(
        copied_paths=copied_paths,
        duplicate_files_ignored=len(plan.duplicate_paths),
        business_dates_considered=len(plan.business_dates),
    )


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
    if text in {
        "high", "coach", "coach now", "falling", "declining", "falling star",
        "below benchmark", "traffic watch", "incomplete week",
    }:
        return soft_fill("F4CCCC")
    if text in {
        "medium", "coach fundamentals", "protect performance", "reinforce improvement",
        "store review", "group review", "upsell watch", "service watch", "mixed watch",
        "building history", "low sample",
    }:
        return soft_fill("FFF2CC")
    if text in {
        "recognize", "recognize & replicate", "share", "rising", "improving",
        "rising star", "above benchmark",
    }:
        return soft_fill("D9EAD3")
    if text in {"review", "data quality", "short week"}:
        return soft_fill("D9EAF7")
    if text in {"monitor", "stable", "stable / mixed", "on track", "not scored"}:
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
        "Server Week-over-Week Detail",
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


def management_trend_classification(
    metric_scores: dict[str, int],
    rank_modifier: int,
    *,
    improving_label: str,
    declining_label: str,
) -> tuple[str, int]:
    composite_score = sum(metric_scores.values()) + rank_modifier
    positive_count = sum(score > 0 for score in metric_scores.values())
    negative_count = sum(score < 0 for score in metric_scores.values())
    if composite_score >= 3 and positive_count >= 2:
        return improving_label, composite_score
    if composite_score <= -3 and negative_count >= 2:
        return declining_label, composite_score
    return "Stable", composite_score


def management_rank_block_movement(
    recent_rows: list[dict[str, Any]],
    comparison_rows: list[dict[str, Any]],
    rank_lookup: dict[tuple[date, str, str], dict[str, Any]],
) -> tuple[float | None, int]:
    movements: list[float] = []
    for rank_field in (
        "check_average_rank",
        "wine_pct_rank",
        "rate_rank",
        "ticket_time_rank",
    ):
        recent_ranks = [
            rank_lookup[(row["week_end"], row["location"], row["raw_user_name"])].get(rank_field)
            for row in recent_rows
            if (row["week_end"], row["location"], row["raw_user_name"]) in rank_lookup
            and rank_lookup[(row["week_end"], row["location"], row["raw_user_name"])].get(rank_field)
            is not None
        ]
        comparison_ranks = [
            rank_lookup[(row["week_end"], row["location"], row["raw_user_name"])].get(rank_field)
            for row in comparison_rows
            if (row["week_end"], row["location"], row["raw_user_name"]) in rank_lookup
            and rank_lookup[(row["week_end"], row["location"], row["raw_user_name"])].get(rank_field)
            is not None
        ]
        if recent_ranks and comparison_ranks:
            movements.append(
                sum(float(value) for value in comparison_ranks) / len(comparison_ranks)
                - sum(float(value) for value in recent_ranks) / len(recent_ranks)
            )
    average_movement = sum(movements) / len(movements) if movements else None
    modifier = (
        1
        if average_movement is not None and average_movement >= 3
        else -1
        if average_movement is not None and average_movement <= -3
        else 0
    )
    return average_movement, modifier


def server_trend_context(recent_momentum: str, long_term_direction: str) -> str:
    contexts = {
        ("Rising", "Improving"): "The recent gain continues a sustained longer-term improvement.",
        ("Rising", "Declining"): "This is a recent rebound against a softer longer-term trend.",
        ("Falling", "Improving"): "This is recent softening within a stronger longer-term trend.",
        ("Falling", "Declining"): "The recent decline continues a sustained longer-term decline.",
        ("Stable", "Improving"): "Recent results are stable while the longer-term direction is improving.",
        ("Stable", "Declining"): "Recent results are stable while the longer-term direction is declining.",
    }
    return contexts.get((recent_momentum, long_term_direction), "")


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
    long_term_limit = int(config.get("dashboard_long_term_full_weeks", 8))
    long_term_block = int(config.get("dashboard_long_term_block_weeks", 4))
    full_min_recent_guests = float(
        config.get("dashboard_long_term_full_min_recent_guests", 100)
    )
    full_min_earlier_guests = float(
        config.get("dashboard_long_term_full_min_earlier_guests", 100)
    )
    developing_min_total = int(
        config.get("dashboard_long_term_developing_min_total_weeks", 6)
    )
    developing_min_recent = int(
        config.get("dashboard_long_term_developing_min_recent_weeks", 3)
    )
    developing_min_earlier = int(
        config.get("dashboard_long_term_developing_min_earlier_weeks", 2)
    )
    developing_min_recent_guests = float(
        config.get("dashboard_long_term_developing_min_recent_guests", 75)
    )
    developing_min_earlier_guests = float(
        config.get("dashboard_long_term_developing_min_earlier_guests", 50)
    )
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
        current_sample_eligible = dashboard_trend_eligible(current, config)
        prominent = (
            full_latest
            and current_sample_eligible
            and len(prior_rows) >= min_prior_weeks
            and prior_guest_count >= min_prior_guests
        )

        metric_scores: dict[str, int] = {}
        changes: dict[str, float | None] = {}
        positive_drivers: list[str] = []
        negative_drivers: list[str] = []
        for field in SERVER_TREND_FIELDS:
            change = current[field] - baseline[field] if baseline else None
            score = directional_score(change, management_threshold(config, field))
            changes[field] = change
            metric_scores[field] = score
            if change is not None and score > 0:
                positive_drivers.append(metric_driver(field, change))
            elif change is not None and score < 0:
                negative_drivers.append(metric_driver(field, change))

        average_rank_move, rank_score = management_rank_block_movement(
            [current], prior_rows, rank_lookup
        )
        scored_momentum, composite_score = management_trend_classification(
            metric_scores,
            rank_score,
            improving_label="Rising",
            declining_label="Falling",
        )
        momentum = scored_momentum if prominent else "Not Scored"

        history_week_ends = sorted(
            week_end
            for week_end in full_by_location.get(location, set())
            if week_end <= latest_week_end
        )[-long_term_limit:]
        recent_history_ends = set(history_week_ends[-long_term_block:])
        earlier_history_ends = set(
            history_week_ends[-(long_term_block * 2) : -long_term_block]
        )
        server_history_rows = [
            row
            for row in weekly_server_rows
            if row["location"] == location
            and row["raw_user_name"] == current["raw_user_name"]
            and row["week_end"] in set(history_week_ends)
        ]
        recent_history_rows = [
            row for row in server_history_rows if row["week_end"] in recent_history_ends
        ]
        earlier_history_rows = [
            row for row in server_history_rows if row["week_end"] in earlier_history_ends
        ]
        recent_history_guests = sum(
            float(row.get("guest_count", 0) or 0) for row in recent_history_rows
        )
        earlier_history_guests = sum(
            float(row.get("guest_count", 0) or 0) for row in earlier_history_rows
        )
        history_total_weeks = len(server_history_rows)
        history_total_guests = recent_history_guests + earlier_history_guests
        full_history = (
            len(recent_history_rows) == long_term_block
            and len(earlier_history_rows) == long_term_block
            and history_total_weeks == long_term_block * 2
            and recent_history_guests >= full_min_recent_guests
            and earlier_history_guests >= full_min_earlier_guests
        )
        developing_history = (
            history_total_weeks >= developing_min_total
            and len(recent_history_rows) >= developing_min_recent
            and len(earlier_history_rows) >= developing_min_earlier
            and recent_history_guests >= developing_min_recent_guests
            and earlier_history_guests >= developing_min_earlier_guests
        )
        history_label = (
            "Full" if full_history else "Developing" if developing_history else "Building History"
        )
        history_through = max(
            (row["week_end"] for row in server_history_rows), default=None
        )
        history_used = (
            f"{history_total_weeks} weeks / {history_total_guests:,.0f} guests"
            + (f" through {history_through:%m/%d/%Y}" if history_through else "")
            + (
                f" (recent {len(recent_history_rows)}w / {recent_history_guests:,.0f}g; "
                f"earlier {len(earlier_history_rows)}w / {earlier_history_guests:,.0f}g)"
            )
        )
        recent_history = aggregate_weekly_rows(recent_history_rows)
        earlier_history = aggregate_weekly_rows(earlier_history_rows)
        long_term_changes: dict[str, float | None] = {}
        long_term_metric_scores: dict[str, int] = {}
        for field in SERVER_TREND_FIELDS:
            change = (
                recent_history[field] - earlier_history[field]
                if recent_history and earlier_history
                else None
            )
            long_term_changes[field] = change
            long_term_metric_scores[field] = directional_score(
                change, management_threshold(config, field)
            )
        long_term_average_rank_move, long_term_rank_modifier = management_rank_block_movement(
            recent_history_rows, earlier_history_rows, rank_lookup
        )
        scored_long_term_direction, long_term_composite_score = management_trend_classification(
            long_term_metric_scores,
            long_term_rank_modifier,
            improving_label="Improving",
            declining_label="Declining",
        )
        long_term_direction = (
            scored_long_term_direction
            if full_latest and history_label in {"Full", "Developing"}
            else "Not Scored"
        )

        benchmark_values: dict[str, float | None] = {}
        benchmark_sources: dict[str, str] = {}
        level_statuses: dict[str, str] = {}
        for field in SERVER_TREND_FIELDS:
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

        if not full_latest:
            priority, action, confidence = "Monitor", "Monitor", "Incomplete Week"
        elif not current_sample_eligible:
            priority, action, confidence = "Monitor", "Monitor", "Low Sample"
        elif not prominent:
            priority, action, confidence = "Monitor", "Monitor", "Building History"
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
        if long_term_direction != "Not Scored":
            why_parts.append(f"8-week direction {long_term_direction} ({history_label})")
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
            "long_term_direction": long_term_direction,
            "long_term_history_label": history_label,
            "history_used": history_used,
            "long_term_changes": long_term_changes,
            "long_term_metric_scores": long_term_metric_scores,
            "long_term_rank_modifier": long_term_rank_modifier,
            "long_term_average_rank_movement": long_term_average_rank_move,
            "long_term_composite_score": long_term_composite_score,
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
        context = server_trend_context(momentum, long_term_direction)
        if context:
            row["recommended_next_step"] = f"{row['recommended_next_step']} {context}"
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
    digest = hashlib.sha256(f"{entity_key}|{first_seen.isoformat()}".encode("utf-8")).hexdigest()
    return digest[:12].upper()


def compact_server_evidence(row: dict[str, Any]) -> str:
    parts = []
    if row["positive_drivers"]:
        parts.append("Improving: " + "; ".join(row["positive_drivers"][:2]))
    if row["negative_drivers"]:
        parts.append("Watch: " + "; ".join(row["negative_drivers"][:2]))
    if row.get("long_term_direction") != "Not Scored":
        parts.append(
            f"8-week {row['long_term_direction']} ({row['long_term_history_label']})"
        )
    parts.append(f"{row['guest_count']:,.0f} guests / {row['active_days']} days")
    return " | ".join(parts)


def build_management_action_signals(
    server_rows: list[dict[str, Any]],
    store_rows: list[dict[str, Any]],
    group_rows: list[dict[str, Any]],
    weekly_location_rows: list[dict[str, Any]],
    readiness: LatestWeekReadiness | None = None,
) -> list[dict[str, Any]]:
    signals: list[dict[str, Any]] = []
    analytical_ready = readiness is None or readiness.ready
    for row in server_rows if analytical_ready else []:
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
                "Signal": (
                    f"{row['momentum']} / {row['performance_level']} / "
                    f"8-week {row['long_term_direction']}"
                ),
                "Why It Matters": compact_server_evidence(row),
                "Recommended Next Step": row["recommended_next_step"],
                "Performance Level": row["performance_level"],
                "Momentum": row["momentum"],
                "Confidence": row["confidence"],
                "Last Seen": row["week_end"],
            }
        )
    analytical_rows = (("store", store_rows), ("group", group_rows)) if analytical_ready else ()
    for action_type, rows in analytical_rows:
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
    if readiness is not None and readiness.latest_week_end is not None:
        latest_by_location = {
            row["location"]: row for row in readiness.latest_location_rows
        }
        for location in readiness.location_gaps:
            row = latest_by_location.get(location)
            source_days = int(row.get("source_days", 0)) if row else 0
            detail = (
                f"{source_days} of {OPERATING_WEEK_DAYS} source days"
                if row
                else "No current-week store data"
            )
            entity_key = f"data-quality|{location}|latest-week".casefold()
            signals.append(
                {
                    "Entity Key": entity_key,
                    "Priority": "Review",
                    "Location": location,
                    "Person / Area": location,
                    "Action": "Data Quality",
                    "Signal": "Incomplete Latest Week",
                    "Why It Matters": detail,
                    "Recommended Next Step": "Confirm the missing reports before using trends for coaching.",
                    "Performance Level": "Preliminary",
                    "Momentum": "Not Scored",
                    "Confidence": "Low Sample",
                    "Last Seen": readiness.latest_week_end,
                }
            )
        if readiness.missing_dates and not readiness.location_gaps:
            signals.append(
                {
                    "Entity Key": "data-quality|all-stores|missing-daily-reports",
                    "Priority": "Review",
                    "Location": "All Stores",
                    "Person / Area": "All Stores",
                    "Action": "Data Quality",
                    "Signal": "Incomplete Latest Week",
                    "Why It Matters": f"Missing {readiness.missing_text}",
                    "Recommended Next Step": "Confirm the missing reports before using trends for coaching.",
                    "Performance Level": "Preliminary",
                    "Momentum": "Not Scored",
                    "Confidence": "Low Sample",
                    "Last Seen": readiness.latest_week_end,
                }
            )
    elif weekly_location_rows:
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
    signals: list[dict[str, Any]],
    state: dict[str, Any],
    readiness: LatestWeekReadiness | None = None,
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
        prior_status = str(prior.get("Status", "")).casefold()
        is_data_quality = (
            key.startswith("data-quality|")
            or str(prior.get("Action") or "").casefold() == "data quality"
        )
        if (
            readiness is not None
            and not readiness.ready
            and prior_status not in completed_statuses
            and not is_data_quality
        ):
            paused = dict(prior)
            prior_signal = str(paused.get("Signal") or "Prior action").strip()
            if not prior_signal.casefold().startswith("paused / carryover"):
                paused["Signal"] = f"PAUSED / CARRYOVER - {prior_signal}"
            missing_text = readiness.missing_text or "latest-week data"
            paused["Priority"] = "Paused"
            paused["Action"] = "Paused Carryover"
            paused["Why It Matters"] = (
                f"Prior action retained; the latest week is incomplete ({missing_text}) "
                "and was not used to confirm or clear it."
            )
            paused["Recommended Next Step"] = (
                "Keep the manual assignment on hold until a complete week can confirm, "
                "update, or clear the prior signal."
            )
            paused["Performance Level"] = "Preliminary"
            paused["Momentum"] = "Not Scored"
            paused["Confidence"] = "Paused"
            paused["Signal State"] = "Paused / Carryover"
            current.append(paused)
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
    priority_order = {
        "High": 0,
        "Medium": 1,
        "Recognize": 2,
        "Share": 3,
        "Review": 4,
        "Monitor": 5,
        "Paused": 6,
    }
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
        ws.cell(row=row, column=5, value="Recent and 8-week trend scoring; benchmark status")
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
        "Action",
        "Location",
        "Server",
        "Current Sample",
        "Confidence",
        "Performance",
        "Recent Momentum",
        "8-Week Direction",
        "History Used",
        "Positive Drivers",
        "Watch Drivers",
        "Recommended Next Step",
    ]
    data = []
    for row in rows:
        data.append(
            [
                row["action"],
                row["location"],
                row["display_name"],
                f"{row['guest_count']:,.0f} guests / {row['active_days']} days",
                row["confidence"],
                row["performance_level"],
                row["momentum"],
                row["long_term_direction"],
                row["history_used"],
                "; ".join(row["positive_drivers"]),
                "; ".join(row["negative_drivers"]),
                row["recommended_next_step"],
            ]
        )
    ws = write_table_sheet(
        wb, "Server Scorecard", headers, data, "ServerScorecard",
        widths={
            "A": 24, "B": 20, "C": 24, "D": 22, "E": 16, "F": 20,
            "G": 18, "H": 18, "I": 54, "J": 40, "K": 40, "L": 48,
        },
    )
    ws.freeze_panes = "D4"
    ws.sheet_view.zoomScale = 80
    for row in range(4, 4 + len(data)):
        for col in (1, 5, 6, 7, 8):
            fill = priority_fill(ws.cell(row=row, column=col).value)
            if fill:
                ws.cell(row=row, column=col).fill = fill
        for col in (4, 9, 10, 11, 12):
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 42


def write_rising_falling_sheet(wb: Workbook, rows: list[dict[str, Any]]) -> None:
    remove_sheet_if_present(wb, "Rising & Falling Stars")
    star_rows = [row for row in rows if row["prominent"] and row["momentum"] in {"Rising", "Falling"}]
    headers = [
        "Category",
        "Action",
        "Location",
        "Server",
        "Current Sample",
        "Performance",
        "Recent Momentum",
        "8-Week Direction",
        "History Used",
        "Positive Drivers",
        "Watch Drivers",
        "Recommended Next Step",
    ]
    data = [
        [
            f"{row['momentum']} Star",
            row["action"],
            row["location"],
            row["display_name"],
            f"{row['guest_count']:,.0f} guests / {row['active_days']} days",
            row["performance_level"],
            row["momentum"],
            row["long_term_direction"],
            row["history_used"],
            "; ".join(row["positive_drivers"]),
            "; ".join(row["negative_drivers"]),
            row["recommended_next_step"],
        ]
        for row in star_rows
    ]
    ws = write_table_sheet(
        wb, "Rising & Falling Stars", headers, data, "RisingFallingStarsV2",
        widths={
            "A": 15, "B": 24, "C": 20, "D": 24, "E": 22, "F": 20,
            "G": 18, "H": 18, "I": 54, "J": 40, "K": 40, "L": 48,
        },
    )
    ws.freeze_panes = "E4"
    ws.sheet_view.zoomScale = 80
    for row in range(4, 4 + len(data)):
        for col in (1, 2, 6, 7, 8):
            fill = priority_fill(ws.cell(row=row, column=col).value)
            if fill:
                ws.cell(row=row, column=col).fill = fill
        for col in (5, 9, 10, 11, 12):
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 42


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
    wb: Workbook,
    rows: list[dict[str, Any]],
    config: dict[str, Any],
    weekly_location_rows: list[dict[str, Any]],
    weekly_group_rows: list[dict[str, Any]],
    readiness: LatestWeekReadiness | None = None,
) -> None:
    remove_sheet_if_present(wb, "Store & Group Scorecards")
    ws = wb.create_sheet("Store & Group Scorecards")
    style_management_title(ws, "Store & Group Scorecards", 7)
    ws.freeze_panes = "A4"
    current_row = 4
    display_items: list[tuple[str, dict[str, Any] | None, str]] = []
    if readiness is None:
        display_items = [(item["entity"], item, "current") for item in rows]
    else:
        by_entity = {item["entity"]: item for item in rows}
        configured = set(readiness.configured_locations)
        for item in rows:
            if item["entity"] not in configured:
                state = "current" if readiness.ready else "preliminary"
                display_items.append((item["entity"], item, state))
        for location in readiness.configured_locations:
            item = by_entity.get(location)
            item_week_end = as_date(item.get("latest", {}).get("week_end")) if item else None
            if item is None or (
                item_week_end is not None and item_week_end != readiness.latest_week_end
            ):
                display_items.append((location, None, "missing"))
            else:
                state = "current" if readiness.ready else "preliminary"
                display_items.append((location, item, state))

    for entity, item, readiness_state in display_items:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        if readiness_state == "missing":
            title_text = f"{entity} | Missing | No current-week data; comparisons are paused."
            title_priority = "Review"
        elif readiness_state == "preliminary":
            missing_text = readiness.missing_text if readiness else "latest-week data"
            title_text = f"{entity} | Preliminary | Comparisons paused; missing {missing_text or 'latest-week data'}."
            title_priority = "Review"
        else:
            title_text = f"{entity} | {item['status']} | {item['recommended_focus']}"
            title_priority = item["priority"]
        title = ws.cell(row=current_row, column=1, value=title_text)
        title.fill = priority_fill(title_priority) or PatternFill("solid", fgColor="E7E6E6")
        title.font = Font(bold=True)
        title.alignment = Alignment(wrap_text=True)
        for col, header in enumerate(["Metric", "Current", "vs Prior", "vs Benchmark", "Benchmark", "Source", "Status"], start=1):
            cell = ws.cell(row=current_row + 1, column=col, value=header)
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        for offset, (field, label, _) in enumerate(MANAGEMENT_METRICS, start=2):
            row = current_row + offset
            ws.cell(row=row, column=1, value=label)
            if readiness_state == "missing":
                ws.cell(row=row, column=6, value="No current-week data")
                ws.cell(row=row, column=7, value="Missing")
                fill = priority_fill("Review")
                if fill:
                    ws.cell(row=row, column=7).fill = fill
                continue

            current_value, number_format = format_management_value(field, item["latest"][field])
            prior_change = item["prior_changes"][field]
            benchmark_change = item["benchmark_changes"][field]
            benchmark_value, benchmark_format = format_management_value(field, item["benchmark_values"][field])
            ws.cell(row=row, column=2, value=current_value).number_format = number_format
            if readiness_state == "preliminary":
                ws.cell(row=row, column=5, value=benchmark_value).number_format = benchmark_format
                ws.cell(row=row, column=6, value=f"Preliminary; {item['benchmark_sources'][field]}")
                status = "Preliminary"
            elif field in {"gross_sales", "guest_count"}:
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
            if readiness_state == "current":
                ws.cell(row=row, column=5, value=benchmark_value).number_format = benchmark_format
                ws.cell(row=row, column=6, value=item["benchmark_sources"][field])
                status = management_metric_status(field, item, config)
            ws.cell(row=row, column=7, value=status)
            fill = priority_fill(
                "Review" if status == "Preliminary" else
                "Medium" if status == "Watch" else
                "Recognize" if status == "Above" else "Monitor"
            )
            if fill:
                ws.cell(row=row, column=7).fill = fill
        current_row += 10
    for column, width in {"A": 24, "B": 16, "C": 16, "D": 18, "E": 16, "F": 20, "G": 14}.items():
        ws.column_dimensions[column].width = width
    ws.column_dimensions["H"].width = 3
    ws.sheet_view.zoomScale = 80
    add_management_scorecard_charts(wb, ws, weekly_location_rows, weekly_group_rows)


def latest_location_completeness(
    weekly_location_rows: list[dict[str, Any]],
    latest_week_end: date | None,
    config: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[str], list[str], bool]:
    latest_rows = [
        row for row in weekly_location_rows if row["week_end"] == latest_week_end
    ]
    latest_by_location = {row["location"]: row for row in latest_rows}
    configured_locations = list(config.get("locations", {}))
    if not configured_locations:
        configured_locations = sorted(latest_by_location)
    location_gaps = [
        location
        for location in configured_locations
        if location not in latest_by_location
        or latest_by_location[location].get("source_days", 0) < OPERATING_WEEK_DAYS
    ]
    complete = bool(configured_locations) and not location_gaps
    return latest_rows, configured_locations, location_gaps, complete


def write_management_data_quality_sheet(
    wb: Workbook,
    weekly_location_rows: list[dict[str, Any]],
    config: dict[str, Any],
    readiness: LatestWeekReadiness | None = None,
) -> None:
    remove_sheet_if_present(wb, "Data Quality")
    ws = wb.create_sheet("Data Quality")
    style_management_title(ws, "Data Quality", 6)
    if readiness is None:
        latest_week_end = max((row["week_end"] for row in weekly_location_rows), default=None)
        latest_rows, configured_locations, _location_gaps, latest_complete = latest_location_completeness(
            weekly_location_rows, latest_week_end, config
        )
    else:
        latest_week_end = readiness.latest_week_end
        latest_rows = list(readiness.latest_location_rows)
        configured_locations = list(readiness.configured_locations)
        latest_complete = readiness.ready
    latest_by_location = {row["location"]: row for row in latest_rows}
    ws.merge_cells("A3:F3")
    ws["A3"] = (
        f"Latest week ending {latest_week_end:%m/%d/%Y} is complete and suitable for management trends."
        if latest_complete and latest_week_end
        else "Latest week is incomplete. Management trends are preliminary and server actions are suppressed."
    )
    report_audit = config.get("_report_audit", {})
    if report_audit:
        ws["A3"] = (
            f"{ws['A3'].value} Input audit: "
            f"{report_audit.get('unique_business_days', 0)} unique business days used; "
            f"{report_audit.get('duplicate_files_ignored', 0)} semantic duplicate files ignored; "
            "no report conflicts."
        )
    ws["A3"].fill = PatternFill("solid", fgColor="D9EAD3" if latest_complete else "F4CCCC")
    ws["A3"].font = Font(bold=True)
    ws["A3"].alignment = Alignment(wrap_text=True)
    for col, header in enumerate(["Latest Week", "Location", "Active Days", "Source Days", "Status", "Management Use"], start=1):
        cell = ws.cell(row=5, column=col, value=header)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.font = Font(bold=True)
    for row_index, location in enumerate(configured_locations, start=6):
        row = latest_by_location.get(location)
        status = (
            "Missing" if row is None
            else "Complete" if row.get("source_days", 0) >= OPERATING_WEEK_DAYS
            else "Short Week"
        )
        values = [
            latest_week_end,
            location,
            row.get("active_days", 0) if row else 0,
            row.get("source_days", 0) if row else 0,
            status,
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
    _, global_full_week_ends = full_week_ends_by_location(weekly_location_rows)
    week_ends = sorted(global_full_week_ends)[-8:]
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

    group_lookup = {row["week_end"]: row for row in weekly_group_rows}
    group_rows = [group_lookup[week_end] for week_end in week_ends if week_end in group_lookup]
    ws.cell(row=1, column=6, value="Week End")
    ws.cell(row=1, column=7, value="Guest Count")
    for row_index, row in enumerate(group_rows, start=2):
        ws.cell(row=row_index, column=6, value=row["week_end"].strftime("%m/%d"))
        ws.cell(row=row_index, column=7, value=row["guest_count"])
    ws.sheet_state = "hidden"
    return len(week_ends), len(group_rows)


def add_management_scorecard_charts(
    wb: Workbook,
    ws,
    weekly_location_rows: list[dict[str, Any]],
    weekly_group_rows: list[dict[str, Any]],
) -> None:
    location_points, group_points = write_management_chart_data(
        wb, weekly_location_rows, weekly_group_rows
    )
    chart_ws = wb["_Dashboard Chart Data"]
    if location_points:
        locations = sorted({row["location"] for row in weekly_location_rows})
        chart = LineChart()
        chart.title = "Complete-Week Sales Trend by Store"
        chart.y_axis.title = "Gross Sales"
        chart.y_axis.numFmt = "$#,##0"
        chart.x_axis.tickLblPos = "low"
        chart.height = 7
        chart.width = 12.5
        data = Reference(
            chart_ws,
            min_col=2,
            max_col=1 + len(locations),
            min_row=1,
            max_row=1 + location_points,
        )
        cats = Reference(chart_ws, min_col=1, min_row=2, max_row=1 + location_points)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        category_formula = f"'_Dashboard Chart Data'!$A$2:$A${1 + location_points}"
        for series in chart.series:
            series.cat = AxDataSource(strRef=StrRef(f=category_formula))
        for series, location, color in zip(
            chart.series, locations, ("4472C4", "C0504D")
        ):
            series.tx = SeriesLabel(v=location.replace("RC ", ""))
            series.graphicalProperties.line.solidFill = color
            series.graphicalProperties.line.width = 28575
        chart.legend.position = "r"
        ws.add_chart(chart, "I4")
    if group_points:
        chart = LineChart()
        chart.title = "Complete-Week All-Stores Guest Trend"
        chart.y_axis.title = "Guests"
        chart.y_axis.numFmt = "#,##0"
        chart.x_axis.tickLblPos = "low"
        chart.height = 7
        chart.width = 12.5
        data = Reference(chart_ws, min_col=7, min_row=1, max_row=1 + group_points)
        cats = Reference(chart_ws, min_col=6, min_row=2, max_row=1 + group_points)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        category_formula = f"'_Dashboard Chart Data'!$F$2:$F${1 + group_points}"
        for series in chart.series:
            series.cat = AxDataSource(strRef=StrRef(f=category_formula))
        if chart.series:
            chart.series[0].graphicalProperties.line.solidFill = "7A1E1E"
            chart.series[0].graphicalProperties.line.width = 28575
        chart.legend = None
        ws.add_chart(chart, "I20")


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


def latest_week_report_coverage(
    records: list[MetricRecord], latest_week_end: date | None
) -> tuple[list[date], set[date], list[date]]:
    if latest_week_end is None:
        return [], set(), []
    week_start = latest_week_end - timedelta(days=OPERATING_WEEK_DAYS - 1)
    expected = [week_start + timedelta(days=offset) for offset in range(OPERATING_WEEK_DAYS)]
    received = {
        record.report_date
        for record in records
        if week_start <= record.report_date <= latest_week_end
        and is_operating_day(record.report_date)
    }
    missing = [report_date for report_date in expected if report_date not in received]
    return expected, received, missing


def latest_week_readiness(
    records: list[MetricRecord],
    weekly_location_rows: list[dict[str, Any]],
    config: dict[str, Any],
) -> LatestWeekReadiness:
    latest_week_end = max((row["week_end"] for row in weekly_location_rows), default=None)
    latest_rows, configured_locations, location_gaps, locations_complete = (
        latest_location_completeness(weekly_location_rows, latest_week_end, config)
    )
    expected_dates, received_dates, missing_dates = latest_week_report_coverage(
        records, latest_week_end
    )
    return LatestWeekReadiness(
        latest_week_end=latest_week_end,
        latest_location_rows=tuple(latest_rows),
        configured_locations=tuple(configured_locations),
        location_gaps=tuple(location_gaps),
        expected_dates=tuple(expected_dates),
        received_dates=frozenset(received_dates),
        missing_dates=tuple(missing_dates),
        ready=locations_complete and not missing_dates,
    )


def deduplicated_dashboard_actions(
    action_rows: list[dict[str, Any]],
    priorities: set[str],
    limit: int | None = None,
) -> list[dict[str, Any]]:
    priority_order = {"High": 0, "Medium": 1, "Review": 2, "Recognize": 3, "Share": 4}
    candidates = [
        row
        for row in action_rows
        if str(row.get("Priority") or "") in priorities
        and str(row.get("Status") or "Open").casefold() not in {"complete", "dismissed"}
    ]
    candidates.sort(
        key=lambda row: (
            priority_order.get(str(row.get("Priority") or ""), 9),
            str(row.get("Location") or ""),
            str(row.get("Person / Area") or ""),
        )
    )
    selected: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in candidates:
        identity = str(row.get("Entity Key") or "").strip().casefold()
        if not identity:
            identity = "|".join(
                str(row.get(field) or "").strip().casefold()
                for field in ("Location", "Person / Area", "Action", "Signal")
            )
        if identity in seen:
            continue
        seen.add(identity)
        selected.append(row)
        if limit is not None and len(selected) >= limit:
            break
    return selected


def dashboard_management_move(item: dict[str, Any]) -> str:
    signal = str(item.get("Signal") or "").strip().rstrip(".:")
    evidence = str(item.get("Why It Matters") or "").split(" | ")[0].strip().rstrip(".")
    next_step = str(item.get("Recommended Next Step") or "").strip().rstrip(".")
    context = ". ".join(part for part in (signal, evidence) if part)
    if next_step:
        return f"{context}. Next: {next_step}." if context else f"Next: {next_step}."
    return f"{context}." if context else "Review and assign the next management step."


def write_dashboard_card(
    ws,
    start_col: int,
    title: str,
    value: str | int,
    note: str,
    accent_color: str,
) -> None:
    end_col = start_col + 3
    ws.merge_cells(start_row=6, start_column=start_col, end_row=6, end_column=end_col)
    ws.merge_cells(start_row=7, start_column=start_col, end_row=8, end_column=end_col)
    ws.merge_cells(start_row=9, start_column=start_col, end_row=9, end_column=end_col)
    title_cell = ws.cell(row=6, column=start_col, value=title)
    title_cell.fill = PatternFill("solid", fgColor=accent_color)
    title_cell.font = Font(bold=True, color="FFFFFF", size=10)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    value_cell = ws.cell(row=7, column=start_col, value=value)
    value_cell.fill = PatternFill("solid", fgColor="F7F7F7")
    value_cell.font = Font(bold=True, color="7A1E1E", size=20)
    value_cell.alignment = Alignment(horizontal="center", vertical="center")
    note_cell = ws.cell(row=9, column=start_col, value=note)
    note_cell.fill = PatternFill("solid", fgColor="F7F7F7")
    note_cell.font = Font(color="595959", size=9)
    note_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_management_dashboard_sheet(
    wb: Workbook,
    records: list[MetricRecord],
    weekly_location_rows: list[dict[str, Any]],
    weekly_group_rows: list[dict[str, Any]],
    server_rows: list[dict[str, Any]],
    store_rows: list[dict[str, Any]],
    group_rows: list[dict[str, Any]],
    action_rows: list[dict[str, Any]],
    config: dict[str, Any] | None = None,
    readiness: LatestWeekReadiness | None = None,
) -> None:
    config = config or DEFAULT_CONFIG
    remove_sheet_if_present(wb, "Dashboard")
    ws = wb.create_sheet("Dashboard")
    style_management_title(ws, "Red Onion Weekly Brief", 12)
    widths = {
        "A": 11, "B": 11, "C": 12, "D": 12, "E": 12, "F": 12,
        "G": 12, "H": 12, "I": 12, "J": 16, "K": 14, "L": 14,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A5"
    readiness = readiness or latest_week_readiness(records, weekly_location_rows, config)
    latest_week_end = readiness.latest_week_end
    latest_location_rows = list(readiness.latest_location_rows)
    configured_locations = list(readiness.configured_locations)
    expected_dates = readiness.expected_dates
    received_dates = readiness.received_dates
    latest_complete = readiness.ready
    missing_text = readiness.missing_text
    current_location_names = {row["location"] for row in latest_location_rows}
    current_store_rows = {
        item["entity"]: item
        for item in store_rows
        if item["entity"] in current_location_names
    }
    _, global_full = full_week_ends_by_location(weekly_location_rows)

    ws.merge_cells("A3:L3")
    ws["A3"] = (
        f"Week ending {latest_week_end:%m/%d/%Y} | {len(received_dates)} of {len(expected_dates)} daily reports | "
        f"{len(global_full)} complete weeks available"
        if latest_week_end
        else "No management data available"
    )
    ws["A3"].font = Font(bold=True, color="595959")
    ws["A3"].alignment = Alignment(vertical="center")
    ws.merge_cells("A4:L4")
    ws["A4"] = (
        "READY FOR MANAGEMENT REVIEW - comparisons, actions, and recognition are active"
        if latest_complete
        else f"PRELIMINARY - missing {missing_text or 'daily reports'}; comparisons, actions, and recognition are paused"
    )
    ws["A4"].fill = PatternFill("solid", fgColor="D9EAD3" if latest_complete else "F4CCCC")
    ws["A4"].font = Font(bold=True)
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 28

    group = group_rows[0] if group_rows else None
    all_action_items = deduplicated_dashboard_actions(
        action_rows, {"High", "Medium", "Review"}
    )
    selected_actions = all_action_items[:3]
    reports_value = f"{len(received_dates)} of {len(expected_dates)}" if expected_dates else "0 of 0"
    reports_note = (
        "All Tuesday-Sunday reports received"
        if latest_complete
        else f"Missing {missing_text or 'daily reports'}"
    )
    write_dashboard_card(
        ws, 1, "Reports Received", reports_value, reports_note,
        "548235" if latest_complete else "C00000",
    )

    traffic_delta = None
    traffic_source = "benchmark"
    if group:
        traffic_delta = safe_pct_delta(
            group["latest"]["guest_count"], group["benchmark_values"].get("guest_count")
        )
        traffic_source = str(group.get("benchmark_sources", {}).get("guest_count") or "benchmark")
    if not latest_complete:
        traffic_value, traffic_note, traffic_color = "PAUSED", f"Missing {missing_text or 'daily reports'}", "7F7F7F"
    elif traffic_delta is None:
        traffic_value, traffic_note, traffic_color = "NO BASELINE", "Guest traffic benchmark is not available", "7F7F7F"
    else:
        traffic_value = f"{traffic_delta:+.1%}"
        traffic_note = f"Guest traffic vs {traffic_source.lower()}"
        traffic_color = "548235" if traffic_delta >= 0 else "C00000" if traffic_delta <= -0.05 else "BF9000"
    write_dashboard_card(ws, 5, "Traffic vs Benchmark", traffic_value, traffic_note, traffic_color)

    if latest_complete:
        action_value: str | int = len(all_action_items)
        action_note = "Open coaching and review items"
        action_color = "C00000" if all_action_items else "548235"
    else:
        action_value, action_note, action_color = "PAUSED", f"Missing {missing_text or 'daily reports'}", "7F7F7F"
    write_dashboard_card(
        ws, 9, "High-Priority Actions", action_value, action_note, action_color
    )

    ws.merge_cells("A11:L11")
    ws["A11"] = "TOP THREE ACTIONS"
    ws["A11"].fill = PatternFill("solid", fgColor="E7E6E6")
    ws["A11"].font = Font(bold=True, color="7A1E1E")
    action_headers = [
        (1, 2, "Priority"), (3, 5, "Person / Location"), (6, 10, "Management Move"),
        (11, 11, "Owner"), (12, 12, "Due"),
    ]
    for start_col, end_col, header in action_headers:
        if end_col > start_col:
            ws.merge_cells(start_row=12, start_column=start_col, end_row=12, end_column=end_col)
        cell = ws.cell(row=12, column=start_col, value=header)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if latest_complete:
        for row_index in range(13, 16):
            item = selected_actions[row_index - 13] if row_index - 13 < len(selected_actions) else None
            ws.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=2)
            ws.merge_cells(start_row=row_index, start_column=3, end_row=row_index, end_column=5)
            ws.merge_cells(start_row=row_index, start_column=6, end_row=row_index, end_column=10)
            if item:
                priority = str(item.get("Priority") or "Review")
                person = str(item.get("Person / Area") or "").strip()
                location = str(item.get("Location") or "").strip()
                person_location = person if person.casefold() == location.casefold() else " | ".join(
                    part for part in (person, location) if part
                )
                owner = str(item.get("Owner") or "Unassigned")
                due_date = as_date(item.get("Due Date"))
                ws.cell(row=row_index, column=1, value=priority)
                ws.cell(row=row_index, column=3, value=person_location)
                ws.cell(row=row_index, column=6, value=dashboard_management_move(item))
                ws.cell(row=row_index, column=11, value=owner)
                if due_date:
                    ws.cell(row=row_index, column=12, value=due_date).number_format = "m/d/yyyy"
                else:
                    ws.cell(row=row_index, column=12, value="Not set")
                fill = priority_fill(priority)
                if fill:
                    ws.cell(row=row_index, column=1).fill = fill
                if owner == "Unassigned":
                    ws.cell(row=row_index, column=11).fill = PatternFill("solid", fgColor="FFF2CC")
                if not due_date:
                    ws.cell(row=row_index, column=12).fill = PatternFill("solid", fgColor="FFF2CC")
            else:
                ws.cell(row=row_index, column=1, value="—")
                ws.cell(row=row_index, column=3, value="No additional priority item")
            for col in (1, 3, 6, 11, 12):
                ws.cell(row=row_index, column=col).alignment = Alignment(
                    wrap_text=True, vertical="center"
                )
            ws.row_dimensions[row_index].height = 42
    else:
        ws.merge_cells("A13:L15")
        ws["A13"] = (
            f"Actions are paused until the latest week is complete. Missing: {missing_text or 'daily reports'}."
        )
        ws["A13"].fill = PatternFill("solid", fgColor="FCE8E6")
        ws["A13"].font = Font(bold=True, color="9C0006")
        ws["A13"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("A17:L17")
    ws["A17"] = "STORE SNAPSHOT"
    ws["A17"].fill = PatternFill("solid", fgColor="E7E6E6")
    ws["A17"].font = Font(bold=True, color="7A1E1E")
    store_headers = [
        (1, 2, "Location"), (3, 4, "Status"), (5, 5, "Sales"), (6, 6, "Guests"),
        (7, 7, "Check Avg"), (8, 9, "Service Pace"), (10, 12, "Management Focus"),
    ]
    for start_col, end_col, header in store_headers:
        if end_col > start_col:
            ws.merge_cells(start_row=18, start_column=start_col, end_row=18, end_column=end_col)
        cell = ws.cell(row=18, column=start_col, value=header)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_index in range(19, 21):
        location_index = row_index - 19
        location = (
            configured_locations[location_index]
            if location_index < len(configured_locations)
            else None
        )
        item = current_store_rows.get(location) if location else None
        ws.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=2)
        ws.merge_cells(start_row=row_index, start_column=3, end_row=row_index, end_column=4)
        ws.merge_cells(start_row=row_index, start_column=8, end_row=row_index, end_column=9)
        ws.merge_cells(start_row=row_index, start_column=10, end_row=row_index, end_column=12)
        if item:
            latest = item["latest"]
            status = item["status"] if latest_complete else "Preliminary"
            focus = item["recommended_focus"] if latest_complete else (
                f"Comparisons paused; missing {missing_text or 'daily reports'}."
            )
            ws.cell(row=row_index, column=1, value=item["entity"])
            ws.cell(row=row_index, column=3, value=status)
            ws.cell(row=row_index, column=5, value=latest["gross_sales"]).number_format = "$#,##0"
            ws.cell(row=row_index, column=6, value=latest["guest_count"]).number_format = "#,##0"
            ws.cell(row=row_index, column=7, value=latest["check_average"]).number_format = "$0.00"
            ws.cell(row=row_index, column=8, value=latest["average_ticket_time_seconds"] / 60).number_format = '0.0 "min"'
            ws.cell(row=row_index, column=10, value=focus)
            fill = priority_fill(item["priority"] if latest_complete else "Review")
            if fill:
                ws.cell(row=row_index, column=3).fill = fill
        else:
            ws.cell(row=row_index, column=1, value=location or "No store data")
            ws.cell(row=row_index, column=3, value="Missing")
            ws.cell(row=row_index, column=10, value="No current-week data; comparisons are paused.")
            fill = priority_fill("Review")
            if fill:
                ws.cell(row=row_index, column=3).fill = fill
        for col in (1, 3, 5, 6, 7, 8, 10):
            ws.cell(row=row_index, column=col).alignment = Alignment(
                wrap_text=True, vertical="center"
            )
        ws.row_dimensions[row_index].height = 36

    ws.merge_cells("A22:L22")
    ws["A22"] = "RECOGNITION / REPLICATE"
    ws["A22"].fill = PatternFill("solid", fgColor="E7E6E6")
    ws["A22"].font = Font(bold=True, color="7A1E1E")
    ws.merge_cells("A23:L23")
    recognition = deduplicated_dashboard_actions(
        action_rows, {"Recognize", "Share"}, limit=1
    )
    if not latest_complete:
        recognition_text = (
            f"Recognition is paused until the latest week is complete. Missing: {missing_text or 'daily reports'}."
        )
        recognition_fill = "FCE8E6"
    elif recognition:
        item = recognition[0]
        person = str(item.get("Person / Area") or item.get("Location") or "Team")
        location = str(item.get("Location") or "")
        identity = person if person.casefold() == location.casefold() else " | ".join(
            part for part in (person, location) if part
        )
        recognition_text = f"{identity} — {dashboard_management_move(item)}"
        recognition_fill = "D9EAD3"
    else:
        recognition_text = "No recognition item met the current complete-week thresholds."
        recognition_fill = "F3F4F6"
    ws["A23"] = recognition_text
    ws["A23"].fill = PatternFill("solid", fgColor=recognition_fill)
    ws["A23"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[23].height = 34

    ws.merge_cells("A24:L24")
    ws["A24"] = (
        f"Data quality: {len(received_dates)} of {len(expected_dates)} reports received; "
        f"{len(global_full)} complete all-store weeks available; incomplete weeks are excluded from comparisons."
        if latest_complete
        else f"Data quality: {len(received_dates)} of {len(expected_dates)} reports received; missing {missing_text or 'daily reports'}."
    )
    ws["A24"].fill = PatternFill("solid", fgColor="D9EAF7" if latest_complete else "F4CCCC")
    ws["A24"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[24].height = 30

    ws.print_area = "A1:L24"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_options.horizontalCentered = True
    ws.sheet_view.zoomScale = 90


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
    report_audit = config.get("_report_audit", {})
    note_rows = [
        ("Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Source Folder", str(source_dir)),
        ("Operating Week", f"{OPERATING_WEEK_LABEL}; Mondays are closed."),
        ("Raw Reports Read", len({record.source_file for record in records})),
        ("Unique Business Days Used", report_audit.get("unique_business_days", len({record.report_date for record in records}))),
        ("Semantic Duplicates Ignored", report_audit.get("duplicate_files_ignored", 0)),
        ("Canonical History Folder", report_audit.get("canonical_archive", "Not recorded")),
        ("Report Conflicts", "None; conflicting same-date reports stop the run before output or archiving."),
        ("Date Coverage", format_date_range(min(r.report_date for r in records), max(r.report_date for r in records))),
        ("Public Snapshot Dates", format_date_range(public_start, public_end)),
        ("Recent Momentum", f"Latest complete Tuesday-Sunday week versus up to {config.get('dashboard_baseline_full_weeks', 4)} prior complete weeks. Partial latest weeks and low current samples are Not Scored."),
        ("Recent Momentum Confidence", f"Latest guests >= {config.get('dashboard_min_guest_count_for_trends', 25)} AND active days >= {config.get('dashboard_min_active_days_for_trends', 3)}, plus at least {config.get('dashboard_min_prior_full_weeks', 2)} prior full weeks and {config.get('dashboard_min_prior_guest_count', 50)} prior guests."),
        ("8-Week Direction", f"Compares the most recent {config.get('dashboard_long_term_block_weeks', 4)} complete weeks with the preceding {config.get('dashboard_long_term_block_weeks', 4)}. Full uses eight server weeks with at least {config.get('dashboard_long_term_full_min_recent_guests', 100)} recent / {config.get('dashboard_long_term_full_min_earlier_guests', 100)} earlier guests; Developing requires at least {config.get('dashboard_long_term_developing_min_total_weeks', 6)} usable weeks, {config.get('dashboard_long_term_developing_min_recent_weeks', 3)} recent / {config.get('dashboard_long_term_developing_min_earlier_weeks', 2)} earlier weeks, and {config.get('dashboard_long_term_developing_min_recent_guests', 75)} recent / {config.get('dashboard_long_term_developing_min_earlier_guests', 50)} earlier guests."),
        ("Targets", "Management Setup targets take precedence; blank targets use the rolling baseline. Edits apply on the next run."),
        ("Trend Scoring", "Both trend horizons use four metric families scored from -2 to +2 using materiality bands; average rank movement contributes only -1, 0, or +1."),
        ("Performance Level", "Latest metrics are assessed separately as Above Benchmark, On Track, or Below Benchmark."),
        ("Technical Trend Detail", "Server Week-over-Week Detail shows adjacent-week changes for audit use. Management coaching uses Recent Momentum and 8-Week Direction instead."),
        ("Action Tracking", "Owner, due date, status, and manager notes carry forward between weekly runs. Cleared signals move to Action History."),
        ("Metric Rule", "Check average and wine percent are recalculated from rolled-up sales, guests, and wine sales."),
        ("Metric Rule", "Rate of sale and ticket time are guest-weighted averages."),
    ]
    for row, (label, value) in enumerate(note_rows, start=4):
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=2, value=value).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 45 if label == "8-Week Direction" else 30


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
        readiness = latest_week_readiness(records, weekly_location_rows, config)
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
            server_rows, store_rows, group_rows, weekly_location_rows, readiness
        )
        current_actions, action_history = merge_management_actions(signals, state, readiness)

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
        write_store_group_scorecards_sheet(
            wb,
            [*group_rows, *store_rows],
            config,
            weekly_location_rows,
            weekly_group_rows,
            readiness,
        )
        write_rising_falling_sheet(wb, server_rows)
        write_action_tracking_sheet(wb, "Action History", action_history, editable=False)
        write_management_data_quality_sheet(wb, weekly_location_rows, config, readiness)
        write_management_run_notes(wb, records, source_dir, public_start, public_end, config)
        write_management_dashboard_sheet(
            wb, records, weekly_location_rows, weekly_group_rows, server_rows,
            store_rows, group_rows, current_actions, config, readiness,
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
    config_path = Path(args.config).resolve()
    config = load_config(config_path)

    migration_sources = [
        Path(path).resolve() for path in (getattr(args, "migrate_history_from", None) or [])
    ]
    migration_only = bool(getattr(args, "migrate_history_only", False))
    if migration_only and not migration_sources:
        raise ValueError("--migrate-history-only requires at least one --migrate-history-from folder.")

    migration_plan = (
        build_history_migration_plan(migration_sources, archive_dir, config)
        if migration_sources
        else None
    )
    if migration_only:
        assert migration_plan is not None
        return list(apply_history_migration_plan(migration_plan))

    active_paths = daily_report_paths(input_dir)
    if not active_paths:
        raise FileNotFoundError(
            f"No active daily reports ({DAILY_REPORT_FORMAT_LABEL}) found in {input_dir}. "
            "Drop current Toast reports into 01 Daily Reports - Drop Here and rerun."
        )

    raw_active_records_by_path = read_reports_by_path(active_paths, config)
    active_resolution = resolve_report_duplicates(raw_active_records_by_path)
    active_records_by_path = active_resolution.records_by_path
    _, active_week_end = active_week_for_paths(active_records_by_path)
    active_records = flatten_report_records(active_records_by_path)

    if migration_plan is not None:
        historical_records_by_path = migration_plan.effective_records_by_path
        historical_duplicate_paths = migration_plan.duplicate_paths
    else:
        archived_paths = archived_daily_report_paths(archive_dir)
        archived_records_by_path = (
            read_reports_by_path(archived_paths, config) if archived_paths else {}
        )
        historical_resolution = resolve_report_duplicates(archived_records_by_path)
        historical_records_by_path = historical_resolution.records_by_path
        historical_duplicate_paths = historical_resolution.duplicate_paths

    combined_records_by_path = dict(historical_records_by_path)
    combined_records_by_path.update(active_records_by_path)
    combined_resolution = resolve_report_duplicates(combined_records_by_path)
    records = flatten_report_records(combined_resolution.records_by_path)

    duplicate_paths = sorted(
        {
            *active_resolution.duplicate_paths,
            *historical_duplicate_paths,
            *combined_resolution.duplicate_paths,
        },
        key=lambda path: str(path).casefold(),
    )
    config["_report_audit"] = {
        "canonical_archive": str(canonical_daily_archive_dir(archive_dir)),
        "unique_business_days": len(combined_resolution.business_dates),
        "duplicate_files_ignored": len(duplicate_paths),
        "duplicate_file_names": [path.name for path in duplicate_paths],
        "conflicts": 0,
    }

    if migration_plan is not None:
        apply_history_migration_plan(migration_plan)

    output_dir.mkdir(parents=True, exist_ok=True)

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
        help=f"Folder containing active raw daily {DAILY_REPORT_FORMAT_LABEL} reports.",
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
    parser.add_argument(
        "--migrate-history-from",
        action="append",
        default=[],
        metavar="FOLDER",
        help=(
            "Copy validated daily reports from a legacy folder into the canonical processed "
            "archive. Repeat this option for additional source folders."
        ),
    )
    parser.add_argument(
        "--migrate-history-only",
        action="store_true",
        help="Perform the copy-only history migration without generating weekly workbooks.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    try:
        generated = run(args)
    except Exception as exc:
        raise SystemExit(f"ERROR: {exc}") from exc
    if args.migrate_history_only:
        if generated:
            print("History copied to the canonical archive:")
            for path in generated:
                print(f"  {path}")
        else:
            print("History migration complete: the canonical archive is already up to date.")
    else:
        print("Generated:")
        for path in generated:
            print(f"  {path}")


if __name__ == "__main__":
    main()
