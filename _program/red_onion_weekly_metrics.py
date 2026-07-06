from __future__ import annotations

import argparse
import filecmp
import json
import re
import shutil
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


DEFAULT_CONFIG: dict[str, Any] = {
    "locations": {
        "RC Richmond": {"short_code": "RVA"},
        "RC Virginia Beach": {"short_code": "VB"},
    },
    "public_min_guest_count": 1,
    "master_min_guest_count_for_rankings": 1,
    "dashboard_exclude_name_contains": ["Banquet", "Server"],
    "public_name_aliases": {
        "Bar 1 Bar 1": "Bar",
        "Bar Server": "Bar",
        "BarPatio Bartender Patio": "Patio",
    },
    "public_exclude_name_contains": ["Banquet", "Takeout", "Server Server"],
}

METRICS = [
    ("gross_sales", "Gross Sales"),
    ("guest_count", "Guest Count"),
    ("check_average", "Check Average"),
    ("wine_sales", "Wine Sales"),
    ("rate_of_sale_by_guest_count", "Rate of Sale by Guest Count"),
    ("average_ticket_time", "Average Ticket Time"),
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


def public_excluded(row: dict[str, Any], config: dict[str, Any]) -> bool:
    if row["guest_count"] < float(config.get("public_min_guest_count", 1)):
        return True
    if is_numbered_server_placeholder(row.get("raw_user_name")) or is_numbered_server_placeholder(
        row.get("display_name")
    ):
        return True
    haystack = f"{row.get('raw_user_name', '')} {row.get('display_name', '')}".casefold()
    for pattern in config.get("public_exclude_name_contains", []):
        if str(pattern).casefold() in haystack:
            return True
    return False


def dashboard_excluded(row: dict[str, Any], config: dict[str, Any]) -> bool:
    haystack = str(row.get("display_name") or row.get("raw_user_name", "")).casefold()
    for pattern in config.get("dashboard_exclude_name_contains", []):
        if str(pattern).casefold() in haystack:
            return True
    return False


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
        ws.cell(row=row_index, column=7, value=f"=F{row_index}/C{row_index}")

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
) -> None:
    ws = wb.create_sheet(title)
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


def server_week_trend_rows(ranked_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in ranked_rows:
        groups[(row["location"], row["raw_user_name"], row["display_name"])].append(row)

    trend_rows: list[dict[str, Any]] = []
    for (_, _, _), rows in groups.items():
        rows.sort(key=lambda row: row["week_end"])
        previous: dict[str, Any] | None = None
        for row in rows:
            check_change = (
                row["check_average"] - previous["check_average"] if previous else None
            )
            wine_change = row["wine_pct"] - previous["wine_pct"] if previous else None
            rate_change = (
                row["rate_of_sale_by_guest_count"] - previous["rate_of_sale_by_guest_count"]
                if previous
                else None
            )
            ticket_change_minutes = (
                (row["average_ticket_time_seconds"] - previous["average_ticket_time_seconds"]) / 60
                if previous
                else None
            )
            trend_rows.append(
                {
                    **row,
                    "check_average_change": check_change,
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

    trend_rows.sort(
        key=lambda row: (
            row["week_end"],
            row["location"],
            row.get("check_average_rank") or 9999,
            row["display_name"],
        )
    )
    return trend_rows


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


def apply_dashboard_number_formats(ws, row_start: int, row_end: int) -> None:
    for row in range(row_start, row_end + 1):
        for col in range(1, ws.max_column + 1):
            header = str(ws.cell(row=row_start - 1, column=col).value or "")
            cell = ws.cell(row=row, column=col)
            if "Date" in header or "Week End" in header:
                cell.number_format = "m/d/yyyy"
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


def write_dashboard_sheet(
    wb: Workbook,
    records: list[MetricRecord],
    weekly_location_rows: list[dict[str, Any]],
    ranked_rows: list[dict[str, Any]],
    config: dict[str, Any],
    source_dir: Path,
    public_start: date,
    public_end: date,
) -> None:
    ws = wb.create_sheet("Dashboard", 0)
    style_title(ws, "Red Onion Weekly Performance Dashboard", 12)
    ws.freeze_panes = "A4"
    for col, width in {
        "A": 18,
        "B": 18,
        "C": 16,
        "D": 16,
        "E": 16,
        "F": 14,
        "G": 14,
        "H": 16,
        "I": 16,
        "J": 16,
        "K": 16,
        "L": 18,
    }.items():
        ws.column_dimensions[col].width = width

    dates = sorted({record.report_date for record in records})
    latest_week_end = max(row["week_end"] for row in weekly_location_rows)
    latest_location_rows = [
        row for row in weekly_location_rows if row["week_end"] == latest_week_end
    ]
    prior_week_end = max(
        (row["week_end"] for row in weekly_location_rows if row["week_end"] < latest_week_end),
        default=None,
    )
    prior_by_location = {
        row["location"]: row
        for row in weekly_location_rows
        if prior_week_end and row["week_end"] == prior_week_end
    }

    summary_rows = [
        ("Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Source Folder", str(source_dir)),
        ("Raw Reports Read", len({record.source_file for record in records})),
        ("Date Coverage", format_date_range(min(dates), max(dates))),
        ("Public Snapshot Dates", format_date_range(public_start, public_end)),
        ("Weeks Tracked", len({row["week_end"] for row in weekly_location_rows})),
    ]
    style_section_header(ws, 3, 1, 2, "Run Summary")
    for row_index, (label, value) in enumerate(summary_rows, start=4):
        ws.cell(row=row_index, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row_index, column=2, value=value)
        ws.cell(row=row_index, column=2).alignment = Alignment(wrap_text=True)

    start_row = 3
    start_col = 4
    headers = [
        "Week End",
        "Location",
        "Gross Sales",
        "Guest Count",
        "Check Average",
        "Check Average Change",
        "Wine %",
        "Wine % Change",
        "Rate of Sale by Guest Count",
        "Rate of Sale Change",
        "Ticket Time",
        "Ticket Time Change (Min)",
    ]
    for offset, header in enumerate(headers):
        cell = ws.cell(row=start_row, column=start_col + offset, value=header)
        cell.fill = PatternFill("solid", fgColor="E1E1E1")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_offset, row in enumerate(
        sorted(latest_location_rows, key=lambda item: item["location"]), start=1
    ):
        prior = prior_by_location.get(row["location"])
        values = [
            row["week_end"],
            row["location"],
            row["gross_sales"],
            row["guest_count"],
            row["check_average"],
            row["check_average"] - prior["check_average"] if prior else None,
            row["wine_pct"],
            row["wine_pct"] - prior["wine_pct"] if prior else None,
            row["rate_of_sale_by_guest_count"],
            row["rate_of_sale_by_guest_count"] - prior["rate_of_sale_by_guest_count"]
            if prior
            else None,
            duration_fraction(row["average_ticket_time_seconds"]),
            (row["average_ticket_time_seconds"] - prior["average_ticket_time_seconds"]) / 60
            if prior
            else None,
        ]
        for col_offset, value in enumerate(values):
            ws.cell(row=start_row + row_offset, column=start_col + col_offset, value=value)
    apply_dashboard_number_formats(ws, start_row + 1, start_row + len(latest_location_rows))

    latest_ranked = [
        row
        for row in ranked_rows
        if row["week_end"] == latest_week_end and not dashboard_excluded(row, config)
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

    table_row = 12
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
        chart.height = 7
        chart.width = 13
        ws.add_chart(chart, "H12")


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


def write_master_workbook(
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
    trend_rows = trend_summary_rows(weekly_server_rows)

    wb = Workbook()
    wb.remove(wb.active)
    write_dashboard_sheet(
        wb,
        records,
        weekly_location_rows,
        ranked_rows,
        config,
        source_dir,
        public_start,
        public_end,
    )

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
        "Check Average Change",
        "Wine %",
        "Wine % Rank",
        "Wine % Change",
        "Rate of Sale by Guest Count",
        "Rate Rank",
        "Rate Change",
        "Average Ticket Time",
        "Ticket Time Rank",
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
            row["check_average_change"],
            row["wine_pct"],
            row["wine_pct_rank"],
            row["wine_pct_change"],
            row["rate_of_sale_by_guest_count"],
            row["rate_rank"],
            row["rate_change"],
            duration_fraction(row["average_ticket_time_seconds"]),
            row["ticket_time_rank"],
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
            "I": 20,
            "L": 16,
            "M": 22,
            "O": 14,
            "P": 18,
            "R": 20,
            "T": 16,
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

    notes = wb.create_sheet("Run Notes", 1)
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
        ("Dashboard", "Dashboard summarizes latest weekly location results and current server leaders."),
        ("Trend Tabs", "Weekly Server Rankings ranks each metric by week/location; Server Week Trends adds week-over-week changes."),
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
