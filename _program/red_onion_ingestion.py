"""Small, side-effect-free contracts for daily report dates and columns."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
import re
from typing import Iterable, Sequence


METRIC_HEADERS = (
    "Gross Sales",
    "Guest Count",
    "Check Average",
    "Wine Sales",
    "Rate of Sale by Guest Count",
    "Average Ticket Time",
)


def parse_single_business_date(values: Iterable[object], filename: str) -> date:
    """Require explicit, consistent single-day metadata; filenames are labels."""
    parsed: set[date] = set()
    day_pattern = r"\d{1,2}/\d{1,2}/(?:\d{4}|\d{2})"
    for value in values:
        if not isinstance(value, str):
            continue
        for field in re.findall(r"Date\(s\):[ \t]*([^\r\n]*)", value, re.IGNORECASE):
            match = re.fullmatch(
                rf"\s*({day_pattern})(?:\s*[-\u2013\u2014]\s*({day_pattern}))?\s*",
                field,
            )
            if not match:
                raise ValueError(
                    f"{filename}: Date(s) must contain one explicit business date "
                    f"(MM/DD/YYYY), not {field!r}. Re-export this daily report."
                )
            endpoints = []
            for text in filter(None, match.groups()):
                for fmt in ("%m/%d/%Y", "%m/%d/%y"):
                    try:
                        endpoints.append(datetime.strptime(text, fmt).date())
                        break
                    except ValueError:
                        continue
                else:
                    raise ValueError(f"{filename}: invalid Date(s) value {text!r}.")
            if len(set(endpoints)) != 1:
                raise ValueError(
                    f"{filename}: Date(s) spans multiple business dates ({field}). "
                    "Use a separate single-day export for each business date."
                )
            parsed.update(endpoints)
    if len(parsed) != 1:
        detail = "conflicting Date(s) fields" if parsed else "no explicit Date(s) field"
        raise ValueError(
            f"{filename}: {detail}. Re-export with a single business date; "
            "a filename does not establish the business date."
        )
    return next(iter(parsed))


@dataclass(frozen=True)
class ReportColumns:
    location: int
    person: int
    blocks: tuple[dict[str, int], ...]


def validated_report_columns(headers: Sequence[str]) -> ReportColumns:
    """Map repeated complete metric blocks by labels, never by numeric offsets."""
    if headers.count("Store") != 1:
        raise ValueError("The report must have exactly one Store column.")
    location = headers.index("Store")
    person_columns = [
        index for index, label in enumerate(headers)
        if label in {"User", "User Name", "Server"}
    ]
    if not person_columns and location + 1 < len(headers) and not headers[location + 1]:
        # Older Red Onion exports leave the identity heading blank next to Store.
        person_columns = [location + 1]
    if len(person_columns) != 1:
        raise ValueError("The report must have one unambiguous User column.")
    occurrences = {
        label: [index for index, header in enumerate(headers) if header == label]
        for label in METRIC_HEADERS
    }
    block_count = len(occurrences["Gross Sales"])
    if not block_count or any(len(columns) != block_count for columns in occurrences.values()):
        raise ValueError(
            "Incomplete or ambiguous metric headers. Every report block must contain "
            + ", ".join(METRIC_HEADERS) + "."
        )
    blocks = [
        {label: columns[index] for label, columns in occurrences.items()}
        for index in range(block_count)
    ]
    if any(max(left.values()) >= min(right.values()) for left, right in zip(blocks, blocks[1:])):
        raise ValueError("Metric headers overlap ambiguously between report blocks.")
    check_columns = [index for index, label in enumerate(headers) if label == "Check Count"]
    if check_columns and len(check_columns) != block_count:
        raise ValueError("Check Count must appear once in every metric block, or be absent.")
    for index, column in enumerate(check_columns):
        blocks[index]["Check Count"] = column
    if any(max(left.values()) >= min(right.values()) for left, right in zip(blocks, blocks[1:])):
        raise ValueError("Check Count is ambiguous between report blocks.")
    return ReportColumns(location, person_columns[0], tuple(blocks))
