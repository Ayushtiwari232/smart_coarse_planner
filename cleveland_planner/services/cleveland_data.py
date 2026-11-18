import os
import re
from datetime import date
from pathlib import Path

import pandas as pd


if os.environ.get("CLEVELAND_SCHEDULE_FILE"):
    raise ValueError(
        "CLEVELAND_SCHEDULE_FILE is no longer supported; unset it and use "
        "CLEVELAND_INPUT_DIR to select the folder containing all Cleveland inputs."
    )

INPUT_DIR = Path(
    os.environ.get(
        "CLEVELAND_INPUT_DIR",
        str(Path(__file__).resolve().parents[1] / "data" / "input" / "Cleveland Documents"),
    )
)
SCHEDULE_FILE = INPUT_DIR / "Course schedule days Cleveland.xlsx"
ASSET_BAY_FILE = INPUT_DIR / "Asset Bay Course Occupancy Cleveland.xlsx"
PLANNING_START = date(2027, 1, 1)
PLANNING_END = date(2027, 3, 31)


def load_bay_bookings() -> list[dict]:
    frame = pd.read_excel(ASSET_BAY_FILE, header=5)
    required = {"Bay", "Course Code", "Start Date", "End Date"}
    if not required.issubset(frame.columns):
        raise ValueError(f"Bay workbook is missing columns: {sorted(required - set(frame.columns))}")
    bookings = []
    for index, row in frame.iterrows():
        if pd.isna(row["Bay"]) and pd.isna(row["Course Code"]):
            continue
        bay = str(row["Bay"]).strip()
        code = str(row["Course Code"]).strip()
        start = pd.to_datetime(row["Start Date"], errors="coerce")
        end = pd.to_datetime(row["End Date"], errors="coerce")
        if pd.isna(row["Bay"]) or pd.isna(row["Course Code"]) or pd.isna(start) or pd.isna(end) or start > end:
            raise ValueError(f"Invalid bay booking on Excel row {index + 7}")
        parts = bay.split("-", 1)
        if len(parts) == 2 and all(part.strip().isdigit() for part in parts):
            first, last = (int(part.strip()) for part in parts)
            if first > last:
                raise ValueError(f"Invalid bay range on Excel row {index + 7}: {bay}")
            bays = [str(number) for number in range(first, last + 1)]
        else:
            bays = [bay]
        for bay_id in bays:
            bookings.append({
                "bay": bay_id,
                "course_code": code,
                "start_date": start.date().isoformat(),
                "end_date": end.date().isoformat(),
            })
    return bookings


def _mandatory_count(value: object, code: str) -> int:
    if pd.isna(value) or not str(value).strip():
        return 0
    text = str(value).strip()
    match = re.fullmatch(r"\d+(?:\.0+)?", text)
    if match:
        return int(float(text))
    raise ValueError(f"Mandatory count must be a non-negative integer for {code}: {value!r}")


def _lab_days(value: object, duration: int, code: str, has_bay: bool) -> list[int]:
    if pd.isna(value) or not str(value).strip():
        return list(range(1, duration + 1)) if has_bay else []
    text = " ".join(str(value).strip().lower().split())
    if text in {"virtual", "no", "none", "no lab"}:
        return []
    if text in {"yes", "full", "entire course", "all days"}:
        return list(range(1, duration + 1))
    if text == "last five days":
        if duration < 5:
            raise ValueError(f"Lab duration exceeds course duration for {code}")
        return list(range(duration - 4, duration + 1))
    if text in {"on the third day", "on the trhird day"}:
        if duration < 3:
            raise ValueError(f"Lab day exceeds course duration for {code}")
        return [3]
    raise ValueError(f"Unrecognized lab requirement for {code}: {value!r}")


def load_course_schedule(bookings: list[dict]) -> dict[str, dict]:
    frame = pd.read_excel(SCHEDULE_FILE)
    frame.columns = [str(column).strip() for column in frame.columns]
    required = {"Course Code", "Course Name", "No. of Days", "Maximal Capacity", "Lab required", "Mandatory"}
    if not required.issubset(frame.columns):
        raise ValueError(f"Cleveland schedule is missing columns: {sorted(required - set(frame.columns))}")
    courses = {}
    bay_courses = {booking["course_code"] for booking in bookings}
    for index, row in frame.iterrows():
        if pd.isna(row["Course Code"]):
            continue
        code = str(row["Course Code"]).strip()
        if code in courses:
            raise ValueError(f"Duplicate course code in Cleveland schedule: {code}")
        duration = pd.to_numeric(row["No. of Days"], errors="coerce")
        capacity = pd.to_numeric(row["Maximal Capacity"], errors="coerce")
        if pd.isna(duration) or duration <= 0 or duration != int(duration) or pd.isna(capacity) or capacity <= 0 or capacity != int(capacity):
            raise ValueError(f"Invalid duration or capacity for {code} on Excel row {index + 2}")
        duration = int(duration)
        courses[code] = {
            "course_title": str(row["Course Name"]).strip(),
            "no_of_days": duration,
            "max_capacity": int(capacity),
            "mandatory_sessions": _mandatory_count(row["Mandatory"], code),
            "lab_days": _lab_days(row["Lab required"], duration, code, code in bay_courses),
        }
    return courses


def count_existing_sessions(bookings: list[dict], start: date, end: date) -> dict[str, int]:
    occurrences = {
        (booking["course_code"], booking["start_date"], booking["end_date"])
        for booking in bookings
        if start <= date.fromisoformat(booking["start_date"]) <= end
    }
    counts: dict[str, int] = {}
    for code, _, _ in occurrences:
        counts[code] = counts.get(code, 0) + 1
    return counts
