import json
import os
import tempfile
import time
import traceback
from datetime import date, timedelta
from dataclasses import dataclass
from fractions import Fraction
from itertools import groupby

import tiktoken
import pandas as pd
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

from prompts import get_course_planner_prompt
from services.cleveland_data import (
    INPUT_DIR,
    PLANNING_END,
    PLANNING_START,
    load_bay_bookings,
)


def _count_tokens(text: str, model: str = "gpt-4o") -> int:
    """Estimate token count using tiktoken."""
    try:
        enc = tiktoken.encoding_for_model(model)
    except KeyError:
        enc = tiktoken.get_encoding("cl100k_base")
    return len(enc.encode(text))

_DEFAULT_OUTPUT_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "data", "output")
)
OUTPUT_DIR = os.environ.get("SMART_PLANNER_OUTPUT_DIR") or _DEFAULT_OUTPUT_DIR

TRAINER_LEAVE_FILE = "trainers_cleveland_q1_2027.json"
PRIORITY_FILE = INPUT_DIR / "Training Priority Matrix Cleveland.xlsx"
QUALIFICATION_FILE = INPUT_DIR / "Cleveland Trainer Course Qualification Matrix.xlsx"
LOCATION_FILE = "Smart course planner Trainer per location.xlsx"
TRAINERS_JSON_FILE = "trainers.json"
COURSE_SCHEDULE_FILE = "course_schedule_days.xlsx"

# --- Pydantic models for structured LLM output ---

class CourseSession(BaseModel):
    course_code: str = Field(description="Course code identifier")
    course_name: str = Field(description="Name of the course")
    session_number: int = Field(description="Session number (1, 2, 3, ...)")
    start_date: str = Field(description="Start date for the session (YYYY-MM-DD)")
    end_date: str = Field(description="End date for the session (YYYY-MM-DD)")
    trainer_name: str = Field(description="Assigned trainer full name")


class CoursePlan(BaseModel):
    sessions: list[CourseSession] = Field(
        description="List of all planned course sessions with dates and trainers"
    )


def _load_trainer_leave_dates() -> tuple[list[dict], str]:
    """Load Cleveland trainer leave dates from the Q1 2027 availability JSON."""
    path = os.path.join(INPUT_DIR, TRAINER_LEAVE_FILE)
    with open(path, encoding="utf-8") as source:
        data = json.load(source)
    trainers = []
    for trainer_data in data.get("trainers", []):
        leave_dates = set()
        for month_data in trainer_data.get("months", {}).values():
            for date_range in month_data.get("on_leave_ranges", []):
                start = pd.to_datetime(date_range.get("start"), errors="coerce")
                end = pd.to_datetime(date_range.get("end"), errors="coerce")
                if pd.isna(start) or pd.isna(end):
                    continue
                current_date = start.date()
                while current_date <= end.date():
                    if PLANNING_START <= current_date <= PLANNING_END:
                        leave_dates.add(current_date.isoformat())
                    current_date += timedelta(days=1)
        trainers.append(
            {
                "trainer": str(trainer_data.get("trainer", "")).strip(),
                "leave_dates": sorted(leave_dates),
            }
        )

    period = f"{PLANNING_START} to {PLANNING_END}"
    return trainers, period


def _load_trainer_priority() -> pd.DataFrame:
    """Load the Cleveland trainer priority matrix."""
    path = PRIORITY_FILE
    df = pd.read_excel(path, header=1)
    df.columns = [
        "course_code",
        "course_name",
        "priority_1",
        "priority_2",
        "priority_3",
        "priority_4",
        "priority_5",
    ]
    df = df.dropna(subset=["course_code"]).reset_index(drop=True)
    df["course_code"] = df["course_code"].astype(str).str.split().str[0]
    for column in df.columns[2:]:
        df[column] = df[column].apply(_split_priority_trainers)
    return df


def _split_priority_trainers(value: object) -> list[str]:
    if isinstance(value, list):
        return [str(name).strip() for name in value if str(name).strip()]
    if pd.isna(value):
        return []
    return [name.strip() for name in str(value).split(",") if name.strip()]


def _load_trainer_qualifications(trainers: list[dict]) -> list[dict]:
    """Load Cleveland course qualifications as trainer/course pairs."""
    path = QUALIFICATION_FILE
    df = pd.read_excel(path, header=3)
    trainer_names = {
        _normalize_trainer_name(trainer["trainer"]): trainer["trainer"]
        for trainer in trainers
    }
    qualifications = []
    for _, row in df.iterrows():
        qualified_course = row.get("Qualified Course")
        if pd.isna(qualified_course):
            continue
        course_code = str(qualified_course).strip().split()[0]
        email_name = str(row.get("Trainer", "")).split("@", 1)[0]
        email_name = " ".join(email_name.replace(".", " ").split()).title()
        profile_name = str(row.get("Philips_Profile", "")).strip()
        if "," in profile_name:
            last_name, first_name = [part.strip() for part in profile_name.split(",", 1)]
            profile_name = f"{first_name} {last_name}"
        display_name = trainer_names.get(
            _normalize_trainer_name(email_name),
            trainer_names.get(_normalize_trainer_name(profile_name), email_name),
        )
        if display_name:
            qualifications.append(
                {"course_code": course_code, "trainer_name": display_name}
            )
    return qualifications


def _priority_trainer_names(priority_df: pd.DataFrame) -> set[str]:
    priority_columns = [
        "priority_1",
        "priority_2",
        "priority_3",
        "priority_4",
        "priority_5",
    ]
    trainer_names = set()
    for column in priority_columns:
        for value in priority_df[column]:
            trainer_names.update(_split_priority_trainers(value))
    return trainer_names


def _normalize_trainer_name(name: str) -> str:
    return " ".join(str(name).strip().lower().split())


def _trainer_name_aliases(name: str) -> set[str]:
    aliases = {_normalize_trainer_name(name)}
    if "," in str(name):
        last_name, first_names = [part.strip() for part in str(name).split(",", 1)]
        aliases.add(_normalize_trainer_name(f"{first_names} {last_name}"))
        last_name_parts = last_name.split()
        if last_name_parts:
            aliases.add(_normalize_trainer_name(f"{first_names} {last_name_parts[0]}"))
    return aliases


def _build_leave_lookup(trainers: list[dict]) -> dict[str, dict]:
    leave_lookup = {}
    for trainer in trainers:
        for alias in _trainer_name_aliases(trainer["trainer"]):
            leave_lookup[alias] = trainer
    return leave_lookup


def _load_trainer_locations() -> dict[str, dict]:
    """Mark each trainer in the Cleveland Q1 roster as Cleveland-based."""
    trainers, _ = _load_trainer_leave_dates()
    result: dict[str, dict] = {}
    for trainer in trainers:
        name = trainer["trainer"]
        info = {
            "location": "Cleveland",
            "display_name": name,
        }
        for alias in _trainer_name_aliases(name):
            result[alias] = info
    return result


def _load_asset_bay_bookings() -> list[dict]:
    return load_bay_bookings()


def _working_days(start: date, end: date) -> list[date]:
    days = []
    while start <= end:
        if start.weekday() < 5:
            days.append(start)
        start += timedelta(days=1)
    return days


def _bay_resources(bookings: list[dict]) -> tuple[dict[str, set[str]], dict[str, set[date]]]:
    course_bays: dict[str, set[str]] = {}
    reserved_by_bay: dict[str, set[date]] = {}
    for booking in bookings:
        bay = booking["bay"]
        course_bays.setdefault(booking["course_code"], set()).add(bay)
        reserved_by_bay.setdefault(bay, set()).update(
            _working_days(date.fromisoformat(booking["start_date"]), date.fromisoformat(booking["end_date"]))
        )
    return course_bays, reserved_by_bay


@dataclass(frozen=True)
class _ResourceSlot:
    days: tuple[date, ...]
    lab_days: frozenset[date]
    trainer: str
    bay: str | None


def _schedule_resource_feasible_sessions(
    courses: list[dict], trainers: list[dict], qualifications: list[dict], bookings: list[dict],
) -> list[dict]:
    """Search dates, trainers and bays jointly without relaxing resource constraints."""
    required = {
        course["course_code"]: course for course in courses if course["sessions_needed"]
    }
    if not required:
        return []
    working_days = _working_days(PLANNING_START, PLANNING_END)
    roster = {_normalize_trainer_name(trainer["trainer"]): trainer for trainer in trainers}
    aliases = {
        alias: key for key, trainer in roster.items()
        for alias in _trainer_name_aliases(trainer["trainer"])
    }
    leave = {
        key: {date.fromisoformat(day) for day in trainer["leave_dates"]}
        for key, trainer in roster.items()
    }
    available = {key: len(set(working_days) - leave[key]) for key in roster}
    qualified: dict[str, set[str]] = {}
    for entry in qualifications:
        key = aliases.get(_normalize_trainer_name(entry["trainer_name"]))
        if key is not None:
            qualified.setdefault(entry["course_code"], set()).add(key)
    course_bays, reserved = _bay_resources(bookings)
    slots: dict[str, list[_ResourceSlot]] = {}
    for code, course in required.items():
        duration = course["no_of_days"]
        choices = []
        for index in range(len(working_days) - duration + 1):
            days = tuple(working_days[index:index + duration])
            if duration in (3, 4) and days[0].isocalendar()[:2] != days[-1].isocalendar()[:2]:
                continue
            lab_days = frozenset(days[day - 1] for day in course["lab_days"])
            bays: list[str | None] = (
                sorted(course_bays.get(code, set())) if lab_days else [None]
            )
            for trainer in sorted(qualified.get(code, set())):
                if available[trainer] == 0 or not leave[trainer].isdisjoint(days):
                    continue
                for bay in bays:
                    if bay is None or reserved.get(bay, set()).isdisjoint(lab_days):
                        choices.append(_ResourceSlot(days, lab_days, trainer, bay))
        if not choices:
            raise ValueError(
                f"No feasible dates, qualified trainer and compatible bay for {code} "
                f"within {PLANNING_START} to {PLANNING_END}; check leave and fixed bay bookings"
            )
        slots[code] = choices

    assigned: dict[str, list[_ResourceSlot]] = {code: [] for code in required}
    busy: dict[str, set[date]] = {key: set() for key in roster}
    loads = {key: 0 for key in roster}
    attempts = 0
    search_limit = 50000

    def search() -> bool:
        nonlocal attempts
        feasible: dict[str, list[_ResourceSlot]] = {}
        for code, course in required.items():
            if len(assigned[code]) == course["sessions_needed"]:
                continue
            previous_start = assigned[code][-1].days[0] if assigned[code] else None
            options = [
                slot for slot in slots[code]
                if (previous_start is None or slot.days[0] > previous_start)
                and busy[slot.trainer].isdisjoint(slot.days)
                and (slot.bay is None or reserved.get(slot.bay, set()).isdisjoint(slot.lab_days))
            ]
            if not options:
                return False
            feasible[code] = options
        if not feasible:
            return True
        code = min(feasible, key=lambda value: (
            len(feasible[value]), -required[value]["no_of_days"], value,
        ))
        course = required[code]
        target_index = round(
            (len(working_days) - course["no_of_days"])
            * (len(assigned[code]) + 0.5) / course["sessions_needed"]
        )
        target = working_days[target_index]
        options = sorted(feasible[code], key=lambda slot: (
            Fraction(loads[slot.trainer] + len(slot.days), available[slot.trainer]),
            abs((slot.days[0] - target).days), slot.days[0], slot.trainer, slot.bay or "",
        ))
        for slot in options:
            attempts += 1
            if attempts > search_limit:
                raise ValueError(
                    f"Resource scheduling search limit ({search_limit}) reached for "
                    + ", ".join(sorted(required))
                    + "; no valid complete plan has been established"
                )
            assigned[code].append(slot)
            busy[slot.trainer].update(slot.days)
            loads[slot.trainer] += len(slot.days)
            if slot.bay is not None:
                reserved.setdefault(slot.bay, set()).update(slot.lab_days)
            if search():
                return True
            if slot.bay is not None:
                reserved[slot.bay].difference_update(slot.lab_days)
            loads[slot.trainer] -= len(slot.days)
            busy[slot.trainer].difference_update(slot.days)
            assigned[code].pop()
        return False

    if not search():
        raise ValueError(
            "No complete resource-feasible Cleveland schedule exists for "
            + ", ".join(sorted(required))
            + "; check mandatory counts, qualifications, leave and fixed bay bookings"
        )
    rows = []
    for code, choices in assigned.items():
        for number, slot in enumerate(choices, start=1):
            rows.append({
                "course_code": code,
                "course_name": required[code]["course_title"],
                "session_number": number,
                "start_date": slot.days[0].isoformat(),
                "end_date": slot.days[-1].isoformat(),
                "trainer_name": roster[slot.trainer]["trainer"],
                "bay_room": f"Bay {slot.bay}" if slot.bay is not None else "Not required",
                "lab_dates": sorted(day.isoformat() for day in slot.lab_days),
            })
    return sorted(rows, key=lambda row: (row["start_date"], row["course_code"], row["session_number"]))


def _assign_bay_rooms(sessions: list[dict], bookings: list[dict], courses: list[dict]) -> None:
    """Allocate compatible bays for the actual lab days, or fail explicitly."""
    course_bays, reserved_by_bay = _bay_resources(bookings)
    requirements = {course["course_code"]: course for course in courses}
    pending = []
    for session in sessions:
        course = requirements[session["course_code"]]
        lab_days = course["lab_days"]
        if not lab_days:
            session["bay_room"] = "Not required"
            session["lab_dates"] = []
            continue
        working_days = _working_days(date.fromisoformat(session["start_date"]), date.fromisoformat(session["end_date"]))
        dates = {working_days[index - 1] for index in lab_days}
        candidates = sorted(course_bays.get(session["course_code"], set()))
        if not candidates:
            raise ValueError(f"No compatible bay is listed for {session['course_code']}")
        pending.append((session, dates, candidates))

    pending.sort(key=lambda item: len(item[2]))

    def assign(index: int) -> bool:
        if index == len(pending):
            return True
        session, dates, candidates = pending[index]
        for bay in candidates:
            reserved = reserved_by_bay.setdefault(bay, set())
            if reserved.isdisjoint(dates):
                reserved.update(dates)
                session["bay_room"] = f"Bay {bay}"
                session["lab_dates"] = sorted(day.isoformat() for day in dates)
                if assign(index + 1):
                    return True
                reserved.difference_update(dates)
        return False

    if not assign(0):
        raise ValueError("Cannot assign compatible bays without conflicting with existing or new lab bookings")


def _validate_sessions(sessions: list[dict], courses: list[dict], trainers: list[dict], qualifications: list[dict]) -> None:
    requirements = {course["course_code"]: course for course in courses}
    leave = _build_leave_lookup(trainers)
    qualified = {(q["course_code"], _normalize_trainer_name(q["trainer_name"])) for q in qualifications}
    occupied: dict[str, set[date]] = {}
    counts: dict[str, int] = {}
    session_ids = set()
    for session in sessions:
        code = session["course_code"].strip()
        session["course_code"] = code
        course = requirements.get(code)
        if course is None or course["sessions_needed"] is None or course["sessions_needed"] <= 0:
            raise ValueError(f"Unexpected course in generated plan: {code}")
        session_id = (code, session["session_number"])
        if session["session_number"] < 1 or session_id in session_ids:
            raise ValueError(f"Invalid or duplicate session number: {session_id}")
        session_ids.add(session_id)
        try:
            start = date.fromisoformat(session["start_date"])
            end = date.fromisoformat(session["end_date"])
        except ValueError as exc:
            raise ValueError(f"Invalid session dates for {code}") from exc
        days = _working_days(start, end)
        if (start < PLANNING_START or end > PLANNING_END or start > end
                or start.weekday() >= 5 or end.weekday() >= 5
                or len(days) != course["no_of_days"]):
            raise ValueError(f"Invalid planning period or duration for {code}: {start} to {end}")
        if len(days) in (3, 4) and start.isocalendar()[:2] != end.isocalendar()[:2]:
            raise ValueError(f"Session for {code} cannot cross a weekend")
        name = session["trainer_name"]
        trainer = next((leave[alias] for alias in _trainer_name_aliases(name) if alias in leave), None)
        if trainer is None or (code, _normalize_trainer_name(trainer["trainer"])) not in qualified:
            raise ValueError(f"Trainer {name} is not a qualified Cleveland roster trainer for {code}")
        trainer_key = _normalize_trainer_name(trainer["trainer"])
        busy = occupied.setdefault(trainer_key, set())
        if not busy.isdisjoint(days) or any(day.isoformat() in trainer["leave_dates"] for day in days):
            raise ValueError(f"Trainer {name} is unavailable for {code}")
        busy.update(days)
        counts[code] = counts.get(code, 0) + 1
        if counts[code] > course["sessions_needed"]:
            raise ValueError(f"Too many new sessions for {code}")
    for course in courses:
        missing = course["mandatory_sessions"] - counts.get(course["course_code"], 0)
        if missing > 0:
            raise ValueError(f"Mandatory minimum not met for {course['course_code']}: {missing} session(s) missing")


def _balance_trainer_assignments(
    sessions: list[dict],
    trainers: list[dict],
    qualifications: list[dict],
    priority_df: pd.DataFrame,
) -> None:
    """Find feasible assignments, then reduce utilization variance by moves and swaps."""
    if not sessions:
        return
    roster = {_normalize_trainer_name(trainer["trainer"]): trainer for trainer in trainers}
    aliases = {
        alias: key for key, trainer in roster.items()
        for alias in _trainer_name_aliases(trainer["trainer"])
    }
    working_days = set(_working_days(PLANNING_START, PLANNING_END))
    leave = {
        key: {date.fromisoformat(day) for day in trainer["leave_dates"]} & working_days
        for key, trainer in roster.items()
    }
    available = {key: len(working_days - leave[key]) for key in roster}
    qualified: dict[str, set[str]] = {}
    for entry in qualifications:
        key = aliases.get(_normalize_trainer_name(entry["trainer_name"]))
        if key is not None:
            qualified.setdefault(entry["course_code"], set()).add(key)
    ranks: dict[tuple[str, str], int] = {}
    for _, row in priority_df.iterrows():
        for rank in range(1, 6):
            for name in _split_priority_trainers(row[f"priority_{rank}"]):
                key = aliases.get(_normalize_trainer_name(name))
                if key is not None:
                    pair = (str(row["course_code"]), key)
                    ranks[pair] = min(rank, ranks.get(pair, rank))

    session_days = []
    candidates = []
    for session in sessions:
        start = date.fromisoformat(session["start_date"])
        end = date.fromisoformat(session["end_date"])
        if start > end or start < PLANNING_START or end > PLANNING_END:
            raise ValueError(f"Invalid planning period for {session['course_code']}: {start} to {end}")
        days = set(_working_days(start, end))
        session_days.append(days)
        eligible = sorted(
            key for key in qualified.get(session["course_code"].strip(), set())
            if available[key] > 0 and leave[key].isdisjoint(days)
        )
        if not eligible:
            raise ValueError(f"No qualified trainer is available for {session['course_code']} on {start} to {end}")
        candidates.append(eligible)

    assignments: list[str | None] = [None] * len(sessions)
    loads = {key: 0 for key in roster}

    def can_assign(index: int, key: str, ignored: tuple[int, ...] = ()) -> bool:
        return all(
            other == index or other in ignored or assigned != key
            or session_days[index].isdisjoint(session_days[other])
            for other, assigned in enumerate(assignments)
        )

    def find_assignment() -> bool:
        remaining = [index for index, key in enumerate(assignments) if key is None]
        if not remaining:
            return True
        options = {
            index: [key for key in candidates[index] if can_assign(index, key)]
            for index in remaining
        }
        index = min(remaining, key=lambda item: (len(options[item]), -len(session_days[item]), item))
        duration = len(session_days[index])
        for key in sorted(options[index], key=lambda name: (
            Fraction(loads[name] + duration, available[name]),
            ranks.get((sessions[index]["course_code"], name), 6), name,
        )):
            assignments[index] = key
            loads[key] += duration
            if find_assignment():
                return True
            loads[key] -= duration
            assignments[index] = None
        return False

    if not find_assignment():
        raise ValueError("No conflict-free qualified trainer assignment exists for the proposed session dates")

    def score() -> tuple[Fraction, int]:
        teaching = {key: 0 for key in roster}
        priority_cost = 0
        for index, key in enumerate(assignments):
            if key is not None:
                teaching[key] += len(session_days[index])
                priority_cost += ranks.get((sessions[index]["course_code"], key), 6)
        ratios = [
            Fraction(teaching[key], available[key]) if available[key] else Fraction(0)
            for key in roster
        ]
        variance = sum((ratio * ratio for ratio in ratios), Fraction(0))
        variance -= sum(ratios, Fraction(0)) ** 2 / len(ratios)
        return variance, priority_cost

    while True:
        best_score = score()
        best_assignment = None
        for index, current in enumerate(assignments):
            for key in candidates[index]:
                if key == current or not can_assign(index, key):
                    continue
                assignments[index] = key
                candidate_score = score()
                if candidate_score < best_score:
                    best_score, best_assignment = candidate_score, list(assignments)
                assignments[index] = current
            for other in range(index + 1, len(sessions)):
                other_key = assignments[other]
                if (other_key == current or other_key not in candidates[index]
                        or current not in candidates[other]):
                    continue
                if (not can_assign(index, other_key, (other,))
                        or not can_assign(other, current, (index,))):
                    continue
                assignments[index], assignments[other] = other_key, current
                candidate_score = score()
                if candidate_score < best_score:
                    best_score, best_assignment = candidate_score, list(assignments)
                assignments[index], assignments[other] = current, other_key
        if best_assignment is None:
            break
        assignments = best_assignment

    for session, key in zip(sessions, assignments):
        if key is None:
            raise ValueError(f"Trainer assignment is missing for {session['course_code']}")
        session["trainer_name"] = roster[key]["trainer"]


def _trainer_location(trainer_name: str, location_lookup: dict[str, dict]) -> str:
    """Return the trainer's assigned location, or Others when it is unknown."""
    for alias in _trainer_name_aliases(trainer_name):
        if alias in location_lookup:
            return location_lookup[alias]["location"]
    return "Others"


def _load_classroom_assignments() -> dict[str, dict[str, str]]:
    """Load default and CTC-specific classroom assignments by course code.

    The CTC room column is optional to retain compatibility with older schedule files.
    """
    path = os.path.join(INPUT_DIR, COURSE_SCHEDULE_FILE)
    try:
        schedule_df = pd.read_excel(path)
        normalized_columns = {
            " ".join(str(column).strip().lower().split()): column
            for column in schedule_df.columns
        }
        ctc_room_column = normalized_columns.get("ctc training rooms per class")
        assignments: dict[str, dict[str, str]] = {}
        for _, row in schedule_df.iterrows():
            course_code = str(row.get("Course Code", "")).strip()
            if not course_code or course_code.lower() == "nan" or course_code in assignments:
                continue
            default_room = str(row.get("Preferred Classroom", "")).strip()
            ctc_room = str(row.get(ctc_room_column, "")).strip() if ctc_room_column else ""
            assignments[course_code] = {
                "default": "" if default_room.lower() == "nan" else default_room,
                "ctc": "" if ctc_room.lower() == "nan" else ctc_room,
            }
        return assignments
    except Exception as e:
        print(f"[PLANNER] WARNING: Could not load classroom assignments: {e}")
        return {}


def _classroom_for_session(
    course_code: str,
    trainer_name: str,
    classroom_assignments: dict[str, dict[str, str]],
    location_lookup: dict[str, dict],
) -> str:
    """Choose the CTC room for CTC trainers; otherwise choose the default room."""
    rooms = classroom_assignments.get(course_code, {})
    if _trainer_location(trainer_name, location_lookup) == "CTC" and rooms.get("ctc"):
        return rooms["ctc"]
    return rooms.get("default", "")


def _load_parttime_days_from_json() -> dict[str, set]:
    """Extract dates with '%' marker from trainers.json (part-time / half-day indicator).
    Returns dict[normalized_alias → set[date]]."""
    path = os.path.join(INPUT_DIR, TRAINERS_JSON_FILE)
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        result: dict[str, set] = {}
        for trainer in data.get("trainers", []):
            tname = trainer["trainer"]
            pdays: set = set()
            for month_data in trainer.get("months", {}).values():
                for range_type in ("available_ranges", "on_leave_ranges"):
                    for r in month_data.get(range_type, []):
                        if r.get("marker") == "%":
                            rd_start = pd.to_datetime(r["start"], errors="coerce")
                            rd_end = pd.to_datetime(r["end"], errors="coerce")
                            if pd.isna(rd_start) or pd.isna(rd_end):
                                continue
                            cur = rd_start.date()
                            while cur <= rd_end.date():
                                pdays.add(cur)
                                cur += timedelta(days=1)
            for alias in _trainer_name_aliases(tname):
                result[alias] = pdays
        return result
    except Exception as e:
        print(f"[PLANNER] WARNING: Could not load parttime days from trainers.json: {e}")
        return {}


def _get_planning_period_from_json() -> tuple:
    """Parse the planning period string from trainers.json. Returns (start_date, end_date) or (None, None)."""
    path = os.path.join(INPUT_DIR, TRAINERS_JSON_FILE)
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        period_str = data.get("period", "")
        if " to " in period_str:
            parts = period_str.split(" to ")
            start = pd.to_datetime(parts[0].strip(), errors="coerce")
            end = pd.to_datetime(parts[1].strip(), errors="coerce")
            if not pd.isna(start) and not pd.isna(end):
                return start.date(), end.date()
    except Exception as e:
        print(f"[PLANNER] WARNING: Could not parse planning period from trainers.json: {e}")
    return None, None


def _parse_classrooms(classroom_raw: str, session_start: date, session_end: date) -> list:
    """Parse classroom assignment string into list of (lab_name, start_date, end_date).
    Handles multi-lab strings like 'QS003 wk1+2\\nQS147 wk3'."""
    lines = [ln.strip() for ln in classroom_raw.strip().split("\n") if ln.strip()]
    if len(lines) <= 1:
        lab = classroom_raw.strip().split()[0] if classroom_raw.strip().split() else classroom_raw
        return [(lab, session_start, session_end)]

    result = []
    for line in lines:
        parts = line.split()
        if not parts:
            continue
        lab = parts[0]
        wk_part = next((p[2:] for p in parts[1:] if p.lower().startswith("wk")), None)
        if wk_part:
            week_nums = []
            for seg in wk_part.replace("+", ",").split(","):
                try:
                    week_nums.append(int(seg.strip()))
                except ValueError:
                    pass
            if week_nums:
                min_wk = min(week_nums) - 1
                max_wk = max(week_nums) - 1
                lab_start = session_start + timedelta(weeks=min_wk)
                lab_end = min(session_start + timedelta(weeks=max_wk + 1) - timedelta(days=1), session_end)
                result.append((lab, lab_start, lab_end))
                continue
        result.append((lab, session_start, session_end))
    return result if result else [(classroom_raw.split()[0] if classroom_raw.split() else classroom_raw, session_start, session_end)]


def _write_trainer_availability_sheet(
    output_path: str,
    trainers: list[dict],
    sessions: list[dict],
    priority_trainer_names: set[str],
) -> None:
    """Append a 'Trainer Availability' sheet to the plan workbook.

    Layout:
      Row 1 : month labels (merged across that month's columns)
      Row 2 : day numbers
      Row 3+ : trainer rows, grouped by location
    Colors:
      light blue   – teaching (course code shown; merged across consecutive days)
      light red    – on leave
      light yellow – part-time / half-day (% marker from source data)
      light green  – weekend
    A legend block is appended below the trainer rows.
    """
    # --- Colour palette ---
    LIGHT_BLUE   = PatternFill(start_color="FFADD8E6", end_color="FFADD8E6", fill_type="solid")
    LIGHT_RED    = PatternFill(start_color="FFFFB6C1", end_color="FFFFB6C1", fill_type="solid")
    WEEKEND_FILL = PatternFill(start_color="FFE8F5E9", end_color="FFE8F5E9", fill_type="solid")
    PARTTIME_FILL= PatternFill(start_color="FFFFF59D", end_color="FFFFF59D", fill_type="solid")
    HEADER_FILL  = PatternFill(start_color="FF1F4E78", end_color="FF1F4E78", fill_type="solid")
    MONTH_FILL   = PatternFill(start_color="FF2E75B6", end_color="FF2E75B6", fill_type="solid")
    DAYNUM_FILL  = PatternFill(start_color="FF4472C4", end_color="FF4472C4", fill_type="solid")
    LOC_FILL_MAP = {
        "PHC":    PatternFill(start_color="FFD6E4BC", end_color="FFD6E4BC", fill_type="solid"),
        "SLC":    PatternFill(start_color="FFD9E1F2", end_color="FFD9E1F2", fill_type="solid"),
        "CTC":    PatternFill(start_color="FFFCE4D6", end_color="FFFCE4D6", fill_type="solid"),
        "Others": PatternFill(start_color="FFF2EFEA", end_color="FFF2EFEA", fill_type="solid"),
    }
    DEFAULT_LOC_FILL = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")

    HEADER_FONT     = Font(bold=True, color="FFFFFFFF")
    SESSION_FONT    = Font(size=7, bold=True)
    SESSION_ALIGN   = Alignment(horizontal="center", vertical="center", wrap_text=False)
    LOC_HEADER_FONT = Font(bold=True, size=10)

    # --- Extra data ---
    location_lookup = _load_trainer_locations()
    parttime_lookup = _load_parttime_days_from_json()
    start_d, end_d = PLANNING_START, PLANNING_END

    date_list: list[date] = []
    cur = start_d
    while cur <= end_d:
        date_list.append(cur)
        cur += timedelta(days=1)

    date_to_col: dict[date, int] = {dt: i + 2 for i, dt in enumerate(date_list)}
    total_cols = len(date_list) + 1  # col A + date columns

    # --- Trainer sessions lookup (alias-aware) ---
    trainer_sessions: dict[str, list[dict]] = {}
    for s in sessions:
        norm = _normalize_trainer_name(s.get("trainer_name", ""))
        if not norm:
            continue
        sd = pd.to_datetime(s.get("start_date"), errors="coerce")
        ed = pd.to_datetime(s.get("end_date"), errors="coerce")
        if pd.isna(sd) or pd.isna(ed):
            continue
        trainer_sessions.setdefault(norm, []).append({
            "start":          sd.date(),
            "end":            ed.date(),
            "course_code":    str(s.get("course_code",    "") or "").strip(),
            "course_name":    str(s.get("course_name",    "") or "").strip(),
            "session_number": s.get("session_number"),
        })

    leave_lookup = _build_leave_lookup(trainers)

    # --- Deduplicated trainer list (fix: alias-aware dedup prevents "Last, First" duplicates) ---
    trainer_names: list[str] = []
    seen: set[str] = set()

    def _add_trainer(name: str) -> None:
        cleaned = str(name).strip()
        if not cleaned:
            return
        aliases = _trainer_name_aliases(cleaned)
        if aliases & seen:   # any alias already seen → skip
            return
        seen.update(aliases)
        trainer_names.append(cleaned)

    for trainer in trainers:
        _add_trainer(trainer["trainer"])
    for name in sorted(priority_trainer_names):
        _add_trainer(name)
    for s in sessions:
        _add_trainer(s.get("trainer_name", ""))

    # --- Group trainers by location ---
    location_groups: dict[str, list[str]] = {}
    others: list[str] = []
    for name in trainer_names:
        info = None
        for alias in _trainer_name_aliases(name):
            if alias in location_lookup:
                info = location_lookup[alias]
                break
        if info:
            loc = info["location"]
            location_groups.setdefault(loc, []).append(name)
        else:
            others.append(name)
    if others:
        location_groups["Others"] = others

    # --- Create / replace sheet ---
    wb = load_workbook(output_path)
    if "Trainer Availability" in wb.sheetnames:
        del wb["Trainer Availability"]
    ws = wb.create_sheet("Trainer Availability")

    # --- Row 1: "Trainer" label (spans rows 1-2) + month headers ---
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    hc = ws.cell(row=1, column=1, value="Trainer")
    hc.fill = HEADER_FILL
    hc.font = HEADER_FONT
    hc.alignment = Alignment(horizontal="center", vertical="center")

    for (yr, mo), grp in groupby(date_list, key=lambda d: (d.year, d.month)):
        month_dates = list(grp)
        col_s = date_to_col[month_dates[0]]
        col_e = date_to_col[month_dates[-1]]
        if col_s < col_e:
            ws.merge_cells(start_row=1, start_column=col_s, end_row=1, end_column=col_e)
        mc = ws.cell(row=1, column=col_s, value=month_dates[0].strftime("%B %Y"))
        mc.fill = MONTH_FILL
        mc.font = HEADER_FONT
        mc.alignment = Alignment(horizontal="center", vertical="center")

    # --- Row 2: day numbers ---
    for dt in date_list:
        col = date_to_col[dt]
        dc = ws.cell(row=2, column=col, value=dt.day)
        dc.fill = DAYNUM_FILL
        dc.font = Font(bold=True, color="FFFFFFFF", size=7)
        dc.alignment = Alignment(horizontal="center", vertical="center")

    # --- Trainer rows, grouped by location ---
    current_row = 3
    trainer_row_indices: list[int] = []
    course_cells: dict[tuple, str] = {}  # (row, col) → course_code value

    for loc, names in location_groups.items():
        # Location group header
        loc_fill = LOC_FILL_MAP.get(loc, DEFAULT_LOC_FILL)
        ws.merge_cells(start_row=current_row, start_column=1,
                       end_row=current_row, end_column=total_cols)
        lc = ws.cell(row=current_row, column=1, value=f"  {loc}")
        lc.fill = loc_fill
        lc.font = LOC_HEADER_FONT
        lc.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[current_row].height = 18
        current_row += 1

        for trainer_name in names:
            # Resolve leave dates via all aliases
            leave_set: set[date] = set()
            for alias in _trainer_name_aliases(trainer_name):
                record = leave_lookup.get(alias)
                if not record:
                    continue
                for raw in record.get("leave_dates", []):
                    parsed = pd.to_datetime(raw, errors="coerce")
                    if not pd.isna(parsed):
                        leave_set.add(parsed.date())

            # Resolve part-time days
            parttime_set: set[date] = set()
            for alias in _trainer_name_aliases(trainer_name):
                parttime_set.update(parttime_lookup.get(alias, set()))

            # Resolve session ranges
            ranges: list[dict] = []
            for alias in _trainer_name_aliases(trainer_name):
                ranges.extend(trainer_sessions.get(alias, []))

            ws.cell(row=current_row, column=1, value=trainer_name)

            for dt in date_list:
                col = date_to_col[dt]
                cell = ws.cell(row=current_row, column=col)
                matching = [r for r in ranges if r["start"] <= dt <= r["end"]]

                if matching:
                    cell.fill = LIGHT_BLUE
                    codes = [r["course_code"] for r in matching if r["course_code"]]
                    if codes:
                        val = " / ".join(dict.fromkeys(codes))
                        cell.value = val
                        cell.font  = SESSION_FONT
                        cell.alignment = SESSION_ALIGN
                        course_cells[(current_row, col)] = val
                    # Hover comment
                    comment_lines = []
                    for r in matching:
                        suffix = f" — Session {r['session_number']}" if r["session_number"] else ""
                        comment_lines.append(
                            f"{r['course_code'] or '?'}: {r['course_name'] or '(unknown)'}{suffix}"
                        )
                    if comment_lines:
                        cmt = Comment("\n".join(comment_lines), "Planner")
                        cmt.width  = 260
                        cmt.height = 20 + 14 * len(comment_lines)
                        cell.comment = cmt

                elif dt in leave_set:
                    cell.fill = LIGHT_RED
                    if dt in parttime_set:
                        cell.value     = "%"
                        cell.font      = Font(size=7, bold=True, color="FF8B0000")
                        cell.alignment = SESSION_ALIGN

                elif dt in parttime_set:
                    cell.fill      = PARTTIME_FILL
                    cell.value     = "%"
                    cell.font      = Font(size=7, color="FF856404")
                    cell.alignment = SESSION_ALIGN

                elif dt.weekday() >= 5:
                    cell.fill = WEEKEND_FILL

            trainer_row_indices.append(current_row)
            current_row += 1

    # --- Merge consecutive course-code cells within each trainer row ---
    for row_idx in trainer_row_indices:
        col = 2
        max_col = len(date_list) + 1
        while col <= max_col:
            pos = (row_idx, col)
            if pos in course_cells:
                val       = course_cells[pos]
                start_col = col
                while col + 1 <= max_col and course_cells.get((row_idx, col + 1)) == val:
                    col += 1
                end_col = col
                if end_col > start_col:
                    ws.merge_cells(start_row=row_idx, start_column=start_col,
                                   end_row=row_idx, end_column=end_col)
                    mc = ws.cell(row=row_idx, column=start_col)
                    mc.value     = val
                    mc.fill      = LIGHT_BLUE
                    mc.font      = SESSION_FONT
                    mc.alignment = SESSION_ALIGN
            col += 1

    # --- Legend ---
    legend_row = current_row + 1
    ws.cell(row=legend_row, column=1, value="Legend").font = Font(bold=True, size=10)
    legend_row += 1
    legend_items = [
        (LIGHT_BLUE,    "Teaching / Course assigned"),
        (LIGHT_RED,     "On leave"),
        (PARTTIME_FILL, "Part-time / half-day available (% marker)"),
        (WEEKEND_FILL,  "Weekend"),
    ]
    thin = Side(style="thin", color="FF999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for fill, label in legend_items:
        color_cell = ws.cell(row=legend_row, column=1, value=" ")
        color_cell.fill   = fill
        color_cell.border = border
        label_cell = ws.cell(row=legend_row, column=2, value=label)
        label_cell.font = Font(size=9)
        legend_row += 1

    # --- Column / row sizing ---
    ws.column_dimensions["A"].width = 28
    for col_idx in range(2, len(date_list) + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = 4
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "B3"

    wb.save(output_path)


def _write_lab_availability_sheet(output_path: str, sessions: list[dict]) -> None:
    """Show existing and newly assigned Cleveland bay occupancy."""
    location_lookup = _load_trainer_locations()

    plan_start, plan_end = PLANNING_START, PLANNING_END

    date_list: list[date] = []
    cur = plan_start
    while cur <= plan_end:
        date_list.append(cur)
        cur += timedelta(days=1)
    date_to_col = {dt: i + 2 for i, dt in enumerate(date_list)}
    total_cols = len(date_list) + 1

    # Build lab assignments and location map
    lab_assignments: dict[str, list[dict]] = {}
    lab_locations:   dict[str, str]        = {}

    for booking in _load_asset_bay_bookings():
        lab = f"Bay {booking['bay']}"
        lab_locations[lab] = "Cleveland"
        lab_assignments.setdefault(lab, []).append({
            "start": date.fromisoformat(booking["start_date"]),
            "end": date.fromisoformat(booking["end_date"]),
            "course_code": booking["course_code"],
            "course_name": "Existing booking",
            "trainer": "",
        })
    for session in sessions:
        if not session["lab_dates"]:
            continue
        lab = session["bay_room"]
        lab_locations[lab] = "Cleveland"
        for lab_date in session["lab_dates"]:
            day = date.fromisoformat(lab_date)
            lab_assignments.setdefault(lab, []).append({
                "start": day,
                "end": day,
                "course_code": session["course_code"],
                "course_name": session["course_name"],
                "trainer": session["trainer_name"],
            })

    if not lab_assignments:
        print("[PLANNER] No lab data to write – skipping Lab Availability sheet")
        return

    # Group labs by location, ensuring all location groups from trainer data are present
    loc_labs: dict[str, list[str]] = {}
    # Pre-populate all known location groups (same as trainer sheet) in order
    all_locations = list(dict.fromkeys(
        info["location"] for info in location_lookup.values() if info.get("location")
    ))
    for loc in all_locations:
        loc_labs.setdefault(loc, [])
    loc_labs.setdefault("Others", [])
    # Assign labs to their location groups
    for lab in sorted(lab_locations):
        loc = lab_locations[lab]
        loc_labs.setdefault(loc, []).append(lab)

    # --- Colours (same palette as trainer sheet) ---
    COURSE_FILL  = PatternFill(start_color="FFADD8E6", end_color="FFADD8E6", fill_type="solid")
    WEEKEND_FILL = PatternFill(start_color="FFE8F5E9", end_color="FFE8F5E9", fill_type="solid")
    HEADER_FILL  = PatternFill(start_color="FF1F4E78", end_color="FF1F4E78", fill_type="solid")
    MONTH_FILL   = PatternFill(start_color="FF2E75B6", end_color="FF2E75B6", fill_type="solid")
    DAYNUM_FILL  = PatternFill(start_color="FF4472C4", end_color="FF4472C4", fill_type="solid")
    LOC_FILL_MAP = {
        "PHC":    PatternFill(start_color="FFD6E4BC", end_color="FFD6E4BC", fill_type="solid"),
        "SLC":    PatternFill(start_color="FFD9E1F2", end_color="FFD9E1F2", fill_type="solid"),
        "CTC":    PatternFill(start_color="FFFCE4D6", end_color="FFFCE4D6", fill_type="solid"),
        "Others": PatternFill(start_color="FFF2EFEA", end_color="FFF2EFEA", fill_type="solid"),
    }
    DEFAULT_LOC_FILL = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")
    HEADER_FONT   = Font(bold=True, color="FFFFFFFF")
    SESSION_FONT  = Font(size=7, bold=True)
    SESSION_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=False)

    wb = load_workbook(output_path)
    if "Lab Availability" in wb.sheetnames:
        del wb["Lab Availability"]
    ws = wb.create_sheet("Lab Availability")

    # Row 1+2 header
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    hc = ws.cell(row=1, column=1, value="Lab / Room")
    hc.fill = HEADER_FILL
    hc.font = HEADER_FONT
    hc.alignment = Alignment(horizontal="center", vertical="center")

    for (yr, mo), grp in groupby(date_list, key=lambda d: (d.year, d.month)):
        month_dates = list(grp)
        col_s = date_to_col[month_dates[0]]
        col_e = date_to_col[month_dates[-1]]
        if col_s < col_e:
            ws.merge_cells(start_row=1, start_column=col_s, end_row=1, end_column=col_e)
        mc = ws.cell(row=1, column=col_s, value=month_dates[0].strftime("%B %Y"))
        mc.fill = MONTH_FILL
        mc.font = HEADER_FONT
        mc.alignment = Alignment(horizontal="center", vertical="center")

    for dt in date_list:
        col = date_to_col[dt]
        dc = ws.cell(row=2, column=col, value=dt.day)
        dc.fill = DAYNUM_FILL
        dc.font = Font(bold=True, color="FFFFFFFF", size=7)
        dc.alignment = Alignment(horizontal="center", vertical="center")

    current_row = 3
    lab_row_indices: list[int] = []
    course_cells: dict[tuple, str] = {}

    for loc, labs in loc_labs.items():
        loc_fill = LOC_FILL_MAP.get(loc, DEFAULT_LOC_FILL)
        ws.merge_cells(start_row=current_row, start_column=1,
                       end_row=current_row, end_column=total_cols)
        lc = ws.cell(row=current_row, column=1, value=f"  {loc}")
        lc.fill = loc_fill
        lc.font = Font(bold=True, size=10)
        lc.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[current_row].height = 18
        current_row += 1

        for lab in labs:
            ws.cell(row=current_row, column=1, value=lab)
            lab_info = lab_assignments.get(lab, [])

            for dt in date_list:
                col  = date_to_col[dt]
                cell = ws.cell(row=current_row, column=col)
                matching = [r for r in lab_info if dt.weekday() < 5 and r["start"] <= dt <= r["end"]]
                if matching:
                    cell.fill = COURSE_FILL
                    codes = [r["course_code"] for r in matching if r["course_code"]]
                    if codes:
                        val = " / ".join(dict.fromkeys(codes))
                        cell.value     = val
                        cell.font      = SESSION_FONT
                        cell.alignment = SESSION_ALIGN
                        course_cells[(current_row, col)] = val
                    comment_lines = [
                        f"{r['course_code']}: {r['course_name']} ({r['trainer']})"
                        for r in matching
                    ]
                    if comment_lines:
                        cmt = Comment("\n".join(comment_lines), "Planner")
                        cmt.width  = 300
                        cmt.height = 20 + 14 * len(comment_lines)
                        cell.comment = cmt
                elif dt.weekday() >= 5:
                    cell.fill = WEEKEND_FILL

            lab_row_indices.append(current_row)
            current_row += 1

    # Merge consecutive same-course cells
    for row_idx in lab_row_indices:
        col = 2
        max_col = len(date_list) + 1
        while col <= max_col:
            pos = (row_idx, col)
            if pos in course_cells:
                val       = course_cells[pos]
                start_col = col
                while col + 1 <= max_col and course_cells.get((row_idx, col + 1)) == val:
                    col += 1
                end_col = col
                if end_col > start_col:
                    ws.merge_cells(start_row=row_idx, start_column=start_col,
                                   end_row=row_idx, end_column=end_col)
                    mc = ws.cell(row=row_idx, column=start_col)
                    mc.value     = val
                    mc.fill      = COURSE_FILL
                    mc.font      = SESSION_FONT
                    mc.alignment = SESSION_ALIGN
            col += 1

    ws.column_dimensions["A"].width = 22
    for col_idx in range(2, len(date_list) + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = 4
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "B3"

    wb.save(output_path)


def _write_course_overview_sheet(
    output_path: str, courses: list[dict], sessions: list[dict]
) -> None:
    """Append a 'Course Overview' sheet with demand data and missed courses.

    Columns:
      Course Code | Course Name | Duration (Days) | Total Students |
      Additional Sessions Needed | New Sessions Planned | Sessions Unscheduled |
      Has Unscheduled Sessions | Mandatory per Quarter | Unscheduled Reason
    """

    HEADER_FILL = PatternFill(start_color="FF1F4E78", end_color="FF1F4E78", fill_type="solid")
    HEADER_FONT = Font(bold=True, color="FFFFFFFF")
    MISSED_FILL = PatternFill(start_color="FFFFCCCC", end_color="FFFFCCCC", fill_type="solid")
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

    # Count sessions actually planned per course code
    planned_counts: dict[str, int] = {}
    for s in sessions:
        code = str(s.get("course_code", "")).strip()
        if code:
            planned_counts[code] = planned_counts.get(code, 0) + 1

    # Build overview rows
    overview_rows = []
    for c in courses:
        code = str(c.get("course_code", "")).strip()
        name = str(c.get("course_title", "")).strip()
        duration = c.get("no_of_days")
        students = c.get("students_interested")
        needed = c.get("sessions_needed")
        planned = planned_counts.get(code, 0)
        missed = max(0, needed - planned) if needed is not None else 0
        missed_flag = "Yes" if missed > 0 else "No"

        overview_rows.append({
            "course_code": code,
            "course_name": name,
            "duration_days": duration,
            "total_students": students,
            "sessions_needed": needed,
            "sessions_planned": planned,
            "sessions_missed": missed,
            "missed_due_to_timeline": missed_flag,
            "mandatory_sessions": c.get("mandatory_sessions", 0),
            "unscheduled_reason": c.get("unscheduled_reason", ""),
        })

    wb = load_workbook(output_path)
    if "Course Overview" in wb.sheetnames:
        del wb["Course Overview"]
    ws = wb.create_sheet("Course Overview")

    headers = [
        "Course Code", "Course Name", "Duration (Days)", "Total Students",
        "Additional Sessions Needed", "New Sessions Planned", "Sessions Unscheduled",
        "Has Unscheduled Sessions", "Mandatory per Quarter", "Unscheduled Reason",
    ]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER

    for row_idx, row_data in enumerate(overview_rows, start=2):
        ws.cell(row=row_idx, column=1, value=row_data["course_code"])
        ws.cell(row=row_idx, column=2, value=row_data["course_name"])
        ws.cell(row=row_idx, column=3, value=row_data["duration_days"]).alignment = ALIGN_CENTER
        ws.cell(row=row_idx, column=4, value=row_data["total_students"]).alignment = ALIGN_CENTER
        ws.cell(row=row_idx, column=5, value=row_data["sessions_needed"]).alignment = ALIGN_CENTER
        ws.cell(row=row_idx, column=6, value=row_data["sessions_planned"]).alignment = ALIGN_CENTER
        missed_cell = ws.cell(row=row_idx, column=7, value=row_data["sessions_missed"])
        missed_cell.alignment = ALIGN_CENTER
        if row_data["sessions_missed"] > 0:
            missed_cell.fill = MISSED_FILL
        flag_cell = ws.cell(row=row_idx, column=8, value=row_data["missed_due_to_timeline"])
        flag_cell.alignment = ALIGN_CENTER
        if row_data["missed_due_to_timeline"] == "Yes":
            flag_cell.fill = MISSED_FILL
        ws.cell(row=row_idx, column=9, value=row_data["mandatory_sessions"]).alignment = ALIGN_CENTER
        ws.cell(row=row_idx, column=10, value=row_data["unscheduled_reason"])

    # Summary row
    total_row = len(overview_rows) + 3
    ws.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=total_row, column=4, value=sum(
        r["total_students"] or 0 for r in overview_rows
    )).font = Font(bold=True)
    ws.cell(row=total_row, column=5, value=sum(
        r["sessions_needed"] or 0 for r in overview_rows
    )).font = Font(bold=True)
    ws.cell(row=total_row, column=6, value=sum(
        r["sessions_planned"] for r in overview_rows
    )).font = Font(bold=True)
    ws.cell(row=total_row, column=7, value=sum(
        r["sessions_missed"] for r in overview_rows
    )).font = Font(bold=True)

    # Auto-size columns
    col_widths = [14, 40, 16, 14, 26, 22, 23, 25, 23, 42]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_path)


def _write_utilization_metrics_sheet(
    output_path: str,
    trainers: list[dict],
    sessions: list[dict],
    priority_trainer_names: set[str],
) -> None:
    """Append a 'Utilization Metrics' sheet with:
    - Trainer utilization: teaching days vs total available days per quarter (%)
    - Lab/Room utilization: booked days vs total available days (%)
    """
    HEADER_FILL = PatternFill(start_color="FF1F4E78", end_color="FF1F4E78", fill_type="solid")
    HEADER_FONT = Font(bold=True, color="FFFFFFFF")
    SECTION_FONT = Font(bold=True, size=11)
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
    PCT_FORMAT = "0.0%"

    # --- Date range ---
    plan_start, plan_end = PLANNING_START, PLANNING_END

    # Count total working days (Mon-Fri) in the planning period
    all_working_days: list[date] = []
    cur = plan_start
    while cur <= plan_end:
        if cur.weekday() < 5:
            all_working_days.append(cur)
        cur += timedelta(days=1)
    total_working_days = len(all_working_days)

    # --- Trainer Utilization ---
    leave_lookup = _build_leave_lookup(trainers)
    location_lookup = _load_trainer_locations()

    # Deduplicated trainer list
    trainer_names: list[str] = []
    seen: set[str] = set()
    for trainer in trainers:
        name = str(trainer["trainer"]).strip()
        aliases = _trainer_name_aliases(name)
        if not (aliases & seen):
            seen.update(aliases)
            trainer_names.append(name)
    for name in sorted(priority_trainer_names):
        aliases = _trainer_name_aliases(name)
        if not (aliases & seen):
            seen.update(aliases)
            trainer_names.append(name)
    for s in sessions:
        name = str(s.get("trainer_name", "")).strip()
        if name:
            aliases = _trainer_name_aliases(name)
            if not (aliases & seen):
                seen.update(aliases)
                trainer_names.append(name)

    # Build teaching days per trainer
    trainer_teaching_days: dict[str, set] = {}
    for s in sessions:
        norm = _normalize_trainer_name(s.get("trainer_name", ""))
        sd = pd.to_datetime(s.get("start_date"), errors="coerce")
        ed = pd.to_datetime(s.get("end_date"), errors="coerce")
        if pd.isna(sd) or pd.isna(ed):
            continue
        days_set = trainer_teaching_days.setdefault(norm, set())
        cur_d = sd.date()
        while cur_d <= ed.date():
            if PLANNING_START <= cur_d <= PLANNING_END and cur_d.weekday() < 5:
                days_set.add(cur_d)
            cur_d += timedelta(days=1)

    trainer_metrics = []
    for name in trainer_names:
        # Leave days
        leave_set: set[date] = set()
        for alias in _trainer_name_aliases(name):
            record = leave_lookup.get(alias)
            if record:
                for raw in record.get("leave_dates", []):
                    parsed = pd.to_datetime(raw, errors="coerce")
                    if not pd.isna(parsed) and parsed.date().weekday() < 5:
                        leave_set.add(parsed.date())

        available_days = total_working_days - len(leave_set)

        # Teaching days
        teaching_days = 0
        for alias in _trainer_name_aliases(name):
            teaching_days += len(trainer_teaching_days.get(alias, set()))

        pct = teaching_days / available_days if available_days > 0 else 0

        # Location
        loc = "Others"
        for alias in _trainer_name_aliases(name):
            if alias in location_lookup:
                loc = location_lookup[alias]["location"]
                break

        trainer_metrics.append({
            "trainer": name,
            "location": loc,
            "total_working_days": total_working_days,
            "leave_days": len(leave_set),
            "available_days": available_days,
            "teaching_days": teaching_days,
            "utilization_pct": pct,
        })

    bay_booked_days: dict[str, set] = {}
    for booking in _load_asset_bay_bookings():
        bay = f"Bay {booking['bay']}"
        days_set = bay_booked_days.setdefault(bay, set())
        cur_d = date.fromisoformat(booking["start_date"])
        end_d = date.fromisoformat(booking["end_date"])
        while cur_d <= end_d:
            if cur_d.weekday() < 5:
                days_set.add(cur_d)
            cur_d += timedelta(days=1)

    for session in sessions:
        bay = str(session.get("bay_room", ""))
        if not bay.startswith("Bay "):
            continue
        days_set = bay_booked_days.setdefault(bay, set())
        days_set.update(date.fromisoformat(day) for day in session["lab_dates"])

    bay_metrics = []
    for bay in sorted(bay_booked_days.keys()):
        occupied = len(bay_booked_days[bay])
        pct = occupied / total_working_days if total_working_days > 0 else 0
        bay_metrics.append({
            "bay": bay,
            "total_working_days": total_working_days,
            "occupied_days": occupied,
            "occupancy_pct": pct,
        })

    # --- Write to Excel ---
    wb = load_workbook(output_path)
    if "Utilization Metrics" in wb.sheetnames:
        del wb["Utilization Metrics"]
    ws = wb.create_sheet("Utilization Metrics")

    # Section 1: Trainer Utilization
    row = 1
    ws.cell(row=row, column=1, value="Trainer Utilization (Teaching vs Available Days)").font = SECTION_FONT
    row += 1

    trainer_headers = [
        "Trainer", "Location", "Total Working Days", "Leave Days",
        "Available Days", "Teaching Days", "Utilization %",
    ]
    for col_idx, header in enumerate(trainer_headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
    row += 1

    for tm in trainer_metrics:
        ws.cell(row=row, column=1, value=tm["trainer"])
        ws.cell(row=row, column=2, value=tm["location"]).alignment = ALIGN_CENTER
        ws.cell(row=row, column=3, value=tm["total_working_days"]).alignment = ALIGN_CENTER
        ws.cell(row=row, column=4, value=tm["leave_days"]).alignment = ALIGN_CENTER
        ws.cell(row=row, column=5, value=tm["available_days"]).alignment = ALIGN_CENTER
        ws.cell(row=row, column=6, value=tm["teaching_days"]).alignment = ALIGN_CENTER
        pct_cell = ws.cell(row=row, column=7, value=tm["utilization_pct"])
        pct_cell.number_format = PCT_FORMAT
        pct_cell.alignment = ALIGN_CENTER
        row += 1

    # Section 2: Cleveland Bay Utilization
    row += 2
    ws.cell(row=row, column=1, value="Cleveland Bay Utilization (Existing + Planned)").font = SECTION_FONT
    row += 1

    bay_headers = ["Bay Room", "Q1 Working Days", "Occupied Days", "Occupancy %"]
    for col_idx, header in enumerate(bay_headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
    row += 1

    for bay_metric in bay_metrics:
        ws.cell(row=row, column=1, value=bay_metric["bay"])
        ws.cell(row=row, column=2, value=bay_metric["total_working_days"]).alignment = ALIGN_CENTER
        ws.cell(row=row, column=3, value=bay_metric["occupied_days"]).alignment = ALIGN_CENTER
        pct_cell = ws.cell(row=row, column=4, value=bay_metric["occupancy_pct"])
        pct_cell.number_format = PCT_FORMAT
        pct_cell.alignment = ALIGN_CENTER
        row += 1

    # Column widths
    col_widths = [28, 12, 20, 12, 16, 16, 16]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_path)


def plan_courses(session_data: dict) -> dict:
    """Main planning function:
    1. Use pre-computed session requirements
    2. Load trainer availability and priority
    3. Use LLM to assign trainers and dates
    4. Save result to Excel
    """
    try:
        # Step 1: Use session requirements from caller
        courses = [dict(course) for course in session_data["courses"]]
        print(f"[PLANNER] Step 1: Got {len(courses)} courses")

        # Step 2: Load trainer data
        print("[PLANNER] Step 2: Loading trainer data...")
        trainers, leave_period = _load_trainer_leave_dates()
        priority_df = _load_trainer_priority()
        qualifications = _load_trainer_qualifications(trainers)
        bay_bookings = _load_asset_bay_bookings()
        roster_names = {_normalize_trainer_name(trainer["trainer"]) for trainer in trainers}
        qualified_codes = {
            entry["course_code"] for entry in qualifications
            if _normalize_trainer_name(entry["trainer_name"]) in roster_names
        }
        skipped_courses = []
        for course in courses:
            if course["sessions_needed"] and course["course_code"] not in qualified_codes:
                reason = "No qualified Cleveland roster trainer"
                course["unscheduled_reason"] = reason
                skipped_courses.append({
                    "course_code": course["course_code"],
                    "sessions_needed": course["sessions_needed"],
                    "mandatory_sessions": course["mandatory_sessions"],
                    "reason": reason,
                })
        if skipped_courses:
            print("[PLANNER] WARNING: No qualified Cleveland roster trainer for: "
                  + ", ".join(course["course_code"] for course in skipped_courses))
        schedulable_courses = [
            course for course in courses if not course.get("unscheduled_reason")
        ]
        course_bays: dict[str, set[str]] = {}
        for booking in bay_bookings:
            course_bays.setdefault(booking["course_code"], set()).add(booking["bay"])

        # Use the configured Q1 2027 planning period.
        period = f"{PLANNING_START} to {PLANNING_END}"
        print(f"[PLANNER] Loaded leave dates for {len(trainers)} trainers, planning period: {period}")

        # Step 3: Build prompt context
        print("[PLANNER] Step 3: Building prompt and invoking LLM...")
        priority_records = priority_df.to_dict(orient="records")
        leave_lookup = _build_leave_lookup(trainers)
        trainer_leave_summary = {}
        for trainer_name in _priority_trainer_names(priority_df):
            leave_record = leave_lookup.get(_normalize_trainer_name(trainer_name))
            trainer_leave_summary.setdefault(
                trainer_name,
                {
                    "trainer": trainer_name,
                    "leave_dates": leave_record["leave_dates"] if leave_record else [],
                },
            )
        for trainer in trainers:
            trainer_leave_summary.setdefault(trainer["trainer"], trainer)
        trainer_leave_records = sorted(
            trainer_leave_summary.values(), key=lambda trainer: trainer["trainer"]
        )

        rows = []
        if any(course["sessions_needed"] for course in schedulable_courses):
            from utils.llm_utils import get_llm_from_env

            prompt = get_course_planner_prompt()
            structured_llm = get_llm_from_env().with_structured_output(CoursePlan)
            chain = prompt | structured_llm
            invoke_input = {
                "period": period,
                "courses_json": json.dumps(schedulable_courses, indent=2),
                "priority_json": json.dumps(priority_records, indent=2, default=str),
                "qualification_json": json.dumps(qualifications, indent=2),
                "bay_occupancy_json": json.dumps(
                    {
                        "compatible_bays_by_course": {
                            code: sorted(bays) for code, bays in course_bays.items()
                        },
                        "existing_bookings": bay_bookings,
                    },
                    indent=2,
                ),
                "trainer_leave_json": json.dumps(trainer_leave_records, indent=2),
                "validation_feedback": "None",
            }
            formatted_messages = prompt.format_messages(**invoke_input)
            full_prompt_text = "\n".join(m.content for m in formatted_messages)
            print(f"[PLANNER] Estimated prompt tokens: {_count_tokens(full_prompt_text)}")
            start_time = time.time()
            result: CoursePlan = chain.invoke(invoke_input)
            print(f"[PLANNER] LLM returned {len(result.sessions)} sessions in {time.time() - start_time:.1f}s")
            rows = [session.model_dump() for session in result.sessions]
            try:
                _balance_trainer_assignments(rows, trainers, qualifications, priority_df)
                _validate_sessions(rows, schedulable_courses, trainers, qualifications)
                _assign_bay_rooms(rows, bay_bookings, schedulable_courses)
            except ValueError as error:
                print(f"[PLANNER] Proposed plan is invalid: {error}")
                print("[PLANNER] Repairing dates, trainers and bays with joint resource scheduling...")
                rows = _schedule_resource_feasible_sessions(
                    schedulable_courses, trainers, qualifications, bay_bookings
                )
                _balance_trainer_assignments(rows, trainers, qualifications, priority_df)
                _validate_sessions(rows, schedulable_courses, trainers, qualifications)
                _assign_bay_rooms(rows, bay_bookings, schedulable_courses)
                print(f"[PLANNER] Resource scheduling produced {len(rows)} validated sessions")
        else:
            _validate_sessions(rows, schedulable_courses, trainers, qualifications)

        # Step 4: Save to Excel
        print("[PLANNER] Step 4: Saving to Excel...")
        for row in rows:
            row["location"] = "Cleveland"
        result_df = pd.DataFrame(rows, columns=[
            "course_code", "course_name", "session_number", "start_date", "end_date",
            "trainer_name", "bay_room", "lab_dates", "location",
        ]).rename(
            columns={
                "course_code": "Course Code",
                "course_name": "Course Name",
                "session_number": "Session Number",
                "start_date": "Start Date",
                "end_date": "End Date",
                "trainer_name": "Trainer",
                "bay_room": "Bay Room",
                "lab_dates": "Lab Dates",
                "location": "Location",
            }
        )
        result_df["Lab Dates"] = result_df["Lab Dates"].apply(lambda dates: ", ".join(dates))

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        output_path = os.path.join(OUTPUT_DIR, "Cleveland_Course_Plan_Q1_2027.xlsx")
        result_df.to_excel(
            output_path, index=False, sheet_name="Cleveland Course Plan"
        )
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="a") as writer:
            pd.DataFrame(bay_bookings, columns=["bay", "course_code", "start_date", "end_date"]).to_excel(
                writer, index=False, sheet_name="Existing Bay Bookings"
            )
            requirements = pd.DataFrame(courses, columns=[
                "course_code", "course_title", "students_interested", "max_capacity",
                "no_of_days", "mandatory_sessions", "demand_sessions", "sessions_needed", "lab_days",
                "unscheduled_reason",
            ])
            requirements["lab_days"] = requirements["lab_days"].apply(
                lambda days: ", ".join(str(day) for day in days)
            )
            requirements.to_excel(writer, index=False, sheet_name="Session Requirements")
        print(f"[PLANNER] Saved to {output_path}")

        # Step 5: Append trainer availability visualization sheet
        try:
            print("[PLANNER] Step 5: Writing trainer availability sheet...")
            _write_trainer_availability_sheet(
                output_path=output_path,
                trainers=trainers,
                sessions=rows,
                priority_trainer_names=_priority_trainer_names(priority_df),
            )
            print("[PLANNER] Trainer availability sheet written")
        except Exception as availability_error:
            print(f"[PLANNER] WARNING: Could not write availability sheet: {availability_error}")
            print(traceback.format_exc())

        # Step 6: Append course overview sheet
        _write_course_overview_sheet(output_path=output_path, courses=courses, sessions=rows)
        _write_lab_availability_sheet(output_path=output_path, sessions=rows)

        # Step 7: Append utilization metrics sheet
        try:
            print("[PLANNER] Step 8: Writing utilization metrics sheet...")
            _write_utilization_metrics_sheet(
                output_path=output_path,
                trainers=trainers,
                sessions=rows,
                priority_trainer_names=_priority_trainer_names(priority_df),
            )
            print("[PLANNER] Utilization metrics sheet written")
        except Exception as util_error:
            print(f"[PLANNER] WARNING: Could not write utilization metrics sheet: {util_error}")
            print(traceback.format_exc())

        return {
            "total_sessions_planned": len(rows),
            "output_file": output_path,
            "plan": rows,
            "skipped_courses": skipped_courses,
        }
    except Exception as e:
        print(f"[PLANNER] ERROR: {e}")
        print(f"[PLANNER] Traceback: {traceback.format_exc()}")
        raise
