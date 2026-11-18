"""Read canonical quarterly tables once, before invoking any scheduling model."""

import math
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

from course_planner.models import (
    AvailabilityRow,
    BookingRow,
    CompatibilityRow,
    Location,
    PriorityRow,
    QualificationRow,
    Rules,
    TrainerRow,
    quarter_dates,
)

INPUT_FILES = (
    "courses.xlsx",
    "trainers.xlsx",
    "qualifications.xlsx",
    "priorities.xlsx",
    "availability.xlsx",
    "lab_compatibility.xlsx",
    "bookings.xlsx",
    "rules.json",
)
SITES = {"best": "BEST", "cleveland": "CL", "singapore": "SLC"}


def read_table(path: Path, columns: set[str]) -> pd.DataFrame:
    frame = pd.read_excel(
        path, dtype={column: str for column in ("lab", "trainer_id", "course_code", "trainer")}
    ).fillna("")
    missing = columns - set(frame.columns)
    if missing:
        raise ValueError(f"{path.name}: missing columns {sorted(missing)}")
    return frame


def read_demand(path: Path) -> pd.DataFrame:
    required = {"Comp Short Ttl", "Req Site Cd", "Local Course Code", "Course Title"}
    with pd.ExcelFile(path, engine="openpyxl") as workbook:
        for sheet in workbook.sheet_names:
            preview = pd.read_excel(workbook, sheet_name=sheet, header=None, nrows=30)
            for index, row in preview.iterrows():
                if required <= {str(v).strip() for v in row.dropna()}:
                    frame = pd.read_excel(workbook, sheet_name=sheet, header=index)
                    frame.columns = frame.columns.astype(str).str.strip()
                    return frame
    raise ValueError("Demand workbook has no SRL/WL table with required columns")


def whole_number(value: object, field: str, minimum: int = 0) -> int:
    try:
        number = float(str(value))
    except ValueError as error:
        raise ValueError(f"{field} must be a whole number >= {minimum}") from error
    if not math.isfinite(number) or not number.is_integer() or number < minimum:
        raise ValueError(f"{field} must be a whole number >= {minimum}")
    return int(number)


@dataclass
class PlanningData:
    location: str
    quarter: str
    start: date
    end: date
    input_dir: Path
    courses: list[dict]
    trainers: list[dict]
    qualifications: list[dict]
    priorities: pd.DataFrame
    bookings: list[dict]
    compatibility: dict[str, set[str]]
    rules: Rules
    warnings: list[str]
    filtered_demand: pd.DataFrame
    excluded_dates: set[date]
    parttime_dates: dict[str, set[date]]


def load_inputs(
    root: Path,
    location: Location,
    quarter: str,
    demand_path: Path | None = None,
    *,
    snapshot_dir: Path | None = None,
) -> PlanningData:
    start, end = quarter_dates(quarter)
    base = snapshot_dir if snapshot_dir is not None else root / "quarters" / quarter / location
    missing = [name for name in INPUT_FILES if not (base / name).is_file()]
    demand_path = demand_path or (
        base / "demand.xlsx" if snapshot_dir is not None else base.parent / "shared" / "demand.xlsx"
    )
    if not demand_path.is_file():
        missing.append("shared/demand.xlsx")
    if missing:
        raise ValueError(f"{location}/{quarter}: missing input files: {', '.join(missing)}")
    rules = Rules.model_validate_json((base / "rules.json").read_text(encoding="utf-8"))
    excluded = set(rules.excluded_dates)
    if rules.skip_first_week:
        excluded.update(start + timedelta(days=n) for n in range(7))
    if any(not start <= day <= end for day in excluded):
        raise ValueError("Excluded dates must belong to the selected quarter")
    warnings = list(rules.assumptions)
    roster = read_table(base / "trainers.xlsx", {"trainer_id", "trainer"})
    if roster["trainer_id"].duplicated().any() or roster["trainer"].duplicated().any():
        raise ValueError("Trainer IDs and display names must be unique")
    roster_rows = [TrainerRow.model_validate(record) for record in roster.to_dict(orient="records")]
    names = {row.trainer_id: row.trainer for row in roster_rows}
    if len(names) != len(roster_rows) or len(set(names.values())) != len(names):
        raise ValueError("Trainer IDs and display names must be unique after normalization")
    if any(not key or not value for key, value in names.items()):
        raise ValueError("Trainer IDs and names must not be blank")
    availability = read_table(
        base / "availability.xlsx", {"trainer_id", "date", "available_fraction"}
    )
    leave = {identity: {d.isoformat() for d in excluded} for identity in names}
    parttime: dict[str, set[date]] = {name: set() for name in names.values()}
    for record in availability.to_dict(orient="records"):
        row = AvailabilityRow.model_validate(record)
        identity = row.trainer_id
        if identity not in names:
            raise ValueError(f"Unknown availability trainer_id: {identity}")
        day = row.date
        fraction = row.available_fraction
        if not start <= day <= end:
            warnings.append(f"Availability date {day} is outside {quarter}; ignored")
        elif fraction < 1:
            # Sessions require a whole teaching day; partial days cannot host one.
            leave[identity].add(day.isoformat())
            if fraction > 0:
                parttime[names[identity]].add(day)
    trainers = [
        {"trainer_id": identity, "trainer": name, "leave_dates": sorted(leave[identity])}
        for identity, name in names.items()
    ]
    qualifications_frame = read_table(base / "qualifications.xlsx", {"course_code", "trainer_id"})
    qualifications = []
    for record in qualifications_frame.to_dict(orient="records"):
        row = QualificationRow.model_validate(record)
        if str(row.trainer_id) not in names:
            raise ValueError(f"Unknown qualification trainer_id: {row.trainer_id}")
        qualifications.append(
            {
                "course_code": str(row.course_code).strip(),
                "trainer_name": names[str(row.trainer_id)],
            }
        )
    priority_frame = read_table(base / "priorities.xlsx", {"course_code", "trainer_id", "rank"})
    priority_rows: dict[str, dict] = {}
    for record in priority_frame.to_dict(orient="records"):
        row = PriorityRow.model_validate(record)
        if str(row.trainer_id) not in names:
            raise ValueError(f"Unknown priority trainer_id: {row.trainer_id}")
        rank = whole_number(row.rank, "rank", 1)
        if rank > 5:
            raise ValueError("Priority rank must be 1 through 5")
        entry = priority_rows.setdefault(
            str(row.course_code),
            {"course_code": str(row.course_code), **{f"priority_{n}": [] for n in range(1, 6)}},
        )
        entry[f"priority_{rank}"].append(names[str(row.trainer_id)])
    priorities = pd.DataFrame(
        priority_rows.values(), columns=["course_code", *[f"priority_{n}" for n in range(1, 6)]]
    )
    compatibility_frame = read_table(base / "lab_compatibility.xlsx", {"course_code", "lab"})
    compatibility: dict[str, set[str]] = {}
    for record in compatibility_frame.to_dict(orient="records"):
        row = CompatibilityRow.model_validate(record)
        if not str(row.lab).strip() or not str(row.course_code).strip():
            raise ValueError("Lab compatibility must have a course and lab")
        compatibility.setdefault(str(row.course_code).strip(), set()).add(str(row.lab).strip())
    bookings_frame = read_table(
        base / "bookings.xlsx", {"course_code", "lab", "start_date", "end_date"}
    )
    bookings = []
    for record in bookings_frame.to_dict(orient="records"):
        row = BookingRow.model_validate(record)
        first, last = row.start_date, row.end_date
        bookings.append(
            {
                "course_code": str(row.course_code),
                "lab": str(row.lab),
                "start_date": first.isoformat(),
                "end_date": last.isoformat(),
            }
        )
    demand = read_demand(demand_path)
    filtered = demand[
        demand["Req Site Cd"].astype(str).str.strip().str.upper().eq(SITES[location])
        & demand["Comp Short Ttl"].astype(str).str.strip().str.upper().eq("MODALITY IXR")
    ].copy()
    if filtered.empty:
        warnings.append(
            f"No MODALITY IXR demand found for site {SITES[location]}; mandatory requirements still apply."
        )
    if filtered["Local Course Code"].isna().any():
        raise ValueError("Filtered demand contains blank course codes")
    filtered["Local Course Code"] = filtered["Local Course Code"].astype(str).str.strip()
    counts = filtered.groupby("Local Course Code").size().to_dict()
    schedule = read_table(
        base / "courses.xlsx",
        {
            "course_code",
            "course_title",
            "no_of_days",
            "max_capacity",
            "mandatory_sessions",
            "lab_days",
        },
    )
    if schedule["course_code"].duplicated().any():
        raise ValueError("Duplicate course codes in courses.xlsx")
    courses = []
    seen = set()
    for row in schedule.to_dict(orient="records"):
        code = str(row["course_code"]).strip()
        if not code:
            raise ValueError("Course code must not be blank")
        seen.add(code)
        mandatory = whole_number(row["mandatory_sessions"], f"{code} mandatory_sessions")
        students = int(counts.get(code, 0))
        course = {
            "course_code": code,
            "course_title": str(row["course_title"]),
            "students_interested": students,
            "mandatory_sessions": mandatory,
            "lab_days": [],
            "classroom": str(row.get("classroom", "")),
        }
        try:
            duration = whole_number(row["no_of_days"], f"{code} no_of_days", 1)
            capacity = whole_number(row["max_capacity"], f"{code} max_capacity", 1)
        except ValueError as error:
            if not (students or mandatory):
                warnings.append(str(error))
                continue
            course.update(
                no_of_days=None,
                max_capacity=None,
                demand_sessions=None,
                sessions_needed=mandatory or None,
                unscheduled_reason=str(error),
            )
        else:
            labs = [
                whole_number(v, f"{code} lab_days", 1)
                for v in str(row["lab_days"]).replace(".0", "").split(",")
                if v.strip()
            ]
            if len(labs) != len(set(labs)) or any(day > duration for day in labs):
                raise ValueError(f"{code}: invalid lab-day indexes")
            needed = math.ceil(students / capacity)
            course.update(
                no_of_days=duration,
                max_capacity=capacity,
                lab_days=labs,
                demand_sessions=needed,
                sessions_needed=max(needed, mandatory),
            )
            if labs and not compatibility.get(code):
                course["unscheduled_reason"] = "No compatible lab configured"
        if students or mandatory:
            courses.append(course)
    for code in sorted(set(counts) - seen):
        courses.append(
            {
                "course_code": code,
                "course_title": str(
                    filtered.loc[filtered["Local Course Code"] == code, "Course Title"].iloc[0]
                ),
                "students_interested": int(counts[code]),
                "mandatory_sessions": 0,
                "max_capacity": None,
                "no_of_days": None,
                "demand_sessions": None,
                "sessions_needed": None,
                "lab_days": [],
                "unscheduled_reason": "Missing course schedule",
            }
        )
    qualified = {q["course_code"] for q in qualifications}
    for course in courses:
        if course["course_code"] not in qualified:
            course.setdefault("unscheduled_reason", "No qualified roster trainer")
    return PlanningData(
        location.title(),
        quarter,
        start,
        end,
        base,
        courses,
        trainers,
        qualifications,
        priorities,
        bookings,
        compatibility,
        rules,
        list(dict.fromkeys(warnings)),
        filtered,
        excluded,
        parttime,
    )
