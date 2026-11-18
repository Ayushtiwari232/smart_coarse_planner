from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date, timedelta
from fractions import Fraction
from math import isnan

import pandas as pd


@dataclass(frozen=True)
class _ResourceSlot:
    days: tuple[date, ...]
    lab_days: frozenset[date]
    trainer: str
    lab: str | None


class SchedulingEngine:
    """Existing resource search and balancing rules, scoped to one run."""

    def __init__(self, start: date, end: date, compatibility: dict[str, set[str]] | None = None):
        self.start = start
        self.end = end
        self.compatibility = compatibility

    def _split_priority_trainers(self, value: object) -> list[str]:
        if isinstance(value, list):
            return [str(name).strip() for name in value if str(name).strip()]
        if value is None or value is pd.NA or isinstance(value, float) and isnan(value):
            return []
        if not isinstance(value, str):
            raise ValueError("Priority groups must contain trainer names")
        return [name.strip() for name in str(value).split(",") if name.strip()]

    def _priority_trainer_names(self, priority_df: pd.DataFrame) -> set[str]:
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
                trainer_names.update(self._split_priority_trainers(value))
        return trainer_names

    def _normalize_trainer_name(self, name: str) -> str:
        return " ".join(str(name).strip().lower().split())

    def _trainer_name_aliases(self, name: str) -> set[str]:
        aliases = {self._normalize_trainer_name(name)}
        if "," in str(name):
            last_name, first_names = [part.strip() for part in str(name).split(",", 1)]
            aliases.add(self._normalize_trainer_name(f"{first_names} {last_name}"))
            last_name_parts = last_name.split()
            if last_name_parts:
                aliases.add(self._normalize_trainer_name(f"{first_names} {last_name_parts[0]}"))
        return aliases

    def _build_leave_lookup(self, trainers: list[dict]) -> dict[str, dict]:
        leave_lookup = {}
        for trainer in trainers:
            for alias in self._trainer_name_aliases(trainer["trainer"]):
                leave_lookup[alias] = trainer
        return leave_lookup

    def _working_days(self, start: date, end: date) -> list[date]:
        days = []
        while start <= end:
            if start.weekday() < 5:
                days.append(start)
            start += timedelta(days=1)
        return days

    def _lab_resources(
        self, bookings: list[dict]
    ) -> tuple[dict[str, set[str]], dict[str, set[date]]]:
        course_labs: dict[str, set[str]] = {
            code: set(labs) for code, labs in (self.compatibility or {}).items()
        }
        reserved_by_lab: dict[str, set[date]] = {}
        for booking in bookings:
            lab = booking["lab"]
            if self.compatibility is None:
                course_labs.setdefault(booking["course_code"], set()).add(lab)
            reserved_by_lab.setdefault(lab, set()).update(
                self._working_days(
                    date.fromisoformat(booking["start_date"]),
                    date.fromisoformat(booking["end_date"]),
                )
            )
        return course_labs, reserved_by_lab

    def _schedule_resource_feasible_sessions(
        self,
        courses: list[dict],
        trainers: list[dict],
        qualifications: list[dict],
        bookings: list[dict],
    ) -> list[dict]:
        """Search dates, trainers and labs jointly without relaxing resource constraints."""
        required = {
            course["course_code"]: course for course in courses if course["sessions_needed"]
        }
        if not required:
            return []
        working_days = self._working_days(self.start, self.end)
        roster = {self._normalize_trainer_name(trainer["trainer"]): trainer for trainer in trainers}
        aliases = {
            alias: key
            for key, trainer in roster.items()
            for alias in self._trainer_name_aliases(trainer["trainer"])
        }
        leave = {
            key: {date.fromisoformat(day) for day in trainer["leave_dates"]}
            for key, trainer in roster.items()
        }
        available = {key: len(set(working_days) - leave[key]) for key in roster}
        qualified: dict[str, set[str]] = {}
        for entry in qualifications:
            key = aliases.get(self._normalize_trainer_name(entry["trainer_name"]))
            if key is not None:
                qualified.setdefault(entry["course_code"], set()).add(key)
        course_labs, reserved = self._lab_resources(bookings)
        slots: dict[str, list[_ResourceSlot]] = {}
        for code, course in required.items():
            duration = course["no_of_days"]
            choices = []
            for index in range(len(working_days) - duration + 1):
                days = tuple(working_days[index : index + duration])
                if duration in (3, 4) and days[0].isocalendar()[:2] != days[-1].isocalendar()[:2]:
                    continue
                lab_days = frozenset(days[day - 1] for day in course["lab_days"])
                labs: Sequence[str | None] = (
                    sorted(course_labs.get(code, set())) if lab_days else [None]
                )
                for trainer in sorted(qualified.get(code, set())):
                    if available[trainer] == 0 or not leave[trainer].isdisjoint(days):
                        continue
                    for lab in labs:
                        if lab is None or reserved.get(lab, set()).isdisjoint(lab_days):
                            choices.append(_ResourceSlot(days, lab_days, trainer, lab))
            if not choices:
                raise ValueError(
                    f"No feasible dates, qualified trainer and compatible lab for {code} "
                    f"within {self.start} to {self.end}; check leave and fixed lab bookings"
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
                    slot
                    for slot in slots[code]
                    if (previous_start is None or slot.days[0] > previous_start)
                    and busy[slot.trainer].isdisjoint(slot.days)
                    and (
                        slot.lab is None or reserved.get(slot.lab, set()).isdisjoint(slot.lab_days)
                    )
                ]
                if not options:
                    return False
                feasible[code] = options
            if not feasible:
                return True
            code = min(
                feasible,
                key=lambda value: (
                    len(feasible[value]),
                    -required[value]["no_of_days"],
                    value,
                ),
            )
            course = required[code]
            target_index = round(
                (len(working_days) - course["no_of_days"])
                * (len(assigned[code]) + 0.5)
                / course["sessions_needed"]
            )
            target = working_days[target_index]
            options = sorted(
                feasible[code],
                key=lambda slot: (
                    Fraction(loads[slot.trainer] + len(slot.days), available[slot.trainer]),
                    abs((slot.days[0] - target).days),
                    slot.days[0],
                    slot.trainer,
                    slot.lab or "",
                ),
            )
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
                if slot.lab is not None:
                    reserved.setdefault(slot.lab, set()).update(slot.lab_days)
                if search():
                    return True
                if slot.lab is not None:
                    reserved[slot.lab].difference_update(slot.lab_days)
                loads[slot.trainer] -= len(slot.days)
                busy[slot.trainer].difference_update(slot.days)
                assigned[code].pop()
            return False

        if not search():
            raise ValueError(
                "No complete resource-feasible location schedule exists for "
                + ", ".join(sorted(required))
                + "; check mandatory counts, qualifications, leave and fixed lab bookings"
            )
        rows = []
        for code, choices in assigned.items():
            for number, slot in enumerate(choices, start=1):
                rows.append(
                    {
                        "course_code": code,
                        "course_name": required[code]["course_title"],
                        "session_number": number,
                        "start_date": slot.days[0].isoformat(),
                        "end_date": slot.days[-1].isoformat(),
                        "trainer_name": roster[slot.trainer]["trainer"],
                        "lab_room": f"Lab {slot.lab}" if slot.lab is not None else "Not required",
                        "lab_dates": sorted(day.isoformat() for day in slot.lab_days),
                    }
                )
        return sorted(
            rows, key=lambda row: (row["start_date"], row["course_code"], row["session_number"])
        )

    def _assign_lab_rooms(
        self, sessions: list[dict], bookings: list[dict], courses: list[dict]
    ) -> None:
        """Allocate compatible labs for the actual lab days, or fail explicitly."""
        course_labs, reserved_by_lab = self._lab_resources(bookings)
        requirements = {course["course_code"]: course for course in courses}
        pending = []
        for session in sessions:
            course = requirements[session["course_code"]]
            lab_days = course["lab_days"]
            if not lab_days:
                session["lab_room"] = "Not required"
                session["lab_dates"] = []
                continue
            working_days = self._working_days(
                date.fromisoformat(session["start_date"]), date.fromisoformat(session["end_date"])
            )
            dates = {working_days[index - 1] for index in lab_days}
            candidates = sorted(course_labs.get(session["course_code"], set()))
            if not candidates:
                raise ValueError(f"No compatible lab is listed for {session['course_code']}")
            pending.append((session, dates, candidates))

        pending.sort(key=lambda item: len(item[2]))

        def assign(index: int) -> bool:
            if index == len(pending):
                return True
            session, dates, candidates = pending[index]
            for lab in candidates:
                reserved = reserved_by_lab.setdefault(lab, set())
                if reserved.isdisjoint(dates):
                    reserved.update(dates)
                    session["lab_room"] = f"Lab {lab}"
                    session["lab_dates"] = sorted(day.isoformat() for day in dates)
                    if assign(index + 1):
                        return True
                    reserved.difference_update(dates)
            return False

        if not assign(0):
            raise ValueError(
                "Cannot assign compatible labs without conflicting with existing or new lab bookings"
            )

    def _validate_sessions(
        self,
        sessions: list[dict],
        courses: list[dict],
        trainers: list[dict],
        qualifications: list[dict],
    ) -> None:
        requirements = {course["course_code"]: course for course in courses}
        leave = self._build_leave_lookup(trainers)
        qualified = {
            (q["course_code"], self._normalize_trainer_name(q["trainer_name"]))
            for q in qualifications
        }
        occupied: dict[str, set[date]] = {}
        counts: dict[str, int] = {}
        session_ids = set()
        for session in sessions:
            code = session["course_code"].strip()
            session["course_code"] = code
            course = requirements.get(code)
            if (
                course is None
                or course["sessions_needed"] is None
                or course["sessions_needed"] <= 0
            ):
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
            days = self._working_days(start, end)
            if (
                start < self.start
                or end > self.end
                or start > end
                or start.weekday() >= 5
                or end.weekday() >= 5
                or len(days) != course["no_of_days"]
            ):
                raise ValueError(
                    f"Invalid planning period or duration for {code}: {start} to {end}"
                )
            if len(days) in (3, 4) and start.isocalendar()[:2] != end.isocalendar()[:2]:
                raise ValueError(f"Session for {code} cannot cross a weekend")
            name = session["trainer_name"]
            trainer = next(
                (leave[alias] for alias in self._trainer_name_aliases(name) if alias in leave), None
            )
            if (
                trainer is None
                or (code, self._normalize_trainer_name(trainer["trainer"])) not in qualified
            ):
                raise ValueError(f"Trainer {name} is not a qualified roster trainer for {code}")
            trainer_key = self._normalize_trainer_name(trainer["trainer"])
            busy = occupied.setdefault(trainer_key, set())
            if not busy.isdisjoint(days) or any(
                day.isoformat() in trainer["leave_dates"] for day in days
            ):
                raise ValueError(f"Trainer {name} is unavailable for {code}")
            busy.update(days)
            counts[code] = counts.get(code, 0) + 1
            if counts[code] > course["sessions_needed"]:
                raise ValueError(f"Too many new sessions for {code}")
        for course in courses:
            missing = course["mandatory_sessions"] - counts.get(course["course_code"], 0)
            if missing > 0:
                raise ValueError(
                    f"Mandatory minimum not met for {course['course_code']}: {missing} session(s) missing"
                )
            if counts.get(course["course_code"], 0) != course["sessions_needed"]:
                raise ValueError(f"Required session count not met for {course['course_code']}")

    def _balance_trainer_assignments(
        self,
        sessions: list[dict],
        trainers: list[dict],
        qualifications: list[dict],
        priority_df: pd.DataFrame,
    ) -> None:
        """Find feasible assignments, then reduce utilization variance by moves and swaps."""
        if not sessions:
            return
        roster = {self._normalize_trainer_name(trainer["trainer"]): trainer for trainer in trainers}
        aliases = {
            alias: key
            for key, trainer in roster.items()
            for alias in self._trainer_name_aliases(trainer["trainer"])
        }
        working_days = set(self._working_days(self.start, self.end))
        leave = {
            key: {date.fromisoformat(day) for day in trainer["leave_dates"]} & working_days
            for key, trainer in roster.items()
        }
        available = {key: len(working_days - leave[key]) for key in roster}
        qualified: dict[str, set[str]] = {}
        for entry in qualifications:
            key = aliases.get(self._normalize_trainer_name(entry["trainer_name"]))
            if key is not None:
                qualified.setdefault(entry["course_code"], set()).add(key)
        ranks: dict[tuple[str, str], int] = {}
        for _, row in priority_df.iterrows():
            for rank in range(1, 6):
                for name in self._split_priority_trainers(row[f"priority_{rank}"]):
                    key = aliases.get(self._normalize_trainer_name(name))
                    if key is not None:
                        pair = (str(row["course_code"]), key)
                        ranks[pair] = min(rank, ranks.get(pair, rank))

        session_days = []
        candidates = []
        for session in sessions:
            start = date.fromisoformat(session["start_date"])
            end = date.fromisoformat(session["end_date"])
            if start > end or start < self.start or end > self.end:
                raise ValueError(
                    f"Invalid planning period for {session['course_code']}: {start} to {end}"
                )
            days = set(self._working_days(start, end))
            session_days.append(days)
            eligible = sorted(
                key
                for key in qualified.get(session["course_code"].strip(), set())
                if available[key] > 0 and leave[key].isdisjoint(days)
            )
            if not eligible:
                raise ValueError(
                    f"No qualified trainer is available for {session['course_code']} on {start} to {end}"
                )
            candidates.append(eligible)

        assignments: list[str | None] = [None] * len(sessions)
        loads = {key: 0 for key in roster}

        def can_assign(index: int, key: str, ignored: tuple[int, ...] = ()) -> bool:
            return all(
                other == index
                or other in ignored
                or assigned != key
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
            index = min(
                remaining, key=lambda item: (len(options[item]), -len(session_days[item]), item)
            )
            duration = len(session_days[index])
            for key in sorted(
                options[index],
                key=lambda name: (
                    Fraction(loads[name] + duration, available[name]),
                    ranks.get((sessions[index]["course_code"], name), 6),
                    name,
                ),
            ):
                assignments[index] = key
                loads[key] += duration
                if find_assignment():
                    return True
                loads[key] -= duration
                assignments[index] = None
            return False

        if not find_assignment():
            raise ValueError(
                "No conflict-free qualified trainer assignment exists for the proposed session dates"
            )

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
                if current is None:
                    raise ValueError("Missing assignment during workload balancing")
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
                    if other_key is None:
                        raise ValueError("Missing assignment during workload balancing")
                    if (
                        other_key == current
                        or other_key not in candidates[index]
                        or current not in candidates[other]
                    ):
                        continue
                    if not can_assign(index, other_key, (other,)) or not can_assign(
                        other, current, (index,)
                    ):
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
