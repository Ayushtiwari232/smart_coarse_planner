import unittest
from datetime import date

import pandas as pd

from course_planner.engine import SchedulingEngine

engine = SchedulingEngine(date(2027, 1, 1), date(2027, 3, 31))
_assign_lab_rooms = engine._assign_lab_rooms
_balance_trainer_assignments = engine._balance_trainer_assignments
_validate_sessions = engine._validate_sessions
_schedule_resource_feasible_sessions = engine._schedule_resource_feasible_sessions
_working_days = engine._working_days


def course(code="COURSE", **overrides):
    result = {
        "course_code": code,
        "course_title": code,
        "no_of_days": 3,
        "max_capacity": 5,
        "mandatory_sessions": 1,
        "existing_sessions": 0,
        "sessions_needed": 1,
        "lab_days": [3],
    }
    result.update(overrides)
    return result


def session(code="COURSE", start="2027-01-04", end="2027-01-06", number=1):
    return {
        "course_code": code,
        "course_name": code,
        "session_number": number,
        "start_date": start,
        "end_date": end,
        "trainer_name": "A Trainer",
    }


def booking(code="COURSE", lab="1", start="2027-01-04", end="2027-01-05"):
    return {"course_code": code, "lab": lab, "start_date": start, "end_date": end}


class SchedulingRegressionTests(unittest.TestCase):
    def test_occupied_preferred_trainer_is_replaced_by_lower_priority(self):
        trainers = [{"trainer": name, "leave_dates": []} for name in ("A Trainer", "B Trainer")]
        qualifications = [
            {"trainer_name": "A Trainer", "course_code": "COURSE"},
            {"trainer_name": "B Trainer", "course_code": "COURSE"},
            {"trainer_name": "A Trainer", "course_code": "RESTRICTED"},
        ]
        priorities = pd.DataFrame(
            [
                {
                    "course_code": "COURSE",
                    "priority_1": "A Trainer",
                    "priority_2": None,
                    "priority_3": None,
                    "priority_4": None,
                    "priority_5": None,
                }
            ]
        )
        rows = [session(), session("RESTRICTED")]
        _balance_trainer_assignments(rows, trainers, qualifications, priorities)
        self.assertEqual([row["trainer_name"] for row in rows], ["B Trainer", "A Trainer"])
        _validate_sessions(rows, [course(), course("RESTRICTED")], trainers, qualifications)

    def test_leave_and_qualifications_override_priority_and_balance(self):
        trainers = [
            {"trainer": "A Trainer", "leave_dates": ["2027-01-04"]},
            {"trainer": "B Trainer", "leave_dates": []},
            {"trainer": "Unqualified Trainer", "leave_dates": []},
        ]
        qualifications = [
            {"trainer_name": name, "course_code": "COURSE"} for name in ("A Trainer", "B Trainer")
        ]
        priorities = pd.DataFrame(
            columns=[
                "course_code",
                "priority_1",
                "priority_2",
                "priority_3",
                "priority_4",
                "priority_5",
            ]
        )
        rows = [session()]
        _balance_trainer_assignments(rows, trainers, qualifications, priorities)
        self.assertEqual(rows[0]["trainer_name"], "B Trainer")
        with self.assertRaisesRegex(ValueError, "No qualified trainer is available"):
            _balance_trainer_assignments([session()], trainers[:1], qualifications, priorities)

    def test_equal_priority_group_balances_all_trainers_and_preserves_dates(self):
        names = ["A Trainer", "B Trainer", "C Trainer"]
        trainers = [{"trainer": name, "leave_dates": []} for name in names]
        qualifications = [{"trainer_name": name, "course_code": "COURSE"} for name in names]
        priorities = pd.DataFrame(
            [
                {
                    "course_code": "COURSE",
                    "priority_1": "A Trainer, B Trainer, C Trainer",
                    "priority_2": None,
                    "priority_3": None,
                    "priority_4": None,
                    "priority_5": None,
                }
            ]
        )
        days = _working_days(date(2027, 1, 4), date(2027, 1, 11))
        rows = [
            session(start=day.isoformat(), end=day.isoformat(), number=index + 1)
            for index, day in enumerate(days)
        ]
        _balance_trainer_assignments(rows, trainers, qualifications, priorities)
        self.assertEqual(
            [sum(row["trainer_name"] == name for row in rows) for name in names], [2, 2, 2]
        )
        self.assertEqual([row["start_date"] for row in rows], [day.isoformat() for day in days])

    def test_priority_breaks_equal_utilization_ties(self):
        trainers = [{"trainer": name, "leave_dates": []} for name in ("A Trainer", "B Trainer")]
        qualifications = [
            {"trainer_name": trainer["trainer"], "course_code": "COURSE"} for trainer in trainers
        ]
        priorities = pd.DataFrame(
            [
                {
                    "course_code": "COURSE",
                    "priority_1": "B Trainer",
                    "priority_2": "A Trainer",
                    "priority_3": None,
                    "priority_4": None,
                    "priority_5": None,
                }
            ]
        )
        rows = [session()]
        _balance_trainer_assignments(rows, trainers, qualifications, priorities)
        self.assertEqual(rows[0]["trainer_name"], "B Trainer")

    def test_impossible_overlap_reports_error_instead_of_double_booking(self):
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        qualifications = [{"trainer_name": "A Trainer", "course_code": "COURSE"}]
        priorities = pd.DataFrame()
        with self.assertRaisesRegex(ValueError, "No conflict-free"):
            _balance_trainer_assignments(
                [session(), session(number=2)], trainers, qualifications, priorities
            )

    def test_partial_lab_can_use_lab_booked_on_other_days(self):
        rows = [session()]
        _assign_lab_rooms(rows, [booking()], [course()])
        self.assertEqual(rows[0]["lab_room"], "Lab 1")
        self.assertEqual(rows[0]["lab_dates"], ["2027-01-06"])

    def test_conflict_and_missing_lab_are_errors(self):
        with self.assertRaisesRegex(ValueError, "Cannot assign"):
            _assign_lab_rooms([session()], [booking(end="2027-01-06")], [course()])
        with self.assertRaisesRegex(ValueError, "No compatible lab"):
            _assign_lab_rooms([session()], [], [course()])

    def test_resource_search_moves_course_off_fixed_lab_bookings(self):
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        qualifications = [{"trainer_name": "A Trainer", "course_code": "COURSE"}]
        fixed = [booking(start="2027-01-01", end="2027-03-26")]
        requirements = [course()]
        with self.assertRaisesRegex(ValueError, "Cannot assign"):
            _assign_lab_rooms([session()], fixed, requirements)
        rows = _schedule_resource_feasible_sessions(requirements, trainers, qualifications, fixed)
        _validate_sessions(rows, requirements, trainers, qualifications)
        _assign_lab_rooms(rows, fixed, requirements)
        self.assertEqual(len(rows), 1)
        self.assertGreater(rows[0]["lab_dates"][0], "2027-03-26")
        self.assertLessEqual(rows[0]["end_date"], "2027-03-31")

    def test_joint_search_handles_shared_labs_trainer_leave_and_weekends(self):
        trainers = [
            {"trainer": "A Trainer", "leave_dates": ["2027-01-04"]},
            {"trainer": "B Trainer", "leave_dates": []},
        ]
        requirements = [
            course(
                "FIRST",
                no_of_days=4,
                lab_days=[1, 2, 3, 4],
                sessions_needed=2,
                mandatory_sessions=2,
            ),
            course("SECOND", no_of_days=3, lab_days=[3], sessions_needed=2, mandatory_sessions=2),
        ]
        qualifications = [
            {"trainer_name": trainer["trainer"], "course_code": item["course_code"]}
            for trainer in trainers
            for item in requirements
        ]
        fixed = [
            booking("FIRST", start="2027-01-01", end="2027-01-08"),
            booking("SECOND", start="2027-01-01", end="2027-01-08"),
        ]
        rows = _schedule_resource_feasible_sessions(requirements, trainers, qualifications, fixed)
        self.assertEqual(len(rows), 4)
        _validate_sessions(rows, requirements, trainers, qualifications)
        _assign_lab_rooms(rows, fixed, requirements)
        lab_dates = [day for row in rows for day in row["lab_dates"]]
        self.assertEqual(len(lab_dates), len(set(lab_dates)))

    def test_joint_search_reports_impossible_resource_requirements(self):
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        qualifications = [{"trainer_name": "A Trainer", "course_code": "COURSE"}]
        with self.assertRaisesRegex(ValueError, "No feasible dates.*COURSE"):
            _schedule_resource_feasible_sessions(
                [course()],
                trainers,
                qualifications,
                [booking(start="2027-01-01", end="2027-03-31")],
            )
        with self.assertRaisesRegex(ValueError, "No feasible dates.*COURSE"):
            _schedule_resource_feasible_sessions([course()], trainers, qualifications, [])

    def test_no_lab_needs_no_lab(self):
        rows = [session()]
        _assign_lab_rooms(rows, [], [course(lab_days=[])])
        self.assertEqual(rows[0]["lab_room"], "Not required")

    def test_mandatory_and_trainer_validation(self):
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        qualifications = [{"trainer_name": "A Trainer", "course_code": "COURSE"}]
        _validate_sessions([session()], [course()], trainers, qualifications)
        with self.assertRaisesRegex(ValueError, "Mandatory minimum"):
            _validate_sessions([], [course(existing_sessions=10)], trainers, qualifications)
        with self.assertRaisesRegex(ValueError, "unavailable"):
            _validate_sessions(
                [session()],
                [course()],
                [{"trainer": "A Trainer", "leave_dates": ["2027-01-06"]}],
                qualifications,
            )
        with self.assertRaisesRegex(ValueError, "qualified"):
            _validate_sessions([session()], [course()], trainers, [])
