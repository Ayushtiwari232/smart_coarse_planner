import base64
import os
import subprocess
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from types import ModuleType, SimpleNamespace
from unittest.mock import patch

import pandas as pd
from openpyxl import load_workbook

import main as fastapi_main
import function_app
from helpers.plan_helpers import start_plan_job
from services import cleveland_data, filter_service, planner_service
from services.cleveland_data import (
    _lab_days,
    _mandatory_count,
    count_existing_sessions,
    load_bay_bookings,
    load_course_schedule,
)
from services.planner_service import (
    CoursePlan, CourseSession, _assign_bay_rooms, _balance_trainer_assignments,
    _load_trainer_priority, _priority_trainer_names, _validate_sessions,
    _schedule_resource_feasible_sessions, _working_days, plan_courses,
)
from services.session_calculator_service import calculate_sessions


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


def booking(code="COURSE", bay="1", start="2027-01-04", end="2027-01-05"):
    return {"course_code": code, "bay": bay, "start_date": start, "end_date": end}


class ClevelandFlowTests(unittest.TestCase):
    def test_comma_separated_priority_cells_are_individual_trainers(self):
        source = pd.DataFrame([[
            "COURSE", "A Course", " A Trainer, B Trainer , ,", "C Trainer",
            None, None, None,
        ]])
        with patch("services.planner_service.pd.read_excel", return_value=source):
            priorities = _load_trainer_priority()
        self.assertEqual(priorities.iloc[0]["priority_1"], ["A Trainer", "B Trainer"])
        self.assertEqual(priorities.iloc[0]["priority_3"], [])
        self.assertEqual(_priority_trainer_names(priorities), {"A Trainer", "B Trainer", "C Trainer"})
        self.assertNotIn("A Trainer, B Trainer", priorities.to_json())

    def test_balancing_utilization_percentages_not_teaching_day_counts(self):
        quarter_days = _working_days(cleveland_data.PLANNING_START, cleveland_data.PLANNING_END)
        self.assertEqual(len(quarter_days), 64)
        trainers = [
            {"trainer": "A Trainer", "leave_dates": []},
            {"trainer": "B Trainer", "leave_dates": [day.isoformat() for day in quarter_days[32:]]},
        ]
        qualifications = [{"trainer_name": trainer["trainer"], "course_code": "COURSE"} for trainer in trainers]
        priorities = pd.DataFrame([{
            "course_code": "COURSE", "priority_1": "A Trainer", "priority_2": "B Trainer",
            "priority_3": None, "priority_4": None, "priority_5": None,
        }])
        rows = [session(start=day.isoformat(), end=day.isoformat(), number=index + 1)
                for index, day in enumerate(quarter_days[1:7])]
        _balance_trainer_assignments(rows, trainers, qualifications, priorities)
        counts = {name: sum(row["trainer_name"] == name for row in rows) for name in ("A Trainer", "B Trainer")}
        self.assertEqual(counts, {"A Trainer": 4, "B Trainer": 2})
        self.assertEqual(counts["A Trainer"] / 64, counts["B Trainer"] / 32)
        _validate_sessions(rows, [course(no_of_days=1, lab_days=[], sessions_needed=6)], trainers, qualifications)
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "plan.xlsx"
            pd.DataFrame().to_excel(output, index=False)
            with patch("services.planner_service._load_trainer_locations", return_value={}), \
                 patch("services.planner_service._load_asset_bay_bookings", return_value=[]):
                planner_service._write_utilization_metrics_sheet(
                    str(output), trainers, rows, {"A Trainer", "B Trainer"}
                )
            workbook = load_workbook(output, data_only=True)
            sheet = workbook["Utilization Metrics"]
            ratios = {
                row[0]: row[6] for row in sheet.iter_rows(values_only=True)
                if row[0] in {"A Trainer", "B Trainer"}
            }
            self.assertEqual(ratios, {"A Trainer": 0.0625, "B Trainer": 0.0625})
            workbook.close()

    def test_occupied_preferred_trainer_is_replaced_by_lower_priority(self):
        trainers = [{"trainer": name, "leave_dates": []} for name in ("A Trainer", "B Trainer")]
        qualifications = [
            {"trainer_name": "A Trainer", "course_code": "COURSE"},
            {"trainer_name": "B Trainer", "course_code": "COURSE"},
            {"trainer_name": "A Trainer", "course_code": "RESTRICTED"},
        ]
        priorities = pd.DataFrame([{
            "course_code": "COURSE", "priority_1": "A Trainer",
            "priority_2": None, "priority_3": None, "priority_4": None, "priority_5": None,
        }])
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
        qualifications = [{"trainer_name": name, "course_code": "COURSE"}
                          for name in ("A Trainer", "B Trainer")]
        priorities = pd.DataFrame(columns=["course_code", "priority_1", "priority_2", "priority_3", "priority_4", "priority_5"])
        rows = [session()]
        _balance_trainer_assignments(rows, trainers, qualifications, priorities)
        self.assertEqual(rows[0]["trainer_name"], "B Trainer")
        with self.assertRaisesRegex(ValueError, "No qualified trainer is available"):
            _balance_trainer_assignments([session()], trainers[:1], qualifications, priorities)

    def test_equal_priority_group_balances_all_trainers_and_preserves_dates(self):
        names = ["A Trainer", "B Trainer", "C Trainer"]
        trainers = [{"trainer": name, "leave_dates": []} for name in names]
        qualifications = [{"trainer_name": name, "course_code": "COURSE"} for name in names]
        priorities = pd.DataFrame([{
            "course_code": "COURSE", "priority_1": "A Trainer, B Trainer, C Trainer",
            "priority_2": None, "priority_3": None, "priority_4": None, "priority_5": None,
        }])
        days = _working_days(date(2027, 1, 4), date(2027, 1, 11))
        rows = [session(start=day.isoformat(), end=day.isoformat(), number=index + 1)
                for index, day in enumerate(days)]
        _balance_trainer_assignments(rows, trainers, qualifications, priorities)
        self.assertEqual([sum(row["trainer_name"] == name for row in rows) for name in names], [2, 2, 2])
        self.assertEqual([row["start_date"] for row in rows], [day.isoformat() for day in days])

    def test_priority_breaks_equal_utilization_ties(self):
        trainers = [{"trainer": name, "leave_dates": []} for name in ("A Trainer", "B Trainer")]
        qualifications = [{"trainer_name": trainer["trainer"], "course_code": "COURSE"} for trainer in trainers]
        priorities = pd.DataFrame([{
            "course_code": "COURSE", "priority_1": "B Trainer", "priority_2": "A Trainer",
            "priority_3": None, "priority_4": None, "priority_5": None,
        }])
        rows = [session()]
        _balance_trainer_assignments(rows, trainers, qualifications, priorities)
        self.assertEqual(rows[0]["trainer_name"], "B Trainer")

    def test_impossible_overlap_reports_error_instead_of_double_booking(self):
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        qualifications = [{"trainer_name": "A Trainer", "course_code": "COURSE"}]
        priorities = pd.DataFrame()
        with self.assertRaisesRegex(ValueError, "No conflict-free"):
            _balance_trainer_assignments([session(), session(number=2)], trainers, qualifications, priorities)

    def test_all_default_input_paths_use_same_folder(self):
        folder = cleveland_data.INPUT_DIR
        self.assertEqual(cleveland_data.SCHEDULE_FILE.parent, folder)
        self.assertEqual(cleveland_data.ASSET_BAY_FILE.parent, folder)
        self.assertEqual(planner_service.PRIORITY_FILE.parent, folder)
        self.assertEqual(planner_service.QUALIFICATION_FILE.parent, folder)
        self.assertEqual(fastapi_main.DEFAULT_INPUT_FILE.parent, folder)
        self.assertEqual(function_app.DEFAULT_INPUT_FILE.parent, folder)
        self.assertEqual(filter_service.INPUT_DIR, folder)
        for name in (
            planner_service.TRAINER_LEAVE_FILE,
            planner_service.TRAINERS_JSON_FILE,
            fastapi_main.DEFAULT_INPUT_FILE.name,
            function_app.DEFAULT_INPUT_FILE.name,
        ):
            self.assertTrue((folder / name).is_file(), name)

    def test_input_folder_override_applies_to_all_entry_points(self):
        with tempfile.TemporaryDirectory() as temp:
            env = os.environ.copy()
            env["CLEVELAND_INPUT_DIR"] = temp
            env.pop("CLEVELAND_SCHEDULE_FILE", None)
            script = (
                "import main, function_app; "
                "from services import cleveland_data, filter_service, planner_service; "
                "folder = cleveland_data.INPUT_DIR; "
                "assert main.DEFAULT_INPUT_FILE.parent == folder; "
                "assert function_app.DEFAULT_INPUT_FILE.parent == folder; "
                "assert filter_service.INPUT_DIR == folder; "
                "assert planner_service.INPUT_DIR == folder; "
                "assert planner_service.PRIORITY_FILE.parent == folder; "
                "assert planner_service.QUALIFICATION_FILE.parent == folder; "
                "assert cleveland_data.SCHEDULE_FILE.parent == folder; "
                "assert cleveland_data.ASSET_BAY_FILE.parent == folder"
            )
            result = subprocess.run(
                [sys.executable, "-c", script], cwd=Path(__file__).resolve().parents[1],
                env=env, capture_output=True, text=True, timeout=60,
            )
            self.assertEqual(result.returncode, 0, result.stderr)

    def test_fastapi_completed_job_endpoints(self):
        response = {
            "status": "success", "job_id": "test-job", "file_name": "plan.xlsx",
            "file_content_base64": base64.b64encode(b"workbook").decode("ascii"),
        }
        with patch("main.start_plan_job", return_value=response):
            result = fastapi_main.plan(fastapi_main.PlanRequest(
                file_name="input.xlsx", file_content_base64="encoded"
            ))
        self.assertEqual(result["job_id"], "test-job")
        status = fastapi_main.get_plan_result("test-job")
        self.assertNotIn("file_content_base64", status)
        download = fastapi_main.get_plan_file("test-job")
        self.assertIn("plan.xlsx", download.headers["content-disposition"])
        fastapi_main._results.clear()

    def test_real_input_schema_and_values(self):
        bookings = load_bay_bookings()
        schedule = load_course_schedule(bookings)
        self.assertEqual(schedule["IGT2BL001"]["mandatory_sessions"], 3)
        self.assertEqual(schedule["IGT2BL001"]["lab_days"], [11, 12, 13, 14, 15])
        self.assertEqual(schedule["IGT2VC012"]["lab_days"], [3])
        self.assertEqual(schedule["IGT2VC016"]["lab_days"], [])

    def test_invalid_values_are_reported(self):
        with self.assertRaises(ValueError):
            _mandatory_count("sometimes", "COURSE")
        for value in (0, 1, 2, 3.0, "2"):
            self.assertEqual(_mandatory_count(value, "COURSE"), int(value))
        for value in (-1, 1.5, "once per qtr"):
            with self.assertRaises(ValueError):
                _mandatory_count(value, "COURSE")
        with self.assertRaises(ValueError):
            _lab_days("unknown", 5, "COURSE", True)
        self.assertEqual(_lab_days(None, 3, "COURSE", True), [1, 2, 3])
        self.assertEqual(_lab_days(None, 3, "COURSE", False), [])

    def test_multiple_bays_count_once_and_outside_quarter_excluded(self):
        bookings = [
            booking(bay="1"), booking(bay="2"),
            booking(bay="3", start="2026-12-30", end="2027-01-04"),
            booking(bay="1", start="2027-02-01", end="2027-02-03"),
        ]
        self.assertEqual(
            count_existing_sessions(bookings, date(2027, 1, 1), date(2027, 3, 31)),
            {"COURSE": 2},
        )

    def test_mandatory_without_demand_and_bay_bookings_do_not_reduce_it(self):
        schedule = {
            "COURSE": {k: v for k, v in course(mandatory_sessions=2).items()
                       if k not in {"course_code", "existing_sessions", "sessions_needed"}},
            "OTHER": {k: v for k, v in course("OTHER", mandatory_sessions=0).items()
                      if k not in {"course_code", "existing_sessions", "sessions_needed"}},
        }
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "filtered.xlsx"
            pd.DataFrame([
                {"Local Course Code": "OTHER", "Course Title": "OTHER"}
                for _ in range(6)
            ]).to_excel(path, index=False)
            with patch("services.session_calculator_service.OUTPUT_DIR", temp), \
                 patch("services.session_calculator_service.load_bay_bookings", return_value=[booking()]), \
                 patch("services.session_calculator_service.load_course_schedule", return_value=schedule):
                result = calculate_sessions(str(path))
        courses = {item["course_code"]: item for item in result["courses"]}
        self.assertEqual(courses["COURSE"]["students_interested"], 0)
        self.assertEqual(courses["COURSE"]["sessions_needed"], 2)
        self.assertEqual(courses["OTHER"]["sessions_needed"], 2)

    def test_partial_lab_can_use_bay_booked_on_other_days(self):
        rows = [session()]
        _assign_bay_rooms(rows, [booking()], [course()])
        self.assertEqual(rows[0]["bay_room"], "Bay 1")
        self.assertEqual(rows[0]["lab_dates"], ["2027-01-06"])

    def test_conflict_and_missing_bay_are_errors(self):
        with self.assertRaisesRegex(ValueError, "Cannot assign"):
            _assign_bay_rooms([session()], [booking(end="2027-01-06")], [course()])
        with self.assertRaisesRegex(ValueError, "No compatible bay"):
            _assign_bay_rooms([session()], [], [course()])

    def test_resource_search_moves_course_off_fixed_bay_bookings(self):
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        qualifications = [{"trainer_name": "A Trainer", "course_code": "COURSE"}]
        fixed = [booking(start="2027-01-01", end="2027-03-26")]
        requirements = [course()]
        with self.assertRaisesRegex(ValueError, "Cannot assign"):
            _assign_bay_rooms([session()], fixed, requirements)
        rows = _schedule_resource_feasible_sessions(requirements, trainers, qualifications, fixed)
        _validate_sessions(rows, requirements, trainers, qualifications)
        _assign_bay_rooms(rows, fixed, requirements)
        self.assertEqual(len(rows), 1)
        self.assertGreater(rows[0]["lab_dates"][0], "2027-03-26")
        self.assertLessEqual(rows[0]["end_date"], "2027-03-31")

    def test_joint_search_handles_shared_bays_trainer_leave_and_weekends(self):
        trainers = [
            {"trainer": "A Trainer", "leave_dates": ["2027-01-04"]},
            {"trainer": "B Trainer", "leave_dates": []},
        ]
        requirements = [
            course("FIRST", no_of_days=4, lab_days=[1, 2, 3, 4], sessions_needed=2, mandatory_sessions=2),
            course("SECOND", no_of_days=3, lab_days=[3], sessions_needed=2, mandatory_sessions=2),
        ]
        qualifications = [
            {"trainer_name": trainer["trainer"], "course_code": item["course_code"]}
            for trainer in trainers for item in requirements
        ]
        fixed = [
            booking("FIRST", start="2027-01-01", end="2027-01-08"),
            booking("SECOND", start="2027-01-01", end="2027-01-08"),
        ]
        rows = _schedule_resource_feasible_sessions(requirements, trainers, qualifications, fixed)
        self.assertEqual(len(rows), 4)
        _validate_sessions(rows, requirements, trainers, qualifications)
        _assign_bay_rooms(rows, fixed, requirements)
        lab_dates = [day for row in rows for day in row["lab_dates"]]
        self.assertEqual(len(lab_dates), len(set(lab_dates)))

    def test_joint_search_reports_impossible_resource_requirements(self):
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        qualifications = [{"trainer_name": "A Trainer", "course_code": "COURSE"}]
        with self.assertRaisesRegex(ValueError, "No feasible dates.*COURSE"):
            _schedule_resource_feasible_sessions(
                [course()], trainers, qualifications,
                [booking(start="2027-01-01", end="2027-03-31")],
            )
        with self.assertRaisesRegex(ValueError, "No feasible dates.*COURSE"):
            _schedule_resource_feasible_sessions([course()], trainers, qualifications, [])

    def test_no_lab_needs_no_bay(self):
        rows = [session()]
        _assign_bay_rooms(rows, [], [course(lab_days=[])])
        self.assertEqual(rows[0]["bay_room"], "Not required")

    def test_mandatory_and_trainer_validation(self):
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        qualifications = [{"trainer_name": "A Trainer", "course_code": "COURSE"}]
        _validate_sessions([session()], [course()], trainers, qualifications)
        with self.assertRaisesRegex(ValueError, "Mandatory minimum"):
            _validate_sessions([], [course(existing_sessions=10)], trainers, qualifications)
        with self.assertRaisesRegex(ValueError, "unavailable"):
            _validate_sessions([session()], [course()], [
                {"trainer": "A Trainer", "leave_dates": ["2027-01-06"]}
            ], qualifications)
        with self.assertRaisesRegex(ValueError, "qualified"):
            _validate_sessions([session()], [course()], trainers, [])

    def test_empty_requirements_generate_workbook_without_llm(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "filtered.xlsx"
            pd.DataFrame(columns=["Local Course Code", "Course Title"]).to_excel(path, index=False)
            with patch("services.session_calculator_service.OUTPUT_DIR", temp), \
                 patch("services.planner_service.OUTPUT_DIR", temp), \
                 patch("services.session_calculator_service.load_course_schedule", return_value={}):
                requirements = calculate_sessions(str(path))
                result = plan_courses(requirements)
            self.assertEqual(result["total_sessions_planned"], 0)
            with pd.ExcelFile(result["output_file"]) as workbook:
                self.assertIn("Existing Bay Bookings", workbook.sheet_names)
                self.assertIn("Lab Availability", workbook.sheet_names)
                self.assertIn("Session Requirements", workbook.sheet_names)
            self.assertFalse((Path(temp) / "session_requirements_v3.xlsx").exists())

    def test_api_pipeline_finds_intermediate_files(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "input.xlsx"
            pd.DataFrame(columns=[
                "Comp Short Ttl", "Req Site Cd", "Local Course Code", "Course Title"
            ]).to_excel(source, index=False)
            output = root / "output"
            with patch("helpers.plan_helpers.UPLOAD_INPUT_DIR", root / "uploads"), \
                 patch("helpers.plan_helpers.OUTPUT_DIR", output), \
                 patch("services.filter_service.OUTPUT_DIR", str(output)), \
                 patch("services.session_calculator_service.OUTPUT_DIR", str(output)), \
                 patch("services.planner_service.OUTPUT_DIR", str(output)), \
                 patch("services.session_calculator_service.load_course_schedule", return_value={}):
                response = start_plan_job(
                    "input.xlsx", base64.b64encode(source.read_bytes()).decode("ascii"), None
                )
            self.assertEqual(response["status"], "success", response.get("message"))
            self.assertTrue(base64.b64decode(response["file_content_base64"]).startswith(b"PK"))

    def test_api_returns_skipped_course_warning(self):
        skipped = [{"course_code": "COURSE", "sessions_needed": 2,
                    "mandatory_sessions": 2, "reason": "No qualified Cleveland roster trainer"}]
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "input.xlsx"
            pd.DataFrame().to_excel(source, index=False)
            with patch("helpers.plan_helpers.UPLOAD_INPUT_DIR", root / "uploads"), \
                 patch("helpers.plan_helpers.OUTPUT_DIR", root), \
                 patch("services.filter_service.apply_filters", return_value={"output_file": "filtered.xlsx"}), \
                 patch("services.session_calculator_service.calculate_sessions", return_value={"courses": []}), \
                 patch("services.planner_service.plan_courses", return_value={
                     "output_file": str(source), "skipped_courses": skipped,
                 }):
                result = start_plan_job("input.xlsx", base64.b64encode(source.read_bytes()).decode("ascii"), None)
            self.assertEqual(result["status"], "success")
            self.assertEqual(result["skipped_courses"], skipped)
            self.assertIn("unscheduled", result["message"])

    def test_all_unqualified_courses_are_reported_without_invoking_llm(self):
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        priorities = pd.DataFrame(columns=["priority_1", "priority_2", "priority_3", "priority_4", "priority_5"])
        with tempfile.TemporaryDirectory() as temp:
            with patch("services.planner_service.OUTPUT_DIR", temp), \
                 patch("services.planner_service._load_trainer_leave_dates", return_value=(trainers, "")), \
                 patch("services.planner_service._load_trainer_priority", return_value=priorities), \
                 patch("services.planner_service._load_trainer_qualifications", return_value=[]), \
                 patch("services.planner_service._load_asset_bay_bookings", return_value=[]), \
                 patch("services.planner_service._write_trainer_availability_sheet"), \
                 patch("services.planner_service._write_lab_availability_sheet"), \
                 patch("services.planner_service._write_utilization_metrics_sheet"):
                result = plan_courses({"courses": [
                    course(students_interested=0, mandatory_sessions=2, sessions_needed=2),
                    course("OPTIONAL", students_interested=5, mandatory_sessions=0),
                ]})
            self.assertEqual(result["total_sessions_planned"], 0)
            self.assertEqual([x["course_code"] for x in result["skipped_courses"]], ["COURSE", "OPTIONAL"])
            with pd.ExcelFile(result["output_file"]) as workbook:
                overview = pd.read_excel(workbook, sheet_name="Course Overview")
                requirements = pd.read_excel(workbook, sheet_name="Session Requirements")
            row = overview.loc[overview["Course Code"] == "COURSE"].iloc[0]
            self.assertEqual(row["Sessions Unscheduled"], 2)
            self.assertEqual(row["Unscheduled Reason"], "No qualified Cleveland roster trainer")
            self.assertEqual(requirements.loc[0, "sessions_needed"], 2)
            self.assertEqual(requirements.loc[0, "unscheduled_reason"], row["Unscheduled Reason"])

    def test_new_partial_lab_session_reaches_all_workbook_sheets(self):
        requirement = course(students_interested=0)
        planned = session()
        trainers = [{"trainer": "A Trainer", "leave_dates": []}]
        qualifications = [{"trainer_name": "A Trainer", "course_code": "COURSE"}]
        priorities = pd.DataFrame([{
            "course_code": "COURSE", "priority_1": "A Trainer",
            "priority_2": None, "priority_3": None, "priority_4": None, "priority_5": None,
        }])

        class FakePrompt:
            feedback = []

            def __or__(self, _other):
                return self

            def format_messages(self, **_kwargs):
                return [SimpleNamespace(content="test")]

            def invoke(self, kwargs):
                self.courses_json = kwargs["courses_json"]
                self.feedback.append(kwargs["validation_feedback"])
                return CoursePlan(sessions=[CourseSession(**planned)])

        fake_prompt = FakePrompt()
        llm_module = ModuleType("utils.llm_utils")
        llm_module.get_llm_from_env = lambda: SimpleNamespace(with_structured_output=lambda _model: None)
        with tempfile.TemporaryDirectory() as temp:
            with patch.dict("sys.modules", {"utils.llm_utils": llm_module}), \
                 patch("services.planner_service.OUTPUT_DIR", temp), \
                 patch("services.planner_service._load_trainer_leave_dates", return_value=(trainers, "")), \
                 patch("services.planner_service._load_trainer_priority", return_value=priorities), \
                 patch("services.planner_service._load_trainer_qualifications", return_value=qualifications), \
                 patch("services.planner_service._load_asset_bay_bookings", return_value=[booking(end="2027-01-06")]), \
                 patch("services.planner_service._load_trainer_locations", return_value={}), \
                 patch("services.planner_service._write_trainer_availability_sheet"), \
                 patch("services.planner_service._write_utilization_metrics_sheet"), \
                 patch("services.planner_service.get_course_planner_prompt", return_value=fake_prompt), \
                 patch("services.planner_service._count_tokens", return_value=1):
                result = plan_courses({"courses": [
                    requirement,
                    course("NO-TRAINER", students_interested=0, mandatory_sessions=1),
                ]})
            row = result["plan"][0]
            self.assertEqual(len(row["lab_dates"]), 1)
            self.assertNotIn(row["lab_dates"][0], ["2027-01-04", "2027-01-05", "2027-01-06"])
            self.assertEqual(len(fake_prompt.feedback), 1)
            self.assertNotIn("NO-TRAINER", fake_prompt.courses_json)
            self.assertEqual(result["skipped_courses"][0]["course_code"], "NO-TRAINER")
            workbook = load_workbook(result["output_file"], data_only=True)
            self.assertEqual(workbook["Cleveland Course Plan"]["G2"].value, "Bay 1")
            self.assertEqual(workbook["Cleveland Course Plan"]["H2"].value, row["lab_dates"][0])
            self.assertEqual(workbook["Course Overview"]["F2"].value, 1)
            self.assertEqual(workbook["Course Overview"]["G3"].value, 1)
            self.assertEqual(workbook["Course Overview"]["J3"].value, "No qualified Cleveland roster trainer")
            self.assertEqual(workbook["Session Requirements"]["A2"].value, "COURSE")
            self.assertEqual(workbook["Session Requirements"]["H2"].value, 1)
            sheet = workbook["Lab Availability"]
            bay_row = next(row for row in range(1, sheet.max_row + 1) if sheet.cell(row, 1).value == "Bay 1")
            lab_column = (date.fromisoformat(row["lab_dates"][0]) - cleveland_data.PLANNING_START).days + 2
            lab_cell = sheet.cell(bay_row, lab_column)
            for merged in sheet.merged_cells.ranges:
                if lab_cell.coordinate in merged:
                    lab_cell = sheet.cell(merged.min_row, merged.min_col)
                    break
            self.assertEqual(lab_cell.value, "COURSE")
            workbook.close()


if __name__ == "__main__":
    unittest.main()
