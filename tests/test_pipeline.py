import base64
import json
import os
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from openpyxl import load_workbook

from course_planner import api
from course_planner.inputs import INPUT_FILES, load_inputs, whole_number
from course_planner.models import PlanRequest, Settings, quarter_dates
from course_planner.pipeline import create_run_directory, get_run, run_plan


def fixture(root: Path, location: str = "cleveland", mandatory: int = 1) -> Path:
    directory = root / "data" / "quarters" / "2027-Q1" / location
    directory.mkdir(parents=True)
    shared = directory.parent / "shared"
    shared.mkdir(exist_ok=True)
    tables = {
        "courses.xlsx": pd.DataFrame(
            [
                {
                    "course_code": "COURSE",
                    "course_title": "A course",
                    "no_of_days": 3,
                    "max_capacity": 5,
                    "mandatory_sessions": mandatory,
                    "lab_days": "3",
                    "classroom": "Room A",
                }
            ]
        ),
        "trainers.xlsx": pd.DataFrame([{"trainer_id": "trainer-a", "trainer": "A Trainer"}]),
        "qualifications.xlsx": pd.DataFrame([{"trainer_id": "trainer-a", "course_code": "COURSE"}]),
        "priorities.xlsx": pd.DataFrame(
            [{"trainer_id": "trainer-a", "course_code": "COURSE", "rank": 1}]
        ),
        "availability.xlsx": pd.DataFrame(columns=["trainer_id", "date", "available_fraction"]),
        "lab_compatibility.xlsx": pd.DataFrame([{"course_code": "COURSE", "lab": "1"}]),
        "bookings.xlsx": pd.DataFrame(columns=["lab", "course_code", "start_date", "end_date"]),
    }
    for name, table in tables.items():
        table.to_excel(directory / name, index=False)
    (directory / "rules.json").write_text(
        json.dumps({"notes": ["Advisory: discuss timing with the planner."]})
    )
    pd.DataFrame(
        [
            {
                "Comp Short Ttl": "MODALITY IXR",
                "Req Site Cd": site,
                "Local Course Code": "COURSE",
                "Course Title": "A course",
            }
            for site in ("BEST", "CL", "SLC", "PHC")
        ]
    ).to_excel(shared / "demand.xlsx", index=False)
    return directory


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.directory = fixture(self.root)
        self.settings = Settings(data_root=self.root / "data", runs_root=self.root / "runs")
        self.environment = patch.dict(os.environ, {"PLANNER_SHAREPOINT_ENABLED": "false"})
        self.environment.start()
        self.addCleanup(self.environment.stop)

    def test_quarter_validation(self):
        self.assertEqual(str(quarter_dates("2028-Q1")[1]), "2028-03-31")
        self.assertEqual(str(quarter_dates("2027-Q4")[1]), "2027-12-31")
        for invalid in ("2027-Q0", "Q1", "../2027", "2027-Q5"):
            with self.assertRaises(ValueError):
                PlanRequest(quarter=invalid)

    def test_all_locations_use_separate_site_filters(self):
        for location in ("best", "singapore"):
            fixture(self.root, location)
        for location in ("best", "cleveland", "singapore"):
            data = load_inputs(self.settings.data_root, location, "2027-Q1")
            self.assertEqual(len(data.filtered_demand), 1)
            self.assertEqual(data.courses[0]["sessions_needed"], 1)

    def test_mandatory_without_demand_is_not_reduced_by_lab_bookings(self):
        path = self.directory.parent / "shared" / "demand.xlsx"
        pd.read_excel(path).iloc[:0].to_excel(path, index=False)
        data = load_inputs(self.settings.data_root, "cleveland", "2027-Q1")
        self.assertEqual(data.courses[0]["sessions_needed"], 1)
        self.assertEqual(data.courses[0]["students_interested"], 0)

    def test_numeric_fields_reject_invalid_values(self):
        for value in ("bad", -1, 1.5, "nan", float("inf")):
            with self.assertRaises(ValueError):
                whole_number(value, "field")

    def test_unknown_trainer_id_fails_explicitly(self):
        pd.DataFrame([{"trainer_id": "missing", "course_code": "COURSE"}]).to_excel(
            self.directory / "qualifications.xlsx", index=False
        )
        with self.assertRaisesRegex(ValueError, "Unknown qualification"):
            load_inputs(self.settings.data_root, "cleveland", "2027-Q1")

    def test_missing_files_are_reported(self):
        (self.directory / "courses.xlsx").unlink()
        with self.assertRaisesRegex(ValueError, "courses.xlsx"):
            load_inputs(self.settings.data_root, "cleveland", "2027-Q1")

    def test_duplicate_courses_are_rejected(self):
        path = self.directory / "courses.xlsx"
        frame = pd.read_excel(path)
        pd.concat([frame, frame]).to_excel(path, index=False)
        with self.assertRaisesRegex(ValueError, "Duplicate"):
            load_inputs(self.settings.data_root, "cleveland", "2027-Q1")

    def test_skip_first_week_and_excluded_dates_are_enforced(self):
        (self.directory / "rules.json").write_text(
            json.dumps(
                {
                    "notes": ["Skip the last week (advisory text only)."],
                    "skip_first_week": True,
                    "excluded_dates": ["2027-02-15"],
                }
            )
        )
        result = run_plan(PlanRequest(location="cleveland", mode="deterministic"), self.settings)
        frame = pd.read_excel(result["output_file"])
        for row in frame.itertuples(index=False, name=None):
            self.assertGreaterEqual(row[3], "2027-01-08")
        data = load_inputs(self.settings.data_root, "cleveland", "2027-Q1")
        self.assertIn("2027-02-15", data.trainers[0]["leave_dates"])
        self.assertNotIn("2027-03-29", data.trainers[0]["leave_dates"])

    def test_out_of_quarter_rule_is_rejected(self):
        (self.directory / "rules.json").write_text('{"excluded_dates": ["2026-01-01"]}')
        with self.assertRaisesRegex(ValueError, "selected quarter"):
            load_inputs(self.settings.data_root, "cleveland", "2027-Q1")

    def test_fractional_availability_blocks_full_day_session(self):
        pd.DataFrame(
            [{"trainer_id": "trainer-a", "date": "2027-01-04", "available_fraction": 0.5}]
        ).to_excel(self.directory / "availability.xlsx", index=False)
        data = load_inputs(self.settings.data_root, "cleveland", "2027-Q1")
        self.assertIn("2027-01-04", data.trainers[0]["leave_dates"])
        self.assertIn(pd.Timestamp("2027-01-04").date(), data.parttime_dates["A Trainer"])

    def test_production_can_reject_mock_quarters(self):
        (self.directory / "rules.json").write_text('{"mock_data": true}')
        self.settings.allow_mock_data = False
        with self.assertRaisesRegex(ValueError, "marked as mock"):
            run_plan(PlanRequest(location="cleveland", mode="deterministic"), self.settings)

    def test_new_quarter_clears_time_sensitive_data_and_refuses_overwrite(self):
        from course_planner.cli import init_quarter

        with patch("course_planner.cli.get_settings", return_value=self.settings):
            target = init_quarter("cleveland", "2027-Q2", "2027-Q1")
            self.assertTrue(pd.read_excel(target / "availability.xlsx").empty)
            self.assertTrue(pd.read_excel(target / "bookings.xlsx").empty)
            self.assertTrue(pd.read_excel(target.parent / "shared" / "demand.xlsx").empty)
            self.assertTrue(json.loads((target / "rules.json").read_text())["mock_data"])
            with self.assertRaises(FileExistsError):
                init_quarter("cleveland", "2027-Q2", "2027-Q1")

    def test_demand_only_ai_under_scheduling_is_repaired(self):
        path = self.directory / "courses.xlsx"
        table = pd.read_excel(path)
        table["mandatory_sessions"] = 0
        table.to_excel(path, index=False)
        with patch("course_planner.llm.propose", return_value=[]):
            result = run_plan(PlanRequest(location="cleveland"), self.settings)
        self.assertEqual(result["total_sessions_planned"], 1)

    def test_existing_bookings_do_not_define_course_compatibility(self):
        pd.DataFrame(
            [
                {
                    "lab": "WRONG",
                    "course_code": "COURSE",
                    "start_date": "2027-01-01",
                    "end_date": "2027-03-31",
                }
            ]
        ).to_excel(self.directory / "bookings.xlsx", index=False)
        result = run_plan(PlanRequest(location="cleveland", mode="deterministic"), self.settings)
        self.assertEqual(pd.read_excel(result["output_file"]).iloc[0]["Lab Room"], "Lab 1")

    def test_output_is_complete_and_status_survives_reload(self):
        result = run_plan(PlanRequest(location="cleveland", mode="deterministic"), self.settings)
        self.assertEqual(result["total_sessions_planned"], 1)
        self.assertEqual(
            get_run(self.settings, result["job_id"])["output_file"], result["output_file"]
        )
        workbook = load_workbook(result["output_file"], read_only=True)
        try:
            self.assertTrue(
                {
                    "Session Requirements",
                    "Notes and Assumptions",
                    "Run Manifest",
                    "Lab Availability",
                    "Trainer Availability",
                    "Course Overview",
                    "Utilization Metrics",
                }
                <= set(workbook.sheetnames)
            )
        finally:
            workbook.close()
        self.assertEqual(
            {path.name for path in (Path(result["run_directory"]) / "inputs").iterdir()},
            {*INPUT_FILES, "demand.xlsx"},
        )
        self.assertFalse((Path(result["run_directory"]) / "pending.xlsx").exists())

    def test_runs_have_distinct_paths_and_immutable_inputs(self):
        one = run_plan(PlanRequest(location="cleveland", mode="deterministic"), self.settings)
        two = run_plan(PlanRequest(location="cleveland", mode="deterministic"), self.settings)
        self.assertNotEqual(one["output_file"], two["output_file"])
        self.assertTrue(Path(one["output_file"]).exists())
        manifest = json.loads((Path(one["run_directory"]) / "manifest.json").read_text())
        self.assertEqual(len(manifest["input_sha256"]), len(INPUT_FILES) + 1)
        directory = Path(one["run_directory"])
        self.assertEqual(
            directory.parent, (self.settings.runs_root / "2027-Q1" / "cleveland").resolve()
        )
        datetime.strptime(directory.name, "%Y-%m-%d_%H%M%S_%f")
        self.assertNotEqual(one["job_id"], directory.name)
        snapshot = directory / "inputs" / "courses.xlsx"
        original = snapshot.read_bytes()
        (self.directory / "courses.xlsx").write_bytes(b"changed after run")
        self.assertEqual(snapshot.read_bytes(), original)

    def test_run_directory_collision_never_overwrites_previous_run(self):
        fixed = datetime(2026, 11, 18, 12, 0, 0, 123456, tzinfo=timezone.utc)
        later = fixed.replace(microsecond=123457)
        request = PlanRequest(location="cleveland")
        with patch("course_planner.pipeline.datetime") as clock:
            clock.now.side_effect = [fixed, fixed, later]
            first = create_run_directory(self.settings, request)
            (first / "existing.txt").write_text("preserved")
            second = create_run_directory(self.settings, request)
        self.assertNotEqual(first, second)
        self.assertEqual((first / "existing.txt").read_text(), "preserved")
        with patch("course_planner.pipeline.datetime") as clock:
            clock.now.return_value = fixed
            with self.assertRaisesRegex(FileExistsError, "unique timestamped"):
                create_run_directory(self.settings, request)

    def test_legacy_job_ids_remain_accessible_after_folder_migration(self):
        result = run_plan(PlanRequest(location="cleveland", mode="deterministic"), self.settings)
        status_path = Path(result["run_directory"]) / "status.json"
        result["job_id"] = "b" * 32
        status_path.write_text(json.dumps(result))
        self.assertEqual(get_run(self.settings, result["job_id"]), result)
        old_directory = status_path.parent.with_name(result["job_id"])
        status_path.parent.rename(old_directory)
        self.assertEqual(get_run(self.settings, result["job_id"])["job_id"], result["job_id"])

    def test_all_locations_use_lab_naming_and_audit_tabs_last(self):
        for location in ("best", "singapore"):
            fixture(self.root, location)
        for location in ("best", "cleveland", "singapore"):
            with self.subTest(location=location):
                result = run_plan(
                    PlanRequest(location=location, mode="deterministic"), self.settings
                )
                workbook = load_workbook(result["output_file"], read_only=True)
                try:
                    self.assertEqual(
                        workbook.sheetnames[-3:],
                        ["Existing Lab Bookings", "Notes and Assumptions", "Run Manifest"],
                    )
                    self.assertEqual(workbook.active.title, f"{location.title()} Course Plan")
                    for sheet in workbook:
                        self.assertNotIn("Bay", sheet.title)
                        for row in sheet.iter_rows(values_only=True):
                            self.assertFalse(any("bay" in str(cell).lower() for cell in row))
                finally:
                    workbook.close()

    def test_unknown_run_ids_and_traversal_are_rejected(self):
        with self.assertRaises(ValueError):
            get_run(self.settings, "../x")
        for invalid in ("2026-11-18_120000_123456/../x", "*" * 32, "2026-11-18_120000"):
            with self.assertRaises(ValueError):
                get_run(self.settings, invalid)
        with self.assertRaises(FileNotFoundError):
            get_run(self.settings, "a" * 32)

    def test_unqualified_courses_are_reported_not_scheduled(self):
        pd.DataFrame(columns=["trainer_id", "course_code"]).to_excel(
            self.directory / "qualifications.xlsx", index=False
        )
        result = run_plan(PlanRequest(location="cleveland", mode="deterministic"), self.settings)
        self.assertEqual(result["total_sessions_planned"], 0)
        self.assertEqual(result["status"], "completed_with_warnings")
        self.assertEqual(
            result["skipped_courses"][0]["unscheduled_reason"], "No qualified roster trainer"
        )

    def test_incomplete_ai_plan_is_repaired(self):
        with patch("course_planner.llm.propose", return_value=[]):
            result = run_plan(PlanRequest(location="cleveland"), self.settings)
        self.assertEqual(result["total_sessions_planned"], 1)
        self.assertTrue(any("repair" in warning for warning in result["warnings"]))

    def test_valid_ai_plan_keeps_dates(self):
        rows = [
            {
                "course_code": "COURSE",
                "course_name": "A course",
                "session_number": 1,
                "start_date": "2027-01-04",
                "end_date": "2027-01-06",
                "trainer_name": "A Trainer",
            }
        ]
        with patch("course_planner.llm.propose", return_value=rows):
            result = run_plan(PlanRequest(location="cleveland"), self.settings)
        self.assertEqual(pd.read_excel(result["output_file"]).iloc[0]["Start Date"], "2027-01-04")

    def test_required_report_failure_never_publishes_success(self):
        with patch(
            "course_planner.workbook.ReportWriter.write", side_effect=RuntimeError("report failed")
        ):
            with self.assertRaisesRegex(RuntimeError, "report failed"):
                run_plan(PlanRequest(location="cleveland", mode="deterministic"), self.settings)
        statuses = list(self.settings.runs_root.rglob("status.json"))
        self.assertEqual(json.loads(statuses[0].read_text())["status"], "failed")
        self.assertEqual(list(self.settings.runs_root.rglob("*Course_Plan*.xlsx")), [])

    def test_uploaded_demand_replaces_only_demand(self):
        payload = base64.b64encode(
            (self.directory.parent / "shared" / "demand.xlsx").read_bytes()
        ).decode()
        result = run_plan(
            PlanRequest(location="cleveland", mode="deterministic", file_content_base64=payload),
            self.settings,
            validate_only=True,
        )
        self.assertEqual(result["demand_rows"], 1)

    def test_graph_pipeline_downloads_inputs_and_uploads_output(self):
        objects = {
            f"2027-Q1/cleveland/{name}": (self.directory / name).read_bytes()
            for name in INPUT_FILES
        }
        objects["2027-Q1/shared/demand.xlsx"] = (
            self.directory.parent / "shared" / "demand.xlsx"
        ).read_bytes()
        with patch("course_planner.pipeline.GraphStorage.from_env") as factory:
            graph = factory.return_value
            graph.enabled = True
            graph.download.side_effect = lambda path: objects[path]
            self.settings.sharepoint_enabled = True
            result = run_plan(
                PlanRequest(location="cleveland", mode="deterministic"), self.settings
            )
            self.assertEqual(graph.download.call_count, len(INPUT_FILES) + 1)
            self.assertEqual(graph.upload.call_count, 2)
            self.assertTrue(graph.upload.call_args_list[0].args[0].endswith(result["file_name"]))
            self.assertEqual(
                graph.upload.call_args_list[0].args[0],
                f"2027-Q1/cleveland/{Path(result['run_directory']).name}/{result['file_name']}",
            )

    def test_fastapi_uses_shared_pipeline_and_persistent_status(self):
        with patch("course_planner.api.get_settings", return_value=self.settings):
            self.assertEqual(api.health()["status"], "ok")
            result = api.plan(PlanRequest(location="cleveland", mode="deterministic"))
            self.assertIn("file_content_base64", result)
            self.assertEqual(api.get_plan_result(result["job_id"])["status"], "completed")
            file_response = api.get_plan_file(result["job_id"])
            self.assertEqual(file_response.status_code, 200)
            self.assertTrue(Path(file_response.path).read_bytes().startswith(b"PK"))
            self.assertTrue(
                {"/health", "/plan", "/plan/{job_id}", "/plan/{job_id}/file"}
                <= set(api.app.openapi()["paths"])
            )


if __name__ == "__main__":
    unittest.main()
