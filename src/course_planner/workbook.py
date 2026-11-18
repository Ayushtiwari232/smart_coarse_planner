import json
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from course_planner.inputs import PlanningData
from course_planner.reporting import WorkbookReports


class ReportWriter(WorkbookReports):
    def __init__(self, data: PlanningData):
        super().__init__(data.start, data.end, data.compatibility)
        self.data = data
        self.location = data.location
        self.input_dir = data.input_dir

    def _load_trainer_locations(self) -> dict[str, dict]:
        return {
            alias: {"location": self.location, "display_name": trainer["trainer"]}
            for trainer in self.data.trainers
            for alias in self._trainer_name_aliases(trainer["trainer"])
        }

    def _load_asset_lab_bookings(self) -> list[dict]:
        return self.data.bookings

    def _load_parttime_days_from_json(self) -> dict[str, set]:
        result = {}
        for name, days in self.data.parttime_dates.items():
            for alias in self._trainer_name_aliases(name):
                result[alias] = days
        return result

    def write(self, path: Path, rows: list[dict], manifest: dict) -> None:
        data = self.data
        columns = [
            "course_code",
            "course_name",
            "session_number",
            "start_date",
            "end_date",
            "trainer_name",
            "lab_room",
            "lab_dates",
            "location",
        ]
        frame = pd.DataFrame(rows, columns=columns)
        frame["lab_dates"] = frame["lab_dates"].apply(lambda values: ", ".join(values))
        frame["location"] = self.location
        if self.location == "Best":
            # Keep the original Best machine-readable column names.
            classrooms = {c["course_code"]: c.get("classroom", "") for c in data.courses}
            frame["training_room"] = frame["course_code"].map(classrooms)
        else:
            frame = frame.rename(
                columns={
                    "course_code": "Course Code",
                    "course_name": "Course Name",
                    "session_number": "Session Number",
                    "start_date": "Start Date",
                    "end_date": "End Date",
                    "trainer_name": "Trainer",
                    "lab_room": "Lab Room",
                    "lab_dates": "Lab Dates",
                    "location": "Location",
                }
            )
        requirements = pd.DataFrame(
            data.courses,
            columns=[
                "course_code",
                "course_title",
                "students_interested",
                "max_capacity",
                "no_of_days",
                "mandatory_sessions",
                "demand_sessions",
                "sessions_needed",
                "lab_days",
                "unscheduled_reason",
            ],
        )
        requirements["lab_days"] = requirements["lab_days"].apply(
            lambda days: ", ".join(map(str, days))
        )
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            frame.to_excel(writer, index=False, sheet_name=f"{self.location} Course Plan")
            requirements.to_excel(writer, index=False, sheet_name="Session Requirements")
            pd.DataFrame(
                data.bookings, columns=["lab", "course_code", "start_date", "end_date"]
            ).to_excel(writer, index=False, sheet_name="Existing Lab Bookings")
            notes = [{"Type": "Location note", "Text": note} for note in data.rules.notes] + [
                {"Type": "Warning / assumption", "Text": note} for note in data.warnings
            ]
            notes.append(
                {
                    "Type": "Scheduling restriction",
                    "Text": "Excluded teaching dates: "
                    + ", ".join(sorted(day.isoformat() for day in data.excluded_dates)),
                }
            )
            pd.DataFrame(notes, columns=["Type", "Text"]).to_excel(
                writer, index=False, sheet_name="Notes and Assumptions"
            )
            pd.DataFrame(
                [
                    {
                        "Field": name,
                        "Value": json.dumps(value) if isinstance(value, (dict, list)) else value,
                    }
                    for name, value in manifest.items()
                ]
            ).to_excel(writer, index=False, sheet_name="Run Manifest")
        priorities = self._priority_trainer_names(data.priorities)
        self._write_trainer_availability_sheet(str(path), data.trainers, rows, priorities)
        self._write_course_overview_sheet(str(path), data.courses, rows)
        report_rows = rows
        if self.location == "Best":
            report_rows = []
            classrooms = {c["course_code"]: c.get("classroom", "") for c in data.courses}
            for row in rows:
                for room, start, end in self._parse_classrooms(
                    classrooms.get(row["course_code"], ""),
                    date.fromisoformat(row["start_date"]),
                    date.fromisoformat(row["end_date"]),
                ):
                    if room:
                        report_rows.append(
                            {
                                **row,
                                "lab_room": f"Lab {room}",
                                "lab_dates": [
                                    day.isoformat() for day in self._working_days(start, end)
                                ],
                            }
                        )
        self._write_lab_availability_sheet(str(path), report_rows)
        self.lab_report_rows = report_rows
        self._write_utilization_metrics_sheet(str(path), data.trainers, rows, priorities)
        workbook = load_workbook(path)
        try:
            if "Lab Availability" not in workbook.sheetnames:
                workbook.create_sheet("Lab Availability").append(["No lab bookings"])
            final_sheets = ["Existing Lab Bookings", "Notes and Assumptions", "Run Manifest"]
            for name in final_sheets:
                workbook.move_sheet(name, offset=len(workbook.sheetnames))
            workbook.active = 0
            workbook.save(path)
        finally:
            workbook.close()
        workbook = load_workbook(path, read_only=True)
        try:
            required = {
                "Trainer Availability",
                "Lab Availability",
                "Course Overview",
                "Utilization Metrics",
                "Session Requirements",
                "Existing Lab Bookings",
                "Notes and Assumptions",
                "Run Manifest",
                f"{self.location} Course Plan",
            }
            if not required <= set(workbook.sheetnames):
                raise RuntimeError("Workbook is missing required report sheets")
            if workbook.sheetnames[-3:] != final_sheets:
                raise RuntimeError("Workbook audit sheets are not last")
        finally:
            workbook.close()
