import math
import os
import tempfile
import traceback

import pandas as pd

INPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "input")

# Azure Functions deploys application files under /home/site/wwwroot, which is
# read-only. Use HOME on Azure and /tmp during local development instead.
_HOME_BASE = os.path.join(os.environ.get("HOME", ""), "smart_planner", "output") if os.environ.get("HOME") else None
_DEFAULT_OUTPUT_DIR = _HOME_BASE or os.path.join(tempfile.gettempdir(), "smart_planner", "output")
OUTPUT_DIR = os.environ.get("SMART_PLANNER_OUTPUT_DIR") or _DEFAULT_OUTPUT_DIR

FILTERED_FILE = "filtered_output_v3.xlsx"
COURSE_SCHEDULE_FILE = "course_schedule_days.xlsx"

MODALITY_VALUES = ["MODALITY IXR"]
SITE_CD_VALUES = ["BEST", "PHC", "CL", "SLC", "VC"]

# Corrections for known data-entry errors in course_schedule_days.xlsx.
# IGT2BL001 (Azurion Essentials) classroom note says "QS003 wk1+2 / QS147 wk3" → 3 weeks = 15 working days,
# but the spreadsheet mistakenly shows 19.
DAYS_OVERRIDE: dict[str, int] = {
    "IGT2BL001": 15,
}


def calculate_sessions(filtered_file: str = None) -> dict:
    """Count students interested per course from filtered SRL data,
    look up max capacity from course_schedule_days, and compute
    the number of sessions needed (ceil(students / max_capacity))."""
    try:
        filtered_path = os.path.join(
            OUTPUT_DIR, filtered_file or FILTERED_FILE
        )
        print(f"[SESSION] Reading filtered file: {filtered_path}")
        filtered_df = pd.read_excel(filtered_path)

        # Count students per course
        student_counts = (
            filtered_df.groupby(["Local Course Code", "Course Title"])
            .size()
            .reset_index(name="students_interested")
        )
        print(f"[SESSION] Found {len(student_counts)} unique courses")

        # Load course schedule for capacity info
        schedule_path = os.path.join(INPUT_DIR, COURSE_SCHEDULE_FILE)
        print(f"[SESSION] Loading schedule from: {schedule_path}")
        schedule_df = pd.read_excel(schedule_path)

        # Build a lookup: Course Code -> Maximal Capacity
        capacity_lookup = (
            schedule_df.groupby("Course Code")["Maximal Capacity"]
            .max()
            .to_dict()
        )

        # Also build a lookup for No. of Days
        days_lookup = (
            schedule_df.groupby("Course Code")["No. of Days"]
            .max()
            .to_dict()
        )

        results = []
        for _, row in student_counts.iterrows():
            course_code = row["Local Course Code"]
            course_title = row["Course Title"]
            students = int(row["students_interested"])
            max_capacity = capacity_lookup.get(course_code)
            no_of_days = days_lookup.get(course_code)
            # Apply known corrections
            no_of_days = DAYS_OVERRIDE.get(course_code, no_of_days)

            if max_capacity and max_capacity > 0:
                sessions_needed = math.ceil(students / max_capacity)
            else:
                # If course not found in schedule, flag it
                sessions_needed = None

            results.append(
                {
                    "course_code": course_code,
                    "course_title": course_title,
                    "students_interested": students,
                    "max_capacity": int(max_capacity) if max_capacity else None,
                    "no_of_days": int(no_of_days) if no_of_days else None,
                    "sessions_needed": sessions_needed,
                }
            )

        result_df = pd.DataFrame(results)

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        output_path = os.path.join(OUTPUT_DIR, "session_requirements_v3.xlsx")
        result_df.to_excel(output_path, index=False)
        print(f"[SESSION] Saved {len(results)} courses to {output_path}")

        return {
            "courses": results,
            "output_file": "session_requirements_v3.xlsx",
        }
    except Exception as e:
        print(f"[SESSION] ERROR: {e}")
        print(f"[SESSION] Traceback: {traceback.format_exc()}")
        raise
