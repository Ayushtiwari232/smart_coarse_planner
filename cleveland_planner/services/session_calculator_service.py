import math
import os
import traceback

import pandas as pd

from services.cleveland_data import (
    load_bay_bookings,
    load_course_schedule,
)

_DEFAULT_OUTPUT_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "data", "output")
)
OUTPUT_DIR = os.environ.get("SMART_PLANNER_OUTPUT_DIR") or _DEFAULT_OUTPUT_DIR

FILTERED_FILE = "filtered_output_v3.xlsx"
def calculate_sessions(filtered_file: str = None) -> dict:
    """Calculate trainer-assigned sessions from demand and quarterly minimums."""
    try:
        filtered_path = os.path.join(
            OUTPUT_DIR, filtered_file or FILTERED_FILE
        )
        print(f"[SESSION] Reading filtered file: {filtered_path}")
        filtered_df = pd.read_excel(filtered_path)

        # Count students per course
        student_counts = filtered_df.groupby("Local Course Code").size().to_dict()
        print(f"[SESSION] Found {len(student_counts)} unique courses")

        bookings = load_bay_bookings()
        schedule = load_course_schedule(bookings)
        results = []
        codes = set(student_counts) | {
            code for code, course in schedule.items() if course["mandatory_sessions"]
        }
        for course_code in sorted(codes):
            students = int(student_counts.get(course_code, 0))
            course = schedule.get(course_code)
            if course is None:
                title = filtered_df.loc[
                    filtered_df["Local Course Code"] == course_code, "Course Title"
                ].iloc[0]
                results.append({
                    "course_code": course_code, "course_title": str(title),
                    "students_interested": students, "max_capacity": None,
                    "no_of_days": None, "mandatory_sessions": 0,
                    "lab_days": [], "sessions_needed": None,
                })
                continue
            demand_sessions = math.ceil(students / course["max_capacity"])
            total_needed = max(demand_sessions, course["mandatory_sessions"])
            results.append(
                {
                    "course_code": course_code,
                    **course,
                    "students_interested": students,
                    "demand_sessions": demand_sessions,
                    "sessions_needed": total_needed,
                }
            )

        return {"courses": results}
    except Exception as e:
        print(f"[SESSION] ERROR: {e}")
        print(f"[SESSION] Traceback: {traceback.format_exc()}")
        raise
