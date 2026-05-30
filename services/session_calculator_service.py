import math
import os

import pandas as pd

INPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "input")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "output")

FILTERED_FILE = "filtered_output.xlsx"
COURSE_SCHEDULE_FILE = "course_schedule_days.xlsx"

MODALITY_VALUES = ["MODALITY IXR"]
SITE_CD_VALUES = ["BEST", "PHC", "CL", "SLC", "VC"]


def calculate_sessions(filtered_file: str = None) -> dict:
    """Count students interested per course from filtered SRL data,
    look up max capacity from course_schedule_days, and compute
    the number of sessions needed (ceil(students / max_capacity))."""

    filtered_path = os.path.join(
        OUTPUT_DIR, filtered_file or FILTERED_FILE
    )
    filtered_df = pd.read_excel(filtered_path)

    # Count students per course
    student_counts = (
        filtered_df.groupby(["Local Course Code", "Course Title"])
        .size()
        .reset_index(name="students_interested")
    )

    # Load course schedule for capacity info
    schedule_path = os.path.join(INPUT_DIR, COURSE_SCHEDULE_FILE)
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
    output_path = os.path.join(OUTPUT_DIR, "session_requirements.xlsx")
    result_df.to_excel(output_path, index=False)

    return {
        "courses": results,
        "output_file": "session_requirements.xlsx",
    }
