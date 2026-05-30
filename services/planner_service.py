import json
import os

import pandas as pd
from langchain_core.prompts import ChatPromptTemplate
from pydantic import BaseModel, Field

from services.llm_utils import get_llm
from services.session_calculator_service import calculate_sessions

INPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "input")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "output")

TRAINERS_FILE = "trainers.json"
PRIORITY_FILE = "priority_to_train_list.xlsx"


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


def _load_trainer_availability() -> list[dict]:
    """Load trainer availability from trainers.json."""
    path = os.path.join(INPUT_DIR, TRAINERS_FILE)
    with open(path, "r") as f:
        data = json.load(f)
    return data["trainers"], data.get("period", "")


def _load_trainer_priority() -> pd.DataFrame:
    """Load trainer priority list. The file has header row in the data,
    so we parse accordingly."""
    path = os.path.join(INPUT_DIR, PRIORITY_FILE)
    df = pd.read_excel(path)
    # First data row is the sub-header: 1, 2, 3, 4, 5
    # Rename columns properly
    df.columns = [
        "course_code",
        "course_name",
        "priority_1",
        "priority_2",
        "priority_3",
        "priority_4",
        "priority_5",
    ]
    # Drop the header row (row 0 which has '1','2','3','4','5')
    df = df.iloc[1:].reset_index(drop=True)
    return df


def plan_courses(filtered_file: str = None) -> dict:
    """Main planning function:
    1. Calculate session requirements
    2. Load trainer availability and priority
    3. Use LLM to assign trainers and dates
    4. Save result to Excel
    """

    # Step 1: Get session requirements
    session_data = calculate_sessions(filtered_file)
    courses = session_data["courses"]

    # Step 2: Load trainer data
    trainers, period = _load_trainer_availability()
    priority_df = _load_trainer_priority()

    # Step 3: Build prompt context
    trainer_availability_summary = []
    for t in trainers:
        avail_dates = []
        for month_name, month_data in t["months"].items():
            for r in month_data.get("available_ranges", []):
                avail_dates.append(f"{r['start']} to {r['end']}")
        trainer_availability_summary.append(
            {
                "trainer": t["trainer"],
                "available_dates": avail_dates,
            }
        )

    priority_records = priority_df.to_dict(orient="records")

    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                """You are a training course planner. Your task is to schedule course sessions
and assign trainers based on the following rules:

1. Each course needs a specific number of sessions (provided below).
2. Each session spans a number of consecutive working days (no_of_days provided).
3. Trainers must be assigned based on priority order (priority_1 is most preferred,
   then priority_2, etc.). Use the highest priority trainer who is available.
4. A trainer can only teach one course at a time - no overlapping assignments.
5. Sessions should be scheduled within the planning period: {period}.
6. Only schedule on weekdays (Monday-Friday).
7. Only assign a trainer if they are available for ALL days of that session.
8. If a course has no matching trainer in the priority list, pick the best available
   trainer from the availability data who is not otherwise occupied.
9. Spread sessions across the planning period when possible.
10. If a course is not found in the schedule (sessions_needed is null), skip it.""",
            ),
            (
                "human",
                """Here are the courses that need sessions planned:

{courses_json}

Here is the trainer priority for each course (priority_1 is most preferred):

{priority_json}

Here is the trainer availability (list of date ranges when each trainer is free):

{availability_json}

Please create a complete course plan assigning dates and trainers to each session.""",
            ),
        ]
    )

    llm = get_llm()
    structured_llm = llm.with_structured_output(CoursePlan)

    chain = prompt | structured_llm

    result: CoursePlan = chain.invoke(
        {
            "period": period,
            "courses_json": json.dumps(courses, indent=2),
            "priority_json": json.dumps(priority_records, indent=2, default=str),
            "availability_json": json.dumps(
                trainer_availability_summary, indent=2
            ),
        }
    )

    # Step 4: Save to Excel
    rows = [session.model_dump() for session in result.sessions]
    result_df = pd.DataFrame(rows)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, "course_plan.xlsx")
    result_df.to_excel(output_path, index=False)

    return {
        "total_sessions_planned": len(result.sessions),
        "output_file": "course_plan.xlsx",
        "plan": rows,
    }
