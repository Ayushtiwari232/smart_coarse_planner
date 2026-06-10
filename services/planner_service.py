import json
import os
import time
import traceback
import tiktoken
import pandas as pd
from pydantic import BaseModel, Field
from prompts import get_course_planner_prompt
from utils.llm_utils import get_llm_from_env


def _count_tokens(text: str, model: str = "gpt-4o") -> int:
    """Estimate token count using tiktoken."""
    try:
        enc = tiktoken.encoding_for_model(model)
    except KeyError:
        enc = tiktoken.get_encoding("cl100k_base")
    return len(enc.encode(text))

INPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "input")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "output")

TRAINER_LEAVE_FILE = "trainer_leave_dates_2026.xlsx"
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


def _load_trainer_leave_dates() -> tuple[list[dict], str]:
    """Load trainer leave dates from trainer_leave_dates_2026.xlsx."""
    path = os.path.join(INPUT_DIR, TRAINER_LEAVE_FILE)
    df = pd.read_excel(path)

    trainers = []
    all_leave_dates = []
    for _, row in df.iterrows():
        trainer_name = row.get("Trainer Name")
        if pd.isna(trainer_name):
            continue

        raw_leave_dates = row.get("Leave Dates")
        leave_dates = []
        if not pd.isna(raw_leave_dates):
            for raw_date in str(raw_leave_dates).split(","):
                parsed_date = pd.to_datetime(raw_date.strip(), errors="coerce")
                if not pd.isna(parsed_date):
                    leave_dates.append(parsed_date.strftime("%Y-%m-%d"))

        all_leave_dates.extend(leave_dates)
        trainers.append(
            {
                "trainer": str(trainer_name).strip(),
                "leave_dates": sorted(set(leave_dates)),
            }
        )

    years = sorted({pd.to_datetime(date).year for date in all_leave_dates})
    if len(years) == 1:
        period = f"{years[0]}-01-01 to {years[0]}-12-31"
    elif years:
        period = f"{min(all_leave_dates)} to {max(all_leave_dates)}"
    else:
        period = ""

    return trainers, period


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


def _priority_trainer_names(priority_df: pd.DataFrame) -> set[str]:
    priority_columns = [
        "priority_1",
        "priority_2",
        "priority_3",
        "priority_4",
        "priority_5",
    ]
    trainer_names = set()
    for column in priority_columns:
        trainer_names.update(
            str(name).strip()
            for name in priority_df[column].dropna().tolist()
            if str(name).strip()
        )
    return trainer_names


def _normalize_trainer_name(name: str) -> str:
    return " ".join(str(name).strip().lower().split())


def _trainer_name_aliases(name: str) -> set[str]:
    aliases = {_normalize_trainer_name(name)}
    if "," in str(name):
        last_name, first_names = [part.strip() for part in str(name).split(",", 1)]
        aliases.add(_normalize_trainer_name(f"{first_names} {last_name}"))
        last_name_parts = last_name.split()
        if last_name_parts:
            aliases.add(_normalize_trainer_name(f"{first_names} {last_name_parts[0]}"))
    return aliases


def _build_leave_lookup(trainers: list[dict]) -> dict[str, dict]:
    leave_lookup = {}
    for trainer in trainers:
        for alias in _trainer_name_aliases(trainer["trainer"]):
            leave_lookup[alias] = trainer
    return leave_lookup


def plan_courses(session_data: dict) -> dict:
    """Main planning function:
    1. Use pre-computed session requirements
    2. Load trainer availability and priority
    3. Use LLM to assign trainers and dates
    4. Save result to Excel
    """
    try:
        # Step 1: Use session requirements from caller
        courses = session_data["courses"]
        print(f"[PLANNER] Step 1: Got {len(courses)} courses")

        # Step 2: Load trainer data
        print("[PLANNER] Step 2: Loading trainer data...")
        trainers, period = _load_trainer_leave_dates()
        priority_df = _load_trainer_priority()
        print(f"[PLANNER] Loaded leave dates for {len(trainers)} trainers, period: {period}")

        # Step 3: Build prompt context
        print("[PLANNER] Step 3: Building prompt and invoking LLM...")
        priority_records = priority_df.to_dict(orient="records")
        leave_lookup = _build_leave_lookup(trainers)
        trainer_leave_summary = {}
        for trainer_name in _priority_trainer_names(priority_df):
            leave_record = leave_lookup.get(_normalize_trainer_name(trainer_name))
            trainer_leave_summary.setdefault(
                trainer_name,
                {
                    "trainer": trainer_name,
                    "leave_dates": leave_record["leave_dates"] if leave_record else [],
                },
            )
        for trainer in trainers:
            trainer_leave_summary.setdefault(trainer["trainer"], trainer)
        trainer_leave_records = sorted(
            trainer_leave_summary.values(), key=lambda trainer: trainer["trainer"]
        )

        prompt = get_course_planner_prompt()

        llm = get_llm_from_env()
        structured_llm = llm.with_structured_output(CoursePlan)

        chain = prompt | structured_llm

        invoke_input = {
            "period": period,
            "courses_json": json.dumps(courses, indent=2),
            "priority_json": json.dumps(priority_records, indent=2, default=str),
            "trainer_leave_json": json.dumps(
                trainer_leave_records, indent=2
            ),
        }

        # Estimate and print token count
        formatted_messages = prompt.format_messages(**invoke_input)
        full_prompt_text = "\n".join(m.content for m in formatted_messages)
        token_count = _count_tokens(full_prompt_text)
        print(f"[PLANNER] Estimated prompt tokens: {token_count}")
        print(f"[PLANNER] Prompt character count: {len(full_prompt_text)}")

        print("[PLANNER] Invoking LLM (timeout=300s)...")
        start_time = time.time()
        result: CoursePlan = chain.invoke(invoke_input)
        elapsed = time.time() - start_time
        print(f"[PLANNER] LLM returned {len(result.sessions)} sessions in {elapsed:.1f}s")

        # Step 4: Save to Excel
        print("[PLANNER] Step 4: Saving to Excel...")
        rows = [session.model_dump() for session in result.sessions]
        result_df = pd.DataFrame(rows)

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        output_path = os.path.join(OUTPUT_DIR, "course_plan.xlsx")
        result_df.to_excel(output_path, index=False)
        print(f"[PLANNER] Saved to {output_path}")

        return {
            "total_sessions_planned": len(result.sessions),
            "output_file": "course_plan.xlsx",
            "plan": rows,
        }
    except Exception as e:
        print(f"[PLANNER] ERROR: {e}")
        print(f"[PLANNER] Traceback: {traceback.format_exc()}")
        raise
