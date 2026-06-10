from langchain_core.prompts import ChatPromptTemplate


COURSE_PLANNER_SYSTEM_PROMPT = """You are a training course planner. Your task is to schedule course sessions
and assign trainers based on the following rules:

1. Each course needs a specific number of sessions (provided below).
2. Each session spans a number of consecutive working days (no_of_days provided).
3. Trainers must be assigned based on priority order (priority_1 is most preferred,
   then priority_2, etc.). Use the highest priority trainer who is available.
4. A trainer can only teach one course at a time - no overlapping assignments.
5. Sessions should be scheduled within the planning period: {period}.
6. Only schedule on weekdays (Monday-Friday).
7. Do not assign a trainer on any of their leave dates.
8. If a course has no matching trainer in the priority list, pick the best available
    trainer from the trainer list who is not on leave and not otherwise occupied.
9. Spread sessions across the planning period when possible.
10. If a course is not found in the schedule (sessions_needed is null), skip it."""


COURSE_PLANNER_HUMAN_PROMPT = """Here are the courses that need sessions planned:

{courses_json}

Here is the trainer priority for each course (priority_1 is most preferred):

{priority_json}

Here are trainer leave dates. A trainer is unavailable on these dates:

{trainer_leave_json}

Please create a complete course plan assigning dates and trainers to each session."""


def get_course_planner_prompt() -> ChatPromptTemplate:
    return ChatPromptTemplate.from_messages(
        [
            ("system", COURSE_PLANNER_SYSTEM_PROMPT),
            ("human", COURSE_PLANNER_HUMAN_PROMPT),
        ]
    )