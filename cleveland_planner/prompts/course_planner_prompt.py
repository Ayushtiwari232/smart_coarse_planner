from langchain_core.prompts import ChatPromptTemplate


COURSE_PLANNER_SYSTEM_PROMPT = """You are a training course planner. Your task is to schedule course sessions
and assign trainers based on the following rules:

1. sessions_needed is the number of trainer-assigned sessions to schedule.
   Existing bay bookings are only blocked resources; do not count them as classes.
   A mandatory_sessions minimum applies even with no student demand.
   Prioritize meeting it before other demand. Do not schedule
   additional sessions when sessions_needed is zero or null.
2. Each session spans a number of consecutive working days (no_of_days provided).
3. Balance trainer utilization percentages (teaching working days divided by
   available working days in the planning quarter, excluding leave). Priority
   cells contain lists of individual trainers at the same rank. Priority is
   only a tie-breaker after workload balance. If a trainer is occupied, assign
   another available qualified roster trainer regardless of priority.
4. A trainer can only teach one course at a time - no overlapping assignments.
5. Sessions should be scheduled within the planning period: {period}.
6. Only schedule on weekdays (Monday-Friday).
7. Do not assign a trainer on any of their leave dates.
8. Assign only trainers in the Cleveland roster who are qualified for the course.
   If a course has no matching trainer in the priority list, pick an available
   qualified roster trainer who is not on leave and not otherwise occupied.
9. lab_days contains the 1-based working-day numbers that require a bay. For
   every such day a single compatible Cleveland bay must be free throughout all
   lab_days of the session. Existing bookings and other new sessions block bays.
   An empty lab_days list means no bay is required.
10. Spread sessions across the planning period when possible.
11. If a course is not found in the schedule (sessions_needed is null), skip it.
12. WEEK-WRAPPING RULE: If a session is 3 or 4 working days long, it MUST fit
    entirely within a single Monday-Friday week. Concretely:
      - A 3-day session must start on Monday, Tuesday, or Wednesday.
      - A 4-day session must start on Monday or Tuesday.
    Never let a 3-4 day session cross a weekend (Saturday/Sunday).
    Sessions of 1-2 days and sessions of 5+ days are not constrained by this
    rule, but still skip weekends (rule 6)."""


COURSE_PLANNER_HUMAN_PROMPT = """Here are the courses that need sessions planned:

{courses_json}

Here are trainer priority groups for each course (use only as workload tie-breakers):

{priority_json}

Here are the Cleveland trainer qualifications. Assign only a trainer listed for that course:

{qualification_json}

Here are the Cleveland bay resources compatible with each course and their fixed bookings:

{bay_occupancy_json}

Here are trainer leave dates. A trainer is unavailable on these dates:

{trainer_leave_json}

Previous validation feedback (fix these issues when present):
{validation_feedback}

Please create a complete course plan assigning dates and trainers to each session."""


def get_course_planner_prompt() -> ChatPromptTemplate:
    return ChatPromptTemplate.from_messages(
        [
            ("system", COURSE_PLANNER_SYSTEM_PROMPT),
            ("human", COURSE_PLANNER_HUMAN_PROMPT),
        ]
    )