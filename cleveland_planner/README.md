# Smart Course Planner

Smart Course Planner automates the creation of a training course plan from an SRL/WL Excel file. It filters the incoming requests, calculates the number of required course sessions, assigns available trainers, and generates a course-plan workbook.

The planning API is hosted in an Azure Function App and is called by a Power Automate flow.

## Planning workflow

1. A user manually triggers the Power Automate flow.
2. Power Automate gets the SRL/WL Excel file.
3. The flow sends the file to the Azure Function Plan API through an HTTP POST request.
4. The Azure Function processes the planning request and generates the course-plan workbook.
5. The API response includes the generated workbook as Base64 content.
6. Power Automate uses the returned content in its Create file action.

## What the planner uses

The planner combines these inputs:

- SRL/WL training requests
- Course capacity and course duration
- Trainer priority per course
- Trainer leave dates and availability
- Trainer location and classroom assignment

The current filter selects `MODALITY IXR` requests for the configured sites.

## What the planner produces

The generated Excel workbook includes:

- Course Plan: course sessions, dates, trainers, locations, and training rooms
- Trainer Availability: teaching assignments, leave dates, weekends, and part-time indicators
- Lab Availability: planned room usage
- Course Overview: training demand and planned sessions
- Utilization Metrics: trainer and lab utilization

## Azure Function API

### `POST /plan`

Starts and completes a course-planning request.

The repository configuration sets the Azure Functions HTTP `routePrefix` to an empty string, so the literal local route is `/plan`. Use the deployed Function App URL configured for the environment.

#### Request body

```json
{
  "file_name": "srl_and_wl.xlsx",
  "file_content_base64": "<Base64 encoded Excel file>",
  "input": "optional identifier or user input"
}
```

- `file_name`: Name of the uploaded Excel workbook.
- `file_content_base64`: Base64-encoded content from the Power Automate Get file content action.
- `input`: Optional value used in generated intermediate-file naming.

#### Successful response

```json
{
  "status": "success",
  "job_id": "abc123def456",
  "file_name": "processed_20260908_120000_course_plan_v3.xlsx",
  "file_content_base64": "<Base64 encoded course-plan workbook>",
  "message": "Plan completed successfully"
}
```

Use `file_content_base64` as the file content in Power Automate's Create file action. The response also includes a `job_id` for traceability.

## Cleveland planning rules

- The Cleveland planner uses Q1 2027, independently of the main planner.
- All Cleveland planning inputs are read from
  `cleveland_planner/data/input/Cleveland Documents`:
  - `Course schedule days Cleveland.xlsx` for course durations, capacities,
    lab requirements and mandatory counts.
  - `Asset Bay Course Occupancy Cleveland.xlsx`, `Training Priority Matrix Cleveland.xlsx`
    and `Cleveland Trainer Course Qualification Matrix.xlsx` for bays and trainers.
  - `trainers_cleveland_q1_2027.json` for the Cleveland trainer roster and leave,
    and `trainers.json` for part-time indicators in the availability sheet.
  - `srl_and_wl.xlsx` as the local/FastAPI default demand file. Azure Functions
    uses `New SRL and WL for Students_51_9219480390659314172.xlsx` by default.
    An uploaded file or `--input-file` overrides the respective default.
- Set `CLEVELAND_INPUT_DIR` to the **folder itself** to use a different input
  folder. This replaces the previous `CLEVELAND_SCHEDULE_FILE` override, which
  must be unset. All listed files must be present in the selected folder.
- Older copies outside this folder are retained but are not used by the
  Cleveland planning pipeline.
- Mandatory is a minimum number of course sessions per quarter, including when
  there are no SRL requests. Total required sessions are the greater of the
  mandatory minimum and demand divided by capacity, rounded up.
- Existing bay bookings only block bays; they do not count as trainer-assigned
  course sessions because their full course dates and trainers are unknown.
- Mandatory values are non-negative whole numbers, with zero meaning no
  mandatory sessions. Blank values are development mock data meaning zero.
  Blank Lab required values are mock data: reserve the
  entire course if the booking workbook lists a compatible bay for that course,
  otherwise assume no lab. Replace these assumptions with confirmed source data.
- "last five days" and "on the trhird day" reserve only those working days;
  "Virtual" needs no bay. Existing bay rows identify compatible bays and fixed
  bookings, and new sessions cannot overlap either fixed or newly assigned labs.
- Session Requirements is a sheet in the output workbook, not a separate file.
  Course Overview shows mandatory and planned sessions; Existing Bay Bookings
  and Lab Availability show the reservations used. Missing schedule entries are
  listed in Session Requirements with an empty sessions-needed value.
- Courses without a qualified Cleveland roster trainer are skipped, not assigned
  to an unqualified trainer. Course Overview and Session Requirements retain
  their requirements and show the reason; the API response also lists skipped
  courses. Update the qualification matrix and rerun to schedule them.
- The AI-generated schedule is validated before saving. If dates, trainers,
  mandatory counts or lab assignments are invalid, a deterministic search repairs
  the plan by scheduling dates, qualified trainers and bays together. It reserves
  only the specified lab days, preserves fixed bay bookings and schedules all
  calculated requirements for courses that have qualified trainers. The search
  backtracks when a choice blocks another course and stops after 50,000 candidate
  assignments with an explicit error if no complete plan has been established.
  It does not keep retrying invalid dates through the AI or bypass conflicts.
- Courses are scheduled on weekdays only.
- A trainer cannot be assigned to overlapping courses.
- Comma-separated names in a trainer priority cell are individual trainers at
  the same rank, not one combined trainer name.
- Trainer assignments balance utilization (teaching working days / available
  working days excluding leave). Priority is only a tie-breaker; any available,
  qualified Cleveland roster trainer can take over for an occupied trainer.
  After the AI proposes dates, code finds conflict-free assignments and improves
  utilization variance using single-session moves and pairwise swaps. Exact
  equality is not guaranteed because qualifications, leave and fixed dates can
  limit feasible assignments, and sessions cannot be split between trainers.
- Trainer leave dates are excluded.
- Three- and four-day sessions must not cross a weekend.

## Local development

Install dependencies from the deployment folder and run Azure Functions Core Tools:

```powershell
cd cleveland_planner
pip install -r requirements.txt
func start
```

Configure the required Azure OpenAI settings in `local.settings.json` before running the planner locally.

Run focused tests from the Cleveland planner directory with
`python -m unittest discover -s tests`.

The local FastAPI entry point can also be started with `python main.py`. Its
`POST /plan` is synchronous and returns the completed workbook. The optional
`GET /plan/{job_id}` and `GET /plan/{job_id}/file` endpoints retain the latest
20 successful results in process memory; they are cleared on restart.