# Smart Course Planner

One Python package for **Best**, **Cleveland**, and **Singapore**. Local commands,
FastAPI, and Azure Functions use the same quarterly inputs, scheduling engine, and
Excel reports. SharePoint/Graph is implemented but **disabled by default**.

## Start here (PowerShell, repository root)

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -c constraints.txt -e ".[dev]"
Copy-Item .env.example .env   # Only if .env does not already exist
.\.venv\Scripts\planner.exe validate --location cleveland --quarter 2027-Q1
.\.venv\Scripts\planner.exe plan --location cleveland --quarter 2027-Q1 --mode deterministic
```

Use a fresh virtual environment for the standardized package rather than reusing
the legacy environment. After cloning, restore approved quarterly datasets into
the ignored operational data folders before planning.

Default scheduling is **AI-first**. Omit `--mode deterministic` to use the Azure
OpenAI settings listed in [.env.example](.env.example). Deterministic mode makes
no model calls. Neither mode relaxes qualifications, trainer availability, lab
conflicts, required session counts, or quarter boundaries. A valid AI proposal
keeps its dates; invalid proposals are explicitly reported and repaired.

**Important current-data result:** the approved Best filter is now exactly
`Req Site Cd = BEST`, not the previous mixed-site filter. The imported SRL file
originally contained **no BEST / MODALITY IXR demand**. For Q1 2027 development,
40 clearly marked synthetic Best requests now produce **six sessions across five
courses**. Best is flagged as mock data; existing Cleveland/Singapore demand is
unchanged. Replace the `BEST-2027-Q1-DEVELOPMENT` rows in shared demand with
confirmed requests before production. The original workbook and Best rules are
no longer retained after the approved legacy cleanup; the marked mock rows can
be identified without changing any other location's requests.

## Structure

```text
src\course_planner\      Shared package: inputs, scheduling, reports, CLI and API
tests\                  Synthetic fixtures, scheduling and adapter regressions
tools\                  Dependency constraints maintenance
docs\                   Input schemas, operations and SharePoint setup
data\quarters\2027-Q1\
  shared\demand.xlsx    One authoritative quarterly SRL/WL workbook
  best\                Standardized location inputs
  cleveland\
  singapore\
runs\<quarter>\<location>\<UTC timestamp>\
  inputs\              Flat immutable snapshot: location files + demand.xlsx
  manifest.json        Checksums, rules, filters, application version
  status.json          Persistent local status
  <location>_Course_Plan_<quarter>.xlsx
```

For example, `runs\2027-Q1\best\2026-11-18_142029_123456`.
Timestamps are UTC; microseconds prevent collisions during simultaneous requests.
No opaque ID directories or repeated quarter/location folders inside snapshots.
Each rerun keeps its own history rather than overwriting an earlier workbook.

Operational data and run outputs are deliberately ignored by Git.
Tests create small synthetic workbooks and do not require company datasets or
credentials. Preserve operational folders in your approved backup/storage system.

## Locations and quarters

| Location | Exact SRL site | Modality |
|---|---|---|
| Best | BEST | MODALITY IXR |
| Cleveland | CL | MODALITY IXR |
| Singapore | SLC | MODALITY IXR |

Demand is scoped by **requested training site**, not student country.

To prepare a new quarter:

```powershell
.\.venv\Scripts\planner.exe init-quarter --location singapore --quarter 2027-Q2 --from-quarter 2027-Q1
```

This refuses to overwrite an existing quarter/location, copies reference tables,
clears holidays/bookings, and marks the result **unconfirmed/mock**. It does not
silently reuse last quarter's demand, holiday dates or bookings. Supply the new
quarter's shared demand and location availability, verify all reference data,
then clear the mock flag. Set `PLANNER_ALLOW_MOCK_DATA=false` in production.

See [input schemas and outstanding location data](docs/inputs.md).

## Notes and scheduling restrictions

Each location/quarter has a `rules.json`:

```json
{
  "notes": ["Discuss course timing with the local planner."],
  "skip_first_week": false,
  "excluded_dates": [],
  "mock_data": false,
  "assumptions": []
}
```

Notes are advisory and appear in **Notes and Assumptions**. They are not interpreted
as executable instructions. `skip_first_week=true` excludes the quarter's first
seven calendar dates; `excluded_dates` accepts explicit ISO dates in that quarter.
These restrictions are disabled in the imported data. They apply to every trainer,
including AI proposals, deterministic scheduling and utilization denominators.

## Output

One workbook contains the location's course plan, **Session Requirements**,
Trainer Availability, Course Overview, Lab Availability and Utilization Metrics.
The final three tabs are **Existing Lab Bookings**, **Notes and Assumptions**,
and **Run Manifest**, in that order. Lab is the consistent resource name in
inputs, scheduling, reports and messages.

- Mandatory sessions apply even without demand: `max(mandatory, ceil(demand/capacity))`.
- Existing lab bookings block labs; they do not count as trainer-assigned classes.
- Missing qualifications, course details or lab compatibility produce explicitly
  unscheduled course rows. Unknown demand quantities remain unknown, not zero.
- Impossible complete resource schedules fail explicitly; constraints are not
  bypassed. The existing bounded backtracking search is retained.
- Cleveland/Singapore balance utilization; priority breaks ties. Best keeps
  priority-first AI instructions. Deterministic mode is an opt-in alternative.
- Exact equality of trainer utilization is not guaranteed by qualifications,
  leave, indivisible sessions or this heuristic.
- Model output is not byte-for-byte reproducible. Filters, corrected validation,
  explicit skip reasons and additional audit sheets are intentional changes.

## API

Local development:

```powershell
.\.venv\Scripts\python.exe main.py
```

Or `uvicorn course_planner.api:app --port 8011 --env-file .env`.
Open `http://127.0.0.1:8011/docs`.

```json
{
  "location": "singapore",
  "quarter": "2027-Q1",
  "mode": "ai-first"
}
```

`POST /plan` is synchronous and returns the completed workbook as Base64, plus
`job_id`, warnings and skipped courses. The API identifier is retained only in
metadata, not used as a folder name. Optional `file_name`,
`file_content_base64` and legacy `input` fields are accepted; an upload overrides
**only demand**. The `input` field is retained for compatibility, not used as a path.
Result statuses are `completed`, `completed_with_warnings`, or `failed`;
preflight-only runs use `validated`.

`GET /plan/{job_id}` reads persisted status. `GET /plan/{job_id}/file` downloads a
completed workbook. Status is retained on the configured run filesystem, not in
process memory. This is **not a durable background-job queue**; see deployment
limitations in [operations](docs/operations.md).

Azure Functions uses the same routes with Function-key authorization. Update
Power Automate to pass `x-functions-key` and the new location/quarter fields.
No live deployment or credential provisioning is performed by this migration.

## Quality checks

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m ruff check src tests tools main.py function_app.py
.\.venv\Scripts\python.exe -m ruff format --check src tests tools main.py function_app.py
.\.venv\Scripts\python.exe -m pyright
```

Pyright currently gates the typed configuration/storage boundary; the extracted
dict-based legacy scheduler/reports are covered by regression tests and can be
typed incrementally. CI runs on Python 3.12 and 3.13, then packages **this same
source**, excluding operational data and credentials, before deployment.

See [Graph setup](docs/sharepoint.md) and [operational guidance](docs/operations.md).
