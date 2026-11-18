# Microsoft Graph / SharePoint setup

The connector is implemented and tested with mocked HTTP. **No real tenant,
credentials or SharePoint site have been provisioned or contacted.**

## Enable later

Copy the setting names from `.env.example` into your local environment or Azure
Function App settings. Never commit actual credential values.

```dotenv
PLANNER_SHAREPOINT_ENABLED=false
GRAPH_TENANT_ID=
GRAPH_CLIENT_ID=
GRAPH_CLIENT_SECRET=
GRAPH_DRIVE_ID=
GRAPH_INPUT_ROOT=CoursePlanner/quarters
GRAPH_OUTPUT_ROOT=CoursePlanner/runs
```

With the toggle false, Graph credentials are not required and no Graph requests
are made. Set it to `true` only after the following setup.

1. Register an application in your organization's Microsoft Entra tenant.
2. Create and securely store a client secret; use environment/App Settings or an
   approved secret store reference, never a workbook.
3. Ask the administrator for least-privilege **application** permissions. Prefer
   `Sites.Selected` and an explicit write grant for the target SharePoint site,
   rather than tenant-wide access. Admin consent and the site grant are separate.
4. Obtain the document library's **drive ID**. This implementation operates on a
   configured drive and does not discover sites/libraries automatically.
5. Populate the input root using the canonical structure:

```text
CoursePlanner/quarters/2027-Q1/shared/demand.xlsx
CoursePlanner/quarters/2027-Q1/best/courses.xlsx
CoursePlanner/quarters/2027-Q1/best/trainers.xlsx
...the other required input files...
CoursePlanner/quarters/2027-Q1/cleveland/...
CoursePlanner/quarters/2027-Q1/singapore/...
```

6. Set the toggle true and run `planner validate` against a test site first.
7. Test a plan, output upload, permission failure, expired credentials and
   throttling in your real environment before production use.

## Flow

```text
Client-credentials token (.default Graph scope)
  -> download required location/quarter files + shared demand
  -> write immutable local run snapshot and SHA-256 manifest
  -> parse and validate
  -> calculate requirements, schedule and validate
  -> produce and verify complete workbook
  -> create output directory hierarchy
  -> upload <quarter>/<location>/<UTC timestamp>/<workbook>
  -> upload corresponding manifest.json
```

Outputs use the same readable UTC timestamp folders as local runs (for example
`2026-11-18_142029_123456`). Upload canonical `lab_compatibility.xlsx`
with columns `course_code, lab`; `bookings.xlsx` also uses `lab`, not the legacy
`bay` column. Existing remote input files must be migrated before enabling Graph.

Power Automate can alternatively keep downloading demand and pass it as Base64;
that overrides only the demand input. Other reference inputs still use the
selected source. There is no silent fallback to local reference data when an
enabled Graph request fails.

The adapter refreshes a rejected token once and bounds retries for 429/503,
respecting a bounded Retry-After delay. Requests have timeouts and errors do not
include tokens, client secrets or response bodies. Files use Graph simple upload
(maximum 250 MiB); larger files require an upload-session implementation.

Input files are snapshotted sequentially. Local checksums preserve exactly what
was used, but this is **not a transactional SharePoint multi-file snapshot**.
Publish completed quarterly folders as a unit or use an operational approval/
freeze process while runs are starting. Graph ETag concurrency/version-pinned
downloads are a future enhancement.

Uploads are also not transactional across files. A manifest upload failure marks
the run failed even if its workbook reached SharePoint. Consumers should require
both workbook and manifest before using a remote run.

Azure Functions uses Function-key auth independently of Graph's app credentials.
Keep both credentials separate. Configure environment variables in Azure;
local `.env` loading is provided by the CLI and `uvicorn --env-file .env`.
