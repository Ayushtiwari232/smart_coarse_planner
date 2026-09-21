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

## Planning rules

- Sessions are planned from September through December 2026.
- Courses are scheduled on weekdays only.
- A trainer cannot be assigned to overlapping courses.
- Trainer priority is applied before selecting another available trainer.
- Trainer leave dates are excluded.
- The number of sessions is calculated as $ceil(\text{students interested} / \text{maximum capacity})$.
- Three- and four-day sessions must not cross a weekend.

## Local development

Install dependencies from the deployment folder and run Azure Functions Core Tools:

```powershell
cd smart_coarse_planner
pip install -r requirements.txt
func start
```

Configure the required Azure OpenAI settings in `local.settings.json` before running the planner locally.