# Quarterly input contract

All location inputs live at `data\quarters\<YYYY-Qn>\<location>`. Only shared
demand is stored at `data\quarters\<YYYY-Qn>\shared\demand.xlsx`.
The same relative structure lives under the configured SharePoint input root.

XLSX files use a single header row, no merged headers, and no embedded titles.
Identifiers are text. Dates are ISO `YYYY-MM-DD` strings or Excel date cells.
Do not use display names as relational keys.

| File | Required columns | Meaning |
|---|---|---|
| courses.xlsx | course_code, course_title, no_of_days, max_capacity, mandatory_sessions, lab_days | Unique course codes; duration/capacity positive integers; mandatory non-negative integer. |
| trainers.xlsx | trainer_id, trainer | Unique stable IDs and unique display names. |
| qualifications.xlsx | course_code, trainer_id | One qualified trainer/course pair per row. |
| priorities.xlsx | course_code, trainer_id, rank | One relationship per row, integer rank 1–5. Multiple trainers can have the same rank. |
| availability.xlsx | trainer_id, date, available_fraction | Overrides to default full weekday availability. Zero means unavailable, 1 fully available; fractional days cannot host a full-day session and are marked part-time. |
| lab_compatibility.xlsx | course_code, lab | Allowed course/lab pairs, independent of historical bookings. |
| bookings.xlsx | lab, course_code, start_date, end_date | Fixed reservations, inclusive dates; no booking count is deducted from session requirements. |
| rules.json | See README | Notes, explicit restrictions, mock-data flag and documented assumptions. |

Optional course columns: `min_capacity`, `classroom`, `offered`, `notes`.
Minimum capacity, offered flags and course notes are informational in this release;
they are not new scheduling gates. This preserves current demand-session semantics.
`lab_days` is a comma-separated list of **working-day indexes** such as `11,12,13,14,15`.
An empty cell explicitly means no lab. No fuzzy course-code matching is performed.

The demand workbook must contain `Comp Short Ttl`, `Req Site Cd`,
`Local Course Code`, and `Course Title`. The parser searches the first 30 rows of
each sheet for these headers. Each filtered row counts as one request, preserving
the existing calculation; it does not invent a student deduplication key.

## Validation and incomplete data

- Missing files, missing columns, unknown trainer IDs, duplicate course IDs and
  invalid structured restrictions are errors.
- Incomplete duration/capacity for a requested or mandatory course is retained
  as unscheduled, with a reason; the planner does not invent a duration/capacity.
- Invalid details on unused courses are warnings.
- Missing qualifications or compatible labs skip the affected course.
- Availability outside the quarter is reported and ignored, never silently remapped.
- Normal planning never edits authoritative inputs or overwrites another quarter.
- The imported Singapore date correction is explicit migrated mock data, not a
  general date-remapping feature.

## Current location follow-up

### Best

- Confirm requested-site coding: the original import had no `BEST`/IXR rows.
  Q1 2027 now includes 40 synthetic requests marked `BEST-2027-Q1-DEVELOPMENT`
  in the shared workbook's `Mock Data` column. They produce six sessions:
  IGT2BL001 twice; IGT2VC016, IGT2IL023, IGT2VC012 and IGT2IL008 once each.
  Best's rules mark the quarter as mock; other sites' demand is unchanged.
  Replace these rows with confirmed BEST-coded requests before clearing the flag.
  The original pre-mock backup was removed during the approved legacy cleanup.
  Remove only the marked mock rows when supplying confirmed demand; retain other
  locations' requests.
- Confirm unmatched trainer aliases reported in quarterly assumptions; no guessed
  qualifications or availability have been assigned to them.
- Confirm the retained IGT2BL001 **15 working-day** correction; its original sheet
  says 19. Duplicate legacy course rows were consolidated by maximum
  duration/capacity, matching the prior calculator.
- Provide authoritative lab-day/room compatibility and fixed room bookings if
  these must be hard scheduling constraints. Current legacy Best classroom
  assignments remain reporting metadata, not newly inferred constraints.
- Confirm quarterly mandatory counts; migrated missing values are zero.

### Cleveland

- Confirm the current qualification roster, mandatory counts and lab requirements.
- Fill missing schedule entries listed in Session Requirements.
- Supply compatible labs for courses reported with missing lab compatibility.
- Confirm whether lab bookings also represent fixed trainer teaching commitments;
  currently they block only labs, preserving the existing rule.
- Supply a confirmed quarterly availability list rather than changing older quarters.

### Singapore

- Confirm actual Q1 2027 holidays: the supplied 2026 month/day values were mapped
  into 2027 only as the approved mock data.
- Confirm Mandatory counts and explicit lab-day requirements; current zero/inferred
  values remain marked as assumptions.
- Verify `IGT2CV001` in the lab file versus `IGT2VC001` in the schedule. They are not
  automatically treated as the same course.
- Confirm blend-course lab requirements for IGT2BL001 rather than assuming a lab.
- Supply missing durations/capacities for the unused incomplete courses.
- Add Malleswari Maddineni only with a confirmed trainer ID, qualifications,
  priorities and availability. Holiday presence alone does not establish eligibility.

## New quarter checklist

1. Use `planner init-quarter`, which refuses to overwrite an existing location.
2. Replace shared demand; do not copy old demand as if it were current.
3. Enter availability and fixed bookings.
4. Review courses, qualifications, priorities and compatibility.
5. Add location notes and structured date restrictions.
6. Remove unconfirmed assumptions and clear `mock_data` only after review.
7. Run `planner validate`, inspect warnings/skipped courses, then generate a plan.
