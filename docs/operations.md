# Operations and maintenance

## Deliberate compatibility changes

There is now one planner package and one Azure Functions v2 host. The old
deployment package and Cleveland fork have been removed, not independently maintained.

- Best demand now uses exactly BEST, as approved. The old root implementation
  actually selected PHC/CL/SLC; that mixed population is intentionally not retained.
- AI-first remains the default. Deterministic mode is explicit; dates/trainers can
  differ between modes. AI output is never guaranteed identical between calls.
- Every required schedulable session is validated, not only mandatory minimums.
- Required report failures now fail a run instead of returning incomplete success.
- All locations receive common audit sheets and explicit unscheduled reasons.
- Azure's anonymous endpoint has been replaced with Function-key authorization.
- Status names use `completed`, `completed_with_warnings`, `failed`, `processing`
  and preflight `validated`; update Power Automate conditions accordingly.
- Stable trainer IDs replace comma-splitting names. The same scheduler and
  report views are used for Cleveland and Singapore.
- Lab is the resource name throughout: `lab_compatibility.xlsx`, the `lab` input
  column and `lab_room` output field replace the old bay-named equivalents.
  Update external Excel consumers and SharePoint reference files accordingly.
- New run folders are readable UTC timestamps with microseconds; API job IDs
  remain internal metadata, not folder names. Existing API job IDs remain
  supported after local folder migration.

## Data retention and reproducibility

Authoritative inputs: `data\quarters\<quarter>`. Each run preserves input bytes,
SHA-256 hashes, application version, filters, mode and effective rules. These run
snapshots are deliberate audit records, not competing authoritative data copies.
The imported quarter also retains source-file hashes in `import_manifest.json`.
Run snapshots are flat: `runs\<quarter>\<location>\<UTC timestamp>\inputs`
contains the location files and `demand.xlsx` directly, with no repeated hierarchy.
Atomic directory creation prevents concurrent runs from overwriting one another.
The approved legacy cleanup retains the latest successful run per location and
removes older runs, legacy source archives and pre-migration backups. Removed
runs are no longer available through the status/download endpoints. Subsequent
runs retain history normally; no automatic deletion policy is enabled.

Do not use operational trainer/student data as committed fixtures. Store the
ignored data and run trees in approved backed-up storage with access controls and
a retention policy appropriate to your organization. Never remove an entire
data/run root as a cleanup shortcut. Remove only reviewed individual expired runs.

Legacy source/output archives and pre-migration backups were removed with explicit
approval after verifying current input completeness. This does not resolve the
outstanding source-data assumptions documented in `inputs.md`; original source-only
details are no longer retained locally. Import hashes remain as provenance, not
recoverable source copies.
The legacy nested environment file is still preserved without reading or merging
its contents at `data\archive\local-secrets\smart_coarse_planner.env`. Credentials
were excluded from cleanup. This file remains ignored and is not used at runtime
or included in deployment packages.
Use `.venv` for all commands; the old `venv` environment was removed after the
retained environment passed dependency checks and the regression suite.

## Hosting limitations

`POST /plan` is synchronous. Long AI calls can exceed your ingress/Azure Functions
HTTP response timeout. The polling endpoints read persistent status, but they do
not turn this into a queued/durable job. A queue-backed worker/Durable Functions
deployment is still needed if your production request duration exceeds hosting
limits. Do not start unmanaged background threads in an HTTP Function.

Local run folders survive process restarts only if the underlying filesystem
survives. Configure `PLANNER_RUNS_ROOT` to a persistent shared volume for
multi-instance status/download access. Instance-local temporary storage is not
a distributed job store. Graph uploads preserve completed outputs but do not by
themselves back the status endpoints.

For Graph-disabled Azure deployments, mount/configure an accessible quarterly
`PLANNER_DATA_ROOT`; operational inputs are deliberately not packaged. With Graph
enabled, credentials and input/output root folders must be configured.

There are no shared fixed intermediate filenames. Each request allocates its run
ID before work starts, so different processes cannot overwrite each other's runs.
Failed runs retain status and snapshots; pending workbooks are not published as
successful results.

## Quality and dependency maintenance

`pyproject.toml` is the dependency source. `requirements.txt` is only the Azure
installation entry point. `constraints.txt` records the tested environment;
constraints do not install unrelated packages.

After deliberate dependency changes, install into a clean environment, run the
tests, regenerate constraints with `python tools\lock_dependencies.py`, and verify
both Windows development and Linux deployment. Review changes before committing.
The dev dependency versions are explicit. CI tests Python 3.12/3.13.

The extracted scheduling/search and workbook view code intentionally retains
legacy algorithms/styles. Newly separated boundaries permit incremental typing
and report simplification without mixing those changes with this migration.
The search is a bounded heuristic, not a proof of globally optimal utilization.

## Production setup checklist

- Review and version the new source, tests and docs; operational data remains ignored.
- The previously tracked machine-specific `local.settings.json` was untracked
  without deleting the local file; never add real credentials to an example.
- Confirm quarterly mock flags and assumptions; disable mock data in production.
- Provision Graph permissions or the approved local data mount.
- Configure persistent shared run storage and retention.
- Configure Function keys and update Power Automate request/status mappings.
- Deploy only after CI passes; smoke-test the packaged artifact in staging.
- Use a queue/durable job design before relying on long-running HTTP requests.
