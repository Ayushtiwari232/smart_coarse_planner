"""Location-independent planning with immutable per-run input snapshots."""

import base64
import hashlib
import json
import logging
import re
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path

from course_planner import __version__
from course_planner.engine import SchedulingEngine
from course_planner.inputs import INPUT_FILES, SITES, load_inputs
from course_planner.models import PlanRequest, Settings
from course_planner.storage import GraphStorage
from course_planner.workbook import ReportWriter

logger = logging.getLogger(__name__)

RUN_TIMESTAMP = "%Y-%m-%d_%H%M%S_%f"
RUN_NAME_PATTERN = r"\d{4}-\d{2}-\d{2}_\d{6}_\d{6}"


def create_run_directory(settings: Settings, request: PlanRequest) -> Path:
    parent = settings.runs_root / request.quarter / request.location
    parent.mkdir(parents=True, exist_ok=True)
    for _ in range(10):
        directory = parent / datetime.now(timezone.utc).strftime(RUN_TIMESTAMP)
        try:
            directory.mkdir()
        except FileExistsError:
            continue
        return directory
    raise FileExistsError("Unable to allocate a unique timestamped run directory")


def write_json(path: Path, data: dict) -> None:
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")
    temporary.replace(path)


def run_plan(
    request: PlanRequest,
    settings: Settings,
    *,
    validate_only: bool = False,
    input_file: Path | None = None,
) -> dict:
    run_dir = create_run_directory(settings, request)
    run_id = uuid.uuid4().hex
    input_dir = run_dir / "inputs"
    input_dir.mkdir()
    status_path = run_dir / "status.json"
    status = {
        "job_id": run_id,
        "location": request.location,
        "quarter": request.quarter,
        "status": "processing",
    }
    write_json(status_path, status)
    try:
        graph = GraphStorage.from_env()
        if graph.enabled != settings.sharepoint_enabled:
            raise ValueError("SharePoint toggle differs between settings and environment")
        hashes = {}
        for name in (*INPUT_FILES, "demand.xlsx"):
            relative = f"{request.quarter}/{'shared' if name == 'demand.xlsx' else request.location}/{name}"
            target = input_dir / name
            if name == "demand.xlsx" and request.file_content_base64:
                try:
                    content = base64.b64decode(request.file_content_base64, validate=True)
                except ValueError as error:
                    raise ValueError("Invalid base64 demand workbook") from error
                target.write_bytes(content)
            elif name == "demand.xlsx" and input_file:
                shutil.copyfile(input_file, target)
            elif graph.enabled:
                target.write_bytes(graph.download(relative))
            else:
                source = settings.data_root / "quarters" / relative
                if not source.is_file():
                    raise ValueError(f"Missing quarterly input: {relative}")
                shutil.copyfile(source, target)
            hashes[relative] = hashlib.sha256(target.read_bytes()).hexdigest()
        data = load_inputs(
            settings.data_root, request.location, request.quarter, snapshot_dir=input_dir
        )
        if data.rules.mock_data and not settings.allow_mock_data:
            raise ValueError(
                "Quarter is marked as mock data; confirm inputs or explicitly allow development mock data"
            )
        manifest = {
            "schema_version": 2,
            "run_id": run_id,
            "application_version": __version__,
            "location": request.location,
            "quarter": request.quarter,
            "mode": request.mode,
            "created_utc": datetime.now(timezone.utc).isoformat(),
            "source": "sharepoint" if graph.enabled else "local",
            "demand_filter": {"site": SITES[request.location], "modality": "MODALITY IXR"},
            "input_sha256": hashes,
            "rules": data.rules.model_dump(mode="json"),
        }
        write_json(run_dir / "manifest.json", manifest)
        skipped = [c for c in data.courses if c.get("unscheduled_reason")]
        schedulable = [c for c in data.courses if not c.get("unscheduled_reason")]
        if validate_only:
            result = {
                **status,
                "status": "validated",
                "demand_rows": len(data.filtered_demand),
                "courses": data.courses,
                "warnings": data.warnings,
                "skipped_courses": skipped,
                "run_directory": str(run_dir.resolve()),
            }
            write_json(status_path, result)
            return result
        engine = SchedulingEngine(data.start, data.end, data.compatibility)
        rows = []
        if schedulable:
            if request.mode == "ai-first":
                from course_planner.llm import propose

                rows = propose(data, schedulable)
                try:
                    if request.location != "best":
                        engine._balance_trainer_assignments(
                            rows, data.trainers, data.qualifications, data.priorities
                        )
                    engine._validate_sessions(rows, schedulable, data.trainers, data.qualifications)
                    engine._assign_lab_rooms(rows, data.bookings, schedulable)
                except ValueError as error:
                    data.warnings.append(
                        f"AI proposal failed validation; deterministic repair used: {error}"
                    )
                    logger.warning("run=%s invalid AI proposal: %s", run_id, error)
                    rows = engine._schedule_resource_feasible_sessions(
                        schedulable, data.trainers, data.qualifications, data.bookings
                    )
            else:
                rows = engine._schedule_resource_feasible_sessions(
                    schedulable, data.trainers, data.qualifications, data.bookings
                )
            if request.location != "best":
                engine._balance_trainer_assignments(
                    rows, data.trainers, data.qualifications, data.priorities
                )
        engine._validate_sessions(rows, schedulable, data.trainers, data.qualifications)
        engine._assign_lab_rooms(rows, data.bookings, schedulable)
        filename = f"{data.location}_Course_Plan_{request.quarter.replace('-', '_')}.xlsx"
        temporary = run_dir / "pending.xlsx"
        ReportWriter(data).write(temporary, rows, manifest)
        output = run_dir / filename
        temporary.replace(output)
        if graph.enabled:
            remote = f"{request.quarter}/{request.location}/{run_dir.name}"
            graph.upload(f"{remote}/{filename}", output.read_bytes())
            graph.upload(f"{remote}/manifest.json", (run_dir / "manifest.json").read_bytes())
        result = {
            **status,
            "status": "completed_with_warnings" if skipped or data.warnings else "completed",
            "file_name": filename,
            "output_file": str(output.resolve()),
            "total_sessions_planned": len(rows),
            "demand_rows": len(data.filtered_demand),
            "warnings": data.warnings,
            "skipped_courses": skipped,
            "run_directory": str(run_dir.resolve()),
        }
        write_json(status_path, result)
        logger.info("run=%s completed sessions=%s", run_id, len(rows))
        return result
    except Exception:
        pending = run_dir / "pending.xlsx"
        if pending.is_file():
            pending.unlink()
        write_json(
            status_path,
            {
                **status,
                "status": "failed",
                "message": "Planning failed; inspect server logs or CLI error.",
            },
        )
        logger.exception("run=%s planning failed", run_id)
        raise


def get_run(settings: Settings, run_id: str) -> dict:
    if not re.fullmatch(rf"(?:[0-9a-f]{{32}}|{RUN_NAME_PATTERN})", run_id):
        raise ValueError("Invalid job ID")
    matches = list(settings.runs_root.glob(f"????-Q?/*/{run_id}/status.json"))
    if not matches and re.fullmatch(r"[0-9a-f]{32}", run_id):
        # Migrated runs retain their original API job IDs.
        matches = [
            path
            for path in settings.runs_root.glob("????-Q?/*/*/status.json")
            if re.fullmatch(RUN_NAME_PATTERN, path.parent.name)
            and json.loads(path.read_text(encoding="utf-8")).get("job_id") == run_id
        ]
    if len(matches) != 1:
        raise FileNotFoundError("Plan job not found")
    return json.loads(matches[0].read_text(encoding="utf-8"))
