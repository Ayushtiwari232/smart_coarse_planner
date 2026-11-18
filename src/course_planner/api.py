"""HTTP adapter; both API hosts call the same synchronous pipeline."""

import base64
import logging
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse

from course_planner import __version__
from course_planner.models import PlanRequest
from course_planner.pipeline import get_run, run_plan
from course_planner.settings import get_settings

app = FastAPI(title="Smart Course Planner", version=__version__)


@app.get("/health")
def health() -> dict:
    return {"status": "ok", "version": __version__}


def create_plan(request: PlanRequest) -> dict:
    result = run_plan(request, get_settings())
    return {
        **result,
        "file_content_base64": base64.b64encode(Path(result["output_file"]).read_bytes()).decode(
            "ascii"
        ),
    }


@app.post("/plan")
def plan(request: PlanRequest) -> dict:
    try:
        return create_plan(request)
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error
    except Exception as error:
        logging.getLogger(__name__).exception("API planning failed")
        raise HTTPException(
            status_code=500, detail="Planning failed; inspect server logs"
        ) from error


@app.get("/plan/{job_id}")
def get_plan_result(job_id: str) -> dict:
    try:
        return get_run(get_settings(), job_id)
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error
    except FileNotFoundError as error:
        raise HTTPException(status_code=404, detail=str(error)) from error


@app.get("/plan/{job_id}/file")
def get_plan_file(job_id: str):
    result = get_plan_result(job_id)
    if result["status"] not in {"completed", "completed_with_warnings"}:
        raise HTTPException(status_code=409, detail="Plan is not complete")
    path = Path(result["output_file"])
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Plan workbook not found")
    return FileResponse(path, filename=result["file_name"])
