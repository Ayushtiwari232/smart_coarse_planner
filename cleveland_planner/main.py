from typing import Optional
import base64
import io
from collections import OrderedDict
from pathlib import Path
from threading import Lock

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from helpers import start_plan_job
from services.cleveland_data import INPUT_DIR


app = FastAPI(title="Smart Course Planner")
_results: OrderedDict[str, dict] = OrderedDict()
_results_lock = Lock()
_MAX_RESULTS = 20

APP_VERSION = "2026-07-20-v2"

DEFAULT_INPUT_FILE = INPUT_DIR / "srl_and_wl.xlsx"


class PlanRequest(BaseModel):
    file_name: Optional[str] = None
    file_content_base64: Optional[str] = None
    input: Optional[str] = None


@app.get("/health")
def health():
    return {
        "status": "ok",
        "message": "FastAPI is running",
        "version": APP_VERSION,
    }


@app.post("/plan")
def plan(req: Optional[PlanRequest] = None):
    """
    Starts a planning job.

    Power Automate should call this with:
    {
      "file_name": "...xlsx",
      "file_content_base64": "<base64 from Get file content $content>",
      "input": ""
    }

    If `file_content_base64` is not provided, the endpoint falls back to
    `data/input/srl_and_wl.xlsx`, base64-encodes it, and uses that.

    The response includes the completed workbook and a job_id.
    """
    if req is None:
        req = PlanRequest()
    file_name = req.file_name
    file_content_base64 = req.file_content_base64

    if not file_content_base64:
        if not DEFAULT_INPUT_FILE.exists():
            return {
                "status": "error",
                "message": (
                    "file_content_base64 not provided and default fallback file "
                    "not found"
                ),
                "expected_file": str(DEFAULT_INPUT_FILE),
            }

        print(
            f"[PLAN] file_content_base64 missing. "
            f"Loading and encoding fallback file: {DEFAULT_INPUT_FILE}"
        )
        file_bytes = DEFAULT_INPUT_FILE.read_bytes()
        file_content_base64 = base64.b64encode(file_bytes).decode("ascii")
        if not file_name:
            file_name = DEFAULT_INPUT_FILE.name

    result = start_plan_job(
        file_name=file_name,
        file_content_base64=file_content_base64,
        user_input=req.input,
    )
    if result.get("status") == "success":
        with _results_lock:
            _results[result["job_id"]] = result
            while len(_results) > _MAX_RESULTS:
                _results.popitem(last=False)
    return result


@app.get("/plan/{job_id}")
def get_plan_result(job_id: str):
    """
    Lightweight polling endpoint.

    Returns the completed result metadata without the workbook content.
    """
    with _results_lock:
        result = _results.get(job_id)
        if result is None:
            raise HTTPException(status_code=404, detail="Plan job not found")
        return {key: value for key, value in result.items() if key != "file_content_base64"}


@app.get("/plan/{job_id}/file")
def get_plan_file(job_id: str):
    """
    Downloads the completed Excel output file.

    Completed jobs are retained in memory until restart or cache eviction.
    """
    with _results_lock:
        result = _results.get(job_id)
        if result is None:
            raise HTTPException(status_code=404, detail="Plan job not found")
        content = base64.b64decode(result["file_content_base64"])
        file_name = result["file_name"]
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{file_name}"'},
    )


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Smart Course Planner")
    parser.add_argument(
        "--local",
        action="store_true",
        help="Run the planning pipeline locally instead of starting the server",
    )
    parser.add_argument(
        "--input-file",
        default=None,
        help="Path to the SRL/WL .xlsx file. Defaults to the Cleveland input folder.",
    )
    parser.add_argument(
        "--input",
        default=None,
        help="Optional input label passed to the filter step",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8011,
        help="Port for local FastAPI server",
    )

    args = parser.parse_args()

    if args.local:
        input_path = Path(args.input_file) if args.input_file else DEFAULT_INPUT_FILE
        response = plan(
            PlanRequest(
                file_name=input_path.name,
                file_content_base64=base64.b64encode(input_path.read_bytes()).decode("ascii"),
                input=args.input,
            )
        )
        if response.get("status") != "success":
            raise RuntimeError(response.get("message", "Planning failed"))
        print(f"Plan completed: {response['file_name']}")
        for course in response.get("skipped_courses", []):
            print(f"WARNING: {course['course_code']} - {course['reason']} "
                  f"({course['sessions_needed']} sessions unscheduled)")
    else:
        import uvicorn

        uvicorn.run(
            "main:app",
            host="0.0.0.0",
            port=args.port,
            reload=True,
            reload_excludes=["data/*"],
        )
