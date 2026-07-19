from typing import Optional
import base64
from pathlib import Path

from fastapi import FastAPI
from pydantic import BaseModel

from helpers import get_plan_file_response, get_plan_status, run_local, start_plan_job


app = FastAPI(title="Smart Course Planner")

DEFAULT_INPUT_FILE = Path(__file__).resolve().parent / "data" / "input" / "srl_and_wl.xlsx"


class PlanRequest(BaseModel):
    file_name: Optional[str] = None
    file_content_base64: Optional[str] = None
    input: Optional[str] = None


@app.get("/health")
def health():
    return {
        "status": "ok",
        "message": "FastAPI is running",
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

    This endpoint returns quickly with job_id and poll_url.
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

    return start_plan_job(
        file_name=file_name,
        file_content_base64=file_content_base64,
        user_input=req.input,
    )


@app.get("/plan/{job_id}")
def get_plan_result(job_id: str):
    """
    Lightweight polling endpoint.

    Power Automate should poll this until:
      status != processing

    This endpoint intentionally does not return base64 file content.
    """
    return get_plan_status(job_id)


@app.get("/plan/{job_id}/file")
def get_plan_file(job_id: str):
    """
    Downloads the completed Excel output file.

    Power Automate should call this only after the poll endpoint returns success.
    """
    return get_plan_file_response(job_id)


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
        help="Path to the SRL/WL .xlsx file. Defaults to data/input/srl_and_wl.xlsx",
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
        run_local(
            input_file=args.input_file,
            user_input=args.input,
        )
    else:
        import uvicorn

        uvicorn.run(
            "main:app",
            host="0.0.0.0",
            port=args.port,
            reload=True,
            reload_excludes=["data/*"],
        )
