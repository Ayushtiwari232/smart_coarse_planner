from typing import Optional

from fastapi import FastAPI
from pydantic import BaseModel

from helpers import get_plan_file_response, get_plan_status, run_local, start_plan_job


app = FastAPI(title="Smart Course Planner")


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
def plan(req: PlanRequest):
    """
    Starts a planning job.

    Power Automate should call this with:
    {
      "file_name": "...xlsx",
      "file_content_base64": "<base64 from Get file content $content>",
      "input": ""
    }

    This endpoint returns quickly with job_id and poll_url.
    """
    return start_plan_job(
        file_name=req.file_name,
        file_content_base64=req.file_content_base64,
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
