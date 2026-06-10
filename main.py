from typing import Optional
import traceback
import base64
import os
import uuid
import zipfile
import threading
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI
from pydantic import BaseModel

app = FastAPI(title="Smart Course Planner")

_plan_lock = threading.Lock()
_jobs: dict[str, dict] = {}

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"


class PlanRequest(BaseModel):
    file_name: Optional[str] = None
    file_content_base64: Optional[str] = None
    input: Optional[str] = None


@app.get("/health")
def health():
    return {
        "status": "ok",
        "message": "FastAPI is running"
    }


def _run_plan(job_id: str, input_file: Path, user_input: Optional[str]):
    """Background worker that runs the full planning pipeline."""
    try:
        from services.filter_service import apply_filters
        from services.session_calculator_service import calculate_sessions
        from services.planner_service import plan_courses

        # Step 1: Apply filters
        print("[PLAN] Step 1: Applying filters...")
        filter_result = apply_filters(
            input_file=str(input_file),
            input=user_input
        )
        print("[PLAN] Step 1 done.")

        filtered_file = filter_result["output_file"]
        filtered_full_path = OUTPUT_DIR / filtered_file

        if not zipfile.is_zipfile(filtered_full_path):
            _jobs[job_id] = {
                "status": "error",
                "message": "Filtered output file is not a valid .xlsx file",
            }
            return

        # Step 2: Calculate sessions
        print(f"[PLAN] Step 2: Calculating sessions for {filtered_file}...")
        session_result = calculate_sessions(filtered_file=filtered_file)
        print("[PLAN] Step 2 done.")

        # Step 3: Plan courses
        print(f"[PLAN] Step 3: Planning courses...")
        plan_result = plan_courses(session_data=session_result)
        print("[PLAN] Step 3 done.")

        # Step 4: Read final Excel and encode as base64
        output_file_name = Path(plan_result["output_file"]).name
        plan_file_path = OUTPUT_DIR / output_file_name

        if not plan_file_path.exists():
            possible_path = Path(plan_result["output_file"])
            if possible_path.exists():
                plan_file_path = possible_path
            else:
                _jobs[job_id] = {
                    "status": "error",
                    "message": "Plan output file not found",
                }
                return

        with open(plan_file_path, "rb") as f:
            file_base64 = base64.b64encode(f.read()).decode("utf-8")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sharepoint_file_name = f"processed_{timestamp}_{plan_file_path.name}"

        print(f"[PLAN] Done. Output: {plan_file_path}")

        _jobs[job_id] = {
            "status": "success",
            "filter": filter_result,
            "sessions": session_result,
            "plan": plan_result,
            "file_name": sharepoint_file_name,
            "file_content_base64": file_base64,
        }

    except Exception as e:
        print(f"[PLAN] ERROR: {e}")
        _jobs[job_id] = {
            "status": "error",
            "error": str(e),
            "traceback": traceback.format_exc()
        }
    finally:
        _plan_lock.release()


@app.post("/plan")
def plan(req: PlanRequest):
    if not _plan_lock.acquire(blocking=False):
        print("[PLAN] Request rejected — already processing a plan")
        return {
            "status": "error",
            "message": "A plan request is already in progress. Please wait."
        }
    try:
        print(f"[PLAN] Request received: input={req.input}, file_name={req.file_name}")

        INPUT_DIR.mkdir(parents=True, exist_ok=True)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        input_file = None

        if req.file_content_base64:
            safe_file_name = Path(req.file_name or "srl_and_wl.xlsx").name
            input_file = INPUT_DIR / safe_file_name

            file_bytes = base64.b64decode(req.file_content_base64)

            if file_bytes[:4] != b'PK\x03\x04':
                try:
                    file_bytes = base64.b64decode(file_bytes)
                    print("[PLAN] Detected double base64 encoding, decoded again")
                except Exception:
                    pass

            input_file.write_bytes(file_bytes)
            print(f"[PLAN] Saved uploaded Excel: {input_file} ({input_file.stat().st_size} bytes)")

            if not zipfile.is_zipfile(input_file):
                _plan_lock.release()
                return {
                    "status": "error",
                    "message": "Received file is not a valid .xlsx file",
                    "file_name": safe_file_name,
                    "file_size": input_file.stat().st_size
                }
        else:
            input_file = INPUT_DIR / "srl_and_wl.xlsx"
            print(f"[PLAN] Using local file: {input_file}")

            if not input_file.exists():
                _plan_lock.release()
                return {
                    "status": "error",
                    "message": "No file uploaded and local fallback file not found",
                    "expected_file": str(input_file)
                }

        job_id = uuid.uuid4().hex[:12]
        _jobs[job_id] = {"status": "processing"}

        thread = threading.Thread(target=_run_plan, args=(job_id, input_file, req.input), daemon=True)
        thread.start()

        print(f"[PLAN] Job started: {job_id}")
        return {
            "status": "accepted",
            "job_id": job_id,
            "poll_url": f"/plan/{job_id}"
        }

    except Exception as e:
        _plan_lock.release()
        print(f"[PLAN] ERROR: {e}")
        return {
            "status": "error",
            "error": str(e),
        }


@app.get("/plan/{job_id}")
def get_plan_result(job_id: str):
    job = _jobs.get(job_id)
    if not job:
        return {"status": "error", "message": "Job not found"}
    return job


def run_local(input_file: Optional[str] = None, user_input: Optional[str] = None) -> dict:
    """Run the full planning pipeline locally without the HTTP layer.

    Args:
        input_file: Path to the SRL/WL .xlsx file. Defaults to
            data/input/srl_and_wl.xlsx if not provided.
        user_input: Optional input label passed to the filter step.
    """
    from services.filter_service import apply_filters
    from services.session_calculator_service import calculate_sessions
    from services.planner_service import plan_courses

    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    resolved = Path(input_file) if input_file else INPUT_DIR / "srl_and_wl.xlsx"
    if not resolved.exists():
        raise FileNotFoundError(f"Input file not found: {resolved}")

    print(f"[LOCAL] Using input file: {resolved}")

    print("[LOCAL] Step 1: Applying filters...")
    filter_result = apply_filters(input_file=str(resolved), input=user_input)
    print("[LOCAL] Step 1 done.")

    print(f"[LOCAL] Step 2: Calculating sessions for {filter_result['output_file']}...")
    session_result = calculate_sessions(filtered_file=filter_result["output_file"])
    print("[LOCAL] Step 2 done.")

    print("[LOCAL] Step 3: Planning courses...")
    plan_result = plan_courses(session_data=session_result)
    print("[LOCAL] Step 3 done.")

    print(f"[LOCAL] Output file: {OUTPUT_DIR / plan_result['output_file']}")
    return {
        "filter": filter_result,
        "sessions": session_result,
        "plan": plan_result,
    }


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
        help="Path to the SRL/WL .xlsx file (defaults to data/input/srl_and_wl.xlsx)",
    )
    parser.add_argument(
        "--input",
        default=None,
        help="Optional input label passed to the filter step",
    )
    args = parser.parse_args()

    if args.local:
        run_local(input_file=args.input_file, user_input=args.input)
    else:
        import uvicorn

        uvicorn.run(
            "main:app",
            host="0.0.0.0",
            port=8011,
            reload=True,
            reload_excludes=["data/*"]
        )