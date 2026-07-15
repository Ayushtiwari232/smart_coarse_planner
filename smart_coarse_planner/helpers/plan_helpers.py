from typing import Optional, Dict, Any
import base64
import os
import tempfile
import threading
import traceback
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

from fastapi.responses import FileResponse


BASE_DIR = Path(__file__).resolve().parent.parent

# Static bundled input files (read-only)
INPUT_DIR = BASE_DIR / "data" / "input"

# Writable directories for Azure Functions
_TMP_BASE = Path(tempfile.gettempdir()) / "smart_planner"

UPLOAD_INPUT_DIR = _TMP_BASE / "input"
OUTPUT_DIR = Path(
    os.environ.get("SMART_PLANNER_OUTPUT_DIR")
    or (_TMP_BASE / "output")
)

UPLOAD_INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Global in-memory job store.
# Good for local/dev tunnel POC.
# For production Azure, replace this with Blob/Table Storage or SharePoint-backed status.
_jobs: Dict[str, Dict[str, Any]] = {}

# Single-job lock to avoid multiple plans running at the same time.
_plan_lock = threading.Lock()


def _safe_join_output(file_value: str) -> Path:
    """
    Resolve a service-returned output file safely.

    Services sometimes return:
    - just a file name, e.g. filtered_output.xlsx
    - relative path
    - absolute path

    This helper handles all three.
    """
    file_path = Path(file_value)

    if file_path.is_absolute():
        return file_path

    candidate_from_base = BASE_DIR / file_path
    if candidate_from_base.exists():
        return candidate_from_base

    return OUTPUT_DIR / file_path.name


def _validate_xlsx(path: Path) -> bool:
    """
    .xlsx files are zip files internally.
    This check catches bad/corrupted file content early.
    """
    return path.exists() and zipfile.is_zipfile(path)


def _decode_excel_base64(file_content_base64: str) -> bytes:
    """
    Decode Power Automate file content.

    SharePoint Get file content usually sends base64 from:
      body('Get_file_content')?['$content']

    Some flows accidentally double-encode the file.
    This tries one decode, then a second decode only if needed.
    """
    file_bytes = base64.b64decode(file_content_base64)

    if file_bytes[:4] == b"PK\x03\x04":
        return file_bytes

    try:
        second_decode = base64.b64decode(file_bytes)
        if second_decode[:4] == b"PK\x03\x04":
            print("[PLAN] Detected double base64 encoding, decoded again")
            return second_decode
    except Exception:
        pass

    return file_bytes


def _release_plan_lock() -> None:
    if _plan_lock.locked():
        _plan_lock.release()


def _run_plan(job_id: str, input_file: Path, user_input: Optional[str]) -> None:
    """
    Background worker that runs the full planning pipeline.
    """
    try:
        print(f"[PLAN:{job_id}] Background job started")
        print(f"[PLAN:{job_id}] Input file: {input_file}")

        from services.filter_service import apply_filters
        from services.session_calculator_service import calculate_sessions
        from services.planner_service import plan_courses

        print(f"[PLAN:{job_id}] Step 1: Applying filters")
        filter_result = apply_filters(
            input_file=str(input_file),
            input=user_input,
        )
        print(f"[PLAN:{job_id}] Step 1 done: {filter_result}")

        filtered_file_value = filter_result.get("output_file")
        if not filtered_file_value:
            _jobs[job_id] = {
                "status": "error",
                "message": "Filter step did not return output_file",
                "filter": filter_result,
            }
            return

        filtered_full_path = _safe_join_output(filtered_file_value)

        if not _validate_xlsx(filtered_full_path):
            _jobs[job_id] = {
                "status": "error",
                "message": "Filtered output file is not a valid .xlsx file",
                "filtered_file": str(filtered_full_path),
                "filter": filter_result,
            }
            return

        print(f"[PLAN:{job_id}] Step 2: Calculating sessions for {filtered_full_path}")
        session_result = calculate_sessions(
            filtered_file=str(filtered_full_path),
        )
        print(f"[PLAN:{job_id}] Step 2 done")

        print(f"[PLAN:{job_id}] Step 3: Planning courses")
        plan_result = plan_courses(
            session_data=session_result,
        )
        print(f"[PLAN:{job_id}] Step 3 done: {plan_result}")

        output_file_value = plan_result.get("output_file")
        if not output_file_value:
            _jobs[job_id] = {
                "status": "error",
                "message": "Plan step did not return output_file",
                "filter": filter_result,
                "sessions": session_result,
                "plan": plan_result,
            }
            return

        plan_file_path = _safe_join_output(output_file_value)

        if not plan_file_path.exists():
            _jobs[job_id] = {
                "status": "error",
                "message": "Plan output file not found",
                "expected_path": str(plan_file_path),
                "filter": filter_result,
                "sessions": session_result,
                "plan": plan_result,
            }
            return

        if not _validate_xlsx(plan_file_path):
            _jobs[job_id] = {
                "status": "error",
                "message": "Plan output file is not a valid .xlsx file",
                "expected_path": str(plan_file_path),
                "filter": filter_result,
                "sessions": session_result,
                "plan": plan_result,
            }
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sharepoint_file_name = f"processed_{timestamp}_{plan_file_path.name}"

        print(f"[PLAN:{job_id}] Completed successfully. Output: {plan_file_path}")

        _jobs[job_id] = {
            "status": "success",
            "file_name": sharepoint_file_name,
            "file_path": str(plan_file_path),
            "filter": filter_result,
            "sessions": session_result,
            "plan": plan_result,
        }

    except Exception as e:
        print(f"[PLAN:{job_id}] ERROR: {e}")
        print(traceback.format_exc())

        _jobs[job_id] = {
            "status": "error",
            "error": str(e),
            "traceback": traceback.format_exc(),
        }

    finally:
        _release_plan_lock()
        print(f"[PLAN:{job_id}] Lock released")


def start_plan_job(
    file_name: Optional[str],
    file_content_base64: Optional[str],
    user_input: Optional[str],
) -> dict:
    if not _plan_lock.acquire(blocking=False):
        print("[PLAN] Request rejected: another plan is already running")
        return {
            "status": "error",
            "message": "A plan request is already in progress. Please wait.",
        }

    try:
        print(f"[PLAN] Request received: input={user_input}, file_name={file_name}")

        UPLOAD_INPUT_DIR.mkdir(parents=True, exist_ok=True)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        if file_content_base64:
            safe_file_name = Path(file_name or "srl_and_wl.xlsx").name
            input_file = UPLOAD_INPUT_DIR / safe_file_name

            file_bytes = _decode_excel_base64(file_content_base64)
            input_file.write_bytes(file_bytes)

            print(f"[PLAN] Saved uploaded Excel: {input_file}")
            print(f"[PLAN] Uploaded file size: {input_file.stat().st_size} bytes")

            if not _validate_xlsx(input_file):
                _release_plan_lock()
                return {
                    "status": "error",
                    "message": "Received file is not a valid .xlsx file",
                    "file_name": safe_file_name,
                    "file_size": input_file.stat().st_size,
                }

        else:
            # Fallback: check writable upload dir first, then static bundled data
            input_file = UPLOAD_INPUT_DIR / "srl_and_wl.xlsx"
            if not input_file.exists():
                input_file = INPUT_DIR / "srl_and_wl.xlsx"
            print(f"[PLAN] No uploaded file received. Using local fallback: {input_file}")

            if not input_file.exists():
                _release_plan_lock()
                return {
                    "status": "error",
                    "message": "No file uploaded and local fallback file not found",
                    "expected_file": str(input_file),
                }

            if not _validate_xlsx(input_file):
                _release_plan_lock()
                return {
                    "status": "error",
                    "message": "Local fallback file is not a valid .xlsx file",
                    "expected_file": str(input_file),
                }

        job_id = uuid.uuid4().hex[:12]

        _jobs[job_id] = {
            "status": "processing",
            "file_name": Path(input_file).name,
            "started_at": datetime.now().isoformat(timespec="seconds"),
        }

        thread = threading.Thread(
            target=_run_plan,
            args=(job_id, input_file, user_input),
            daemon=True,
        )
        thread.start()

        print(f"[PLAN] Job accepted: {job_id}")

        return {
            "status": "accepted",
            "job_id": job_id,
            "poll_url": f"/plan/{job_id}",
            "file_url": f"/plan/{job_id}/file",
        }

    except Exception as e:
        _release_plan_lock()

        print(f"[PLAN] ERROR: {e}")
        print(traceback.format_exc())

        return {
            "status": "error",
            "error": str(e),
            "traceback": traceback.format_exc(),
        }


def get_plan_status(job_id: str) -> dict:
    job = _jobs.get(job_id)

    if not job:
        return {
            "status": "error",
            "message": "Job not found",
            "job_id": job_id,
        }

    response = {
        "status": job.get("status"),
        "job_id": job_id,
    }

    if job.get("status") == "success":
        response["file_name"] = job.get("file_name")
        response["file_url"] = f"/plan/{job_id}/file"

    elif job.get("status") == "error":
        response["message"] = job.get("message")
        response["error"] = job.get("error")
        response["traceback"] = job.get("traceback")

    else:
        response["started_at"] = job.get("started_at")

    return response



def get_plan_file_response(job_id: str) -> dict:
    job = _jobs.get(job_id)

    if not job:
        return {
            "status": "error",
            "message": "Job not found",
            "job_id": job_id,
        }

    if job.get("status") != "success":
        return {
            "status": job.get("status"),
            "message": "File is not ready",
            "job_id": job_id,
        }

    file_path_value = job.get("file_path")
    file_name = job.get("file_name") or "processed_course_plan.xlsx"

    if not file_path_value:
        return {
            "status": "error",
            "message": "No file path found for job",
            "job_id": job_id,
        }

    file_path = Path(file_path_value)

    if not file_path.exists():
        return {
            "status": "error",
            "message": "Output file not found",
            "file_path": str(file_path),
            "job_id": job_id,
        }

    return {
        "status": "success",
        "file_path": str(file_path),
        "file_name": file_name,
    }


def run_local(input_file: Optional[str] = None, user_input: Optional[str] = None) -> dict:
    """
    Run the full planning pipeline locally without HTTP/Power Automate.
    """
    from services.filter_service import apply_filters
    from services.session_calculator_service import calculate_sessions
    from services.planner_service import plan_courses

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    resolved = Path(input_file) if input_file else INPUT_DIR / "srl_and_wl.xlsx"

    if not resolved.exists():
        raise FileNotFoundError(f"Input file not found: {resolved}")

    if not _validate_xlsx(resolved):
        raise ValueError(f"Input file is not a valid .xlsx file: {resolved}")

    print(f"[LOCAL] Using input file: {resolved}")

    print("[LOCAL] Step 1: Applying filters")
    filter_result = apply_filters(
        input_file=str(resolved),
        input=user_input,
    )
    print("[LOCAL] Step 1 done")

    filtered_path = _safe_join_output(filter_result["output_file"])

    print(f"[LOCAL] Step 2: Calculating sessions for {filtered_path}")
    session_result = calculate_sessions(
        filtered_file=str(filtered_path),
    )
    print("[LOCAL] Step 2 done")

    print("[LOCAL] Step 3: Planning courses")
    plan_result = plan_courses(
        session_data=session_result,
    )
    print("[LOCAL] Step 3 done")

    output_path = _safe_join_output(plan_result["output_file"])

    print(f"[LOCAL] Output file: {output_path}")

    return {
        "filter": filter_result,
        "sessions": session_result,
        "plan": plan_result,
        "output_file": str(output_path),
    }