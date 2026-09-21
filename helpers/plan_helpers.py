import base64
import os
import tempfile
import threading
import traceback
import uuid
import zipfile
from pathlib import Path
from typing import Optional


_TMP_BASE = Path(tempfile.gettempdir()) / "smart_planner"
UPLOAD_INPUT_DIR = _TMP_BASE / "input"
OUTPUT_DIR = Path(
    os.environ.get("SMART_PLANNER_OUTPUT_DIR") or (_TMP_BASE / "output")
)
_plan_lock = threading.Lock()


def _validate_xlsx(path: Path) -> bool:
    return path.exists() and zipfile.is_zipfile(path)


def _decode_excel_base64(file_content_base64: str) -> bytes:
    file_bytes = base64.b64decode(file_content_base64)
    if file_bytes[:4] == b"PK\x03\x04":
        return file_bytes

    try:
        decoded_again = base64.b64decode(file_bytes)
        if decoded_again[:4] == b"PK\x03\x04":
            return decoded_again
    except Exception:
        pass

    return file_bytes


def _resolve_output_file(output_file: str) -> Path:
    path = Path(output_file)
    return path if path.is_absolute() else OUTPUT_DIR / path.name


def start_plan_job(
    file_name: Optional[str],
    file_content_base64: Optional[str],
    user_input: Optional[str],
) -> dict:
    """Run the complete plan pipeline and return the generated workbook."""
    if not _plan_lock.acquire(blocking=False):
        return {
            "status": "error",
            "message": "A plan request is already in progress. Please wait.",
        }

    try:
        if not file_content_base64:
            return {
                "status": "error",
                "message": "file_content_base64 is required",
            }

        UPLOAD_INPUT_DIR.mkdir(parents=True, exist_ok=True)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        input_file = UPLOAD_INPUT_DIR / Path(
            file_name or "srl_and_wl.xlsx"
        ).name
        input_file.write_bytes(_decode_excel_base64(file_content_base64))

        if not _validate_xlsx(input_file):
            return {
                "status": "error",
                "message": "Received file is not a valid .xlsx file",
                "file_name": input_file.name,
            }

        from services.filter_service import apply_filters
        from services.session_calculator_service import calculate_sessions
        from services.planner_service import plan_courses

        filter_result = apply_filters(
            input_file=str(input_file),
            input=user_input,
        )
        filtered_file = filter_result.get("output_file")
        if not filtered_file:
            raise RuntimeError("Filter step did not return an output file")

        session_result = calculate_sessions(
            filtered_file=str(_resolve_output_file(filtered_file)),
        )
        plan_result = plan_courses(session_data=session_result)
        output_file = plan_result.get("output_file")
        if not output_file:
            raise RuntimeError("Plan step did not return an output file")

        plan_file = _resolve_output_file(output_file)
        if not _validate_xlsx(plan_file):
            raise RuntimeError("Plan output file is not a valid .xlsx file")

        return {
            "status": "success",
            "job_id": uuid.uuid4().hex[:12],
            "file_name": plan_file.name,
            "file_content_base64": base64.b64encode(
                plan_file.read_bytes()
            ).decode("ascii"),
            "message": "Plan completed successfully",
        }
    except Exception as error:
        return {
            "status": "error",
            "message": str(error),
            "traceback": traceback.format_exc(),
        }
    finally:
        _plan_lock.release()
