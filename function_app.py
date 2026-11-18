"""Azure Functions v2 adapter for the shared application."""

import json
import logging
from pathlib import Path

import azure.functions as func
from pydantic import ValidationError

from course_planner.api import create_plan
from course_planner.models import PlanRequest
from course_planner.pipeline import get_run
from course_planner.settings import get_settings

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)


def response(payload: dict, status: int = 200) -> func.HttpResponse:
    return func.HttpResponse(json.dumps(payload), mimetype="application/json", status_code=status)


@app.route(route="plan", methods=["POST"])
def plan(req: func.HttpRequest) -> func.HttpResponse:
    try:
        request = PlanRequest.model_validate(req.get_json())
        return response(create_plan(request))
    except (ValueError, ValidationError) as error:
        return response({"status": "failed", "message": str(error)}, 400)
    except Exception:
        logging.exception("Azure plan request failed")
        return response(
            {"status": "failed", "message": "Planning failed; inspect server logs"}, 500
        )


@app.route(route="plan/{job_id}", methods=["GET"])
def status(req: func.HttpRequest) -> func.HttpResponse:
    try:
        return response(get_run(get_settings(), req.route_params["job_id"]))
    except ValueError as error:
        return response({"message": str(error)}, 400)
    except FileNotFoundError as error:
        return response({"message": str(error)}, 404)


@app.route(route="plan/{job_id}/file", methods=["GET"])
def download(req: func.HttpRequest) -> func.HttpResponse:
    try:
        result = get_run(get_settings(), req.route_params["job_id"])
        if result["status"] not in {"completed", "completed_with_warnings"}:
            return response({"message": "Plan is not complete"}, 409)
        return func.HttpResponse(
            Path(result["output_file"]).read_bytes(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{result["file_name"]}"'},
        )
    except ValueError as error:
        return response({"message": str(error)}, 400)
    except FileNotFoundError as error:
        return response({"message": str(error)}, 404)
