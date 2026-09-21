import azure.functions as func
import base64
import json
import traceback
from pathlib import Path

from helpers import start_plan_job


app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

DEFAULT_INPUT_FILE = (
    Path(__file__).resolve().parent
    / "data"
    / "input"
    / "New SRL and WL for Students_51_9219480390659314172.xlsx"
)


@app.function_name(name="plan")
@app.route(route="plan", methods=["POST"])
def plan(req: func.HttpRequest) -> func.HttpResponse:
    try:
        request_body = req.get_json()
    except ValueError:
        request_body = {}

    try:
        file_name = request_body.get("file_name")
        file_content_base64 = request_body.get("file_content_base64")
        user_input = request_body.get("input")

        if not file_content_base64:
            if not DEFAULT_INPUT_FILE.exists():
                return func.HttpResponse(
                    body=json.dumps({
                        "status": "error",
                        "message": "Default SRL/WL input file was not found",
                        "expected_file": str(DEFAULT_INPUT_FILE),
                    }),
                    mimetype="application/json",
                    status_code=400,
                )

            file_content_base64 = base64.b64encode(
                DEFAULT_INPUT_FILE.read_bytes()
            ).decode("ascii")
            file_name = file_name or DEFAULT_INPUT_FILE.name

        result = start_plan_job(
            file_name=file_name,
            file_content_base64=file_content_base64,
            user_input=user_input,
        )
        status_code = 200 if result.get("status") == "success" else 500
        return func.HttpResponse(
            body=json.dumps(result),
            mimetype="application/json",
            status_code=status_code,
        )
    except Exception as error:
        return func.HttpResponse(
            body=json.dumps({
                "status": "error",
                "message": str(error),
                "traceback": traceback.format_exc(),
            }),
            mimetype="application/json",
            status_code=500,
        )