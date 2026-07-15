# import azure.functions as func
# import json
# import logging


# def main(req: func.HttpRequest) -> func.HttpResponse:
#     logging.info("Health endpoint called.")

#     response = {
#         "status": "ok",
#         "message": "Azure Function is running"
#     }

#     return func.HttpResponse(
#         body=json.dumps(response),
#         mimetype="application/json",
#         status_code=200
#     )

import azure.functions as func
import logging
import json
import base64
from pathlib import Path

from .helpers import start_plan_job


DEFAULT_INPUT_FILE = (
    Path(__file__).resolve().parent
    / "data"
    / "input"
    / "srl_and_wl.xlsx"
)


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Plan endpoint called.")

    try:
        req_body = req.get_json()

        file_name = req_body.get("file_name")
        file_content_base64 = req_body.get("file_content_base64")
        user_input = req_body.get("input")

        if not file_content_base64:

            if not DEFAULT_INPUT_FILE.exists():
                return func.HttpResponse(
                    body=json.dumps({
                        "status": "error",
                        "message": (
                            "file_content_base64 not provided and default "
                            "fallback file not found"
                        ),
                        "expected_file": str(DEFAULT_INPUT_FILE),
                    }),
                    mimetype="application/json",
                    status_code=400,
                )

            logging.info(
                f"Loading fallback file: {DEFAULT_INPUT_FILE}"
            )

            file_bytes = DEFAULT_INPUT_FILE.read_bytes()

            file_content_base64 = base64.b64encode(
                file_bytes
            ).decode("ascii")

            if not file_name:
                file_name = DEFAULT_INPUT_FILE.name

        result = start_plan_job(
            file_name=file_name,
            file_content_base64=file_content_base64,
            user_input=user_input,
        )

        return func.HttpResponse(
            body=json.dumps(result),
            mimetype="application/json",
            status_code=200,
        )

    except Exception as ex:
        logging.exception("Error processing plan request")

        return func.HttpResponse(
            body=json.dumps({
                "status": "error",
                "message": str(ex),
            }),
            mimetype="application/json",
            status_code=500,
        )