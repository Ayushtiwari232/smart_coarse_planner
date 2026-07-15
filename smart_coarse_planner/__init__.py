import azure.functions as func
import json
import logging


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Health endpoint called.")

    response = {
        "status": "ok",
        "message": "Azure Function is running"
    }

    return func.HttpResponse(
        body=json.dumps(response),
        mimetype="application/json",
        status_code=200
    )