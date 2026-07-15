import azure.functions as func
import json
import logging

from helpers import get_plan_status


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Plan status endpoint called.")

    try:
        job_id = req.route_params.get("job_id")

        if not job_id:
            return func.HttpResponse(
                body=json.dumps({
                    "status": "error",
                    "message": "job_id is required"
                }),
                mimetype="application/json",
                status_code=400
            )

        result = get_plan_status(job_id)

        return func.HttpResponse(
            body=json.dumps(result),
            mimetype="application/json",
            status_code=200
        )

    except Exception as ex:
        return func.HttpResponse(
            body=json.dumps({
                "status": "error",
                "message": str(ex)
            }),
            mimetype="application/json",
            status_code=500
        )