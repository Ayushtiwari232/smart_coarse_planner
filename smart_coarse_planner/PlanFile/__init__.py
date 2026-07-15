import azure.functions as func
import logging
import traceback
import json

from helpers import get_plan_file_response


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Plan file endpoint called.")

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

        result = get_plan_file_response(job_id)

        if result.get("status") != "success":
            return func.HttpResponse(
                body=json.dumps(result),
                mimetype="application/json",
                status_code=404
            )

        with open(result["file_path"], "rb") as file:
            file_bytes = file.read()

        return func.HttpResponse(
            body=file_bytes,
            status_code=200,
            headers={
                "Content-Disposition": f'attachment; filename="{result["file_name"]}"',
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
        )

    except Exception as ex:
        logging.exception("Error retrieving plan file")

        return func.HttpResponse(
            body=json.dumps({
                "status": "error",
                "message": str(ex),
                "traceback": traceback.format_exc()
            }),
            mimetype="application/json",
            status_code=500
        )