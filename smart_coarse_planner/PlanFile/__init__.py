# # import azure.functions as func
# # import logging
# # import traceback
# # import json

# # from helpers import get_plan_file_response


# # def main(req: func.HttpRequest) -> func.HttpResponse:
# #     logging.info("Plan file endpoint called.")

# #     try:
# #         job_id = req.route_params.get("job_id")
# #         logging.info(f"Route params: {req.route_params}")
# #         if not job_id:
# #             logging.warning("job_id is required")
# #             return func.HttpResponse(
# #                 body=json.dumps({
# #                     "status": "error",
# #                     "message": "job_id is required"
# #                 }),
# #                 mimetype="application/json",
# #                 status_code=400
# #             )

# #         result = get_plan_file_response(job_id)

# #         if result.get("status") != "success":
# #             return func.HttpResponse(
# #                 body=json.dumps(result),
# #                 mimetype="application/json",
# #                 status_code=404
# #             )

# #         with open(result["file_path"], "rb") as file:
# #             file_bytes = file.read()

# #         return func.HttpResponse(
# #             body=file_bytes,
# #             status_code=200,
# #             headers={
# #                 "Content-Disposition": f'attachment; filename="{result["file_name"]}"',
# #                 "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# #             }
# #         )

# #     except Exception as ex:
# #         logging.exception("Error retrieving plan file")

# #         return func.HttpResponse(
# #             body=json.dumps({
# #                 "status": "error",
# #                 "message": str(ex),
# #                 "traceback": traceback.format_exc()
# #             }),
# #             mimetype="application/json",
# #             status_code=500
# #         )
# import logging
# import traceback


# def get_plan_file_response(job_id: str):
#     try:
#         logging.info(f"get_plan_file_response called for job_id={job_id}")

#         job = _jobs.get(job_id)

#         if not job:
#             logging.warning(f"Job not found: {job_id}")

#             return {
#                 "status": "error",
#                 "message": "Job not found",
#                 "job_id": job_id
#             }

#         logging.info(f"Job found: {job}")
#         logging.info(f"Job status: {job.get('status')}")

#         file_path = job.get("file_path")
#         file_name = job.get("file_name")

#         logging.info(f"file_path={file_path}")
#         logging.info(f"file_name={file_name}")

#         if job.get("status") != "success":
#             return {
#                 "status": job.get("status"),
#                 "message": "File is not ready",
#                 "job_id": job_id
#             }

#         if not file_path:
#             logging.error(f"No file_path found for job {job_id}")

#             return {
#                 "status": "error",
#                 "message": "No file path found",
#                 "job_id": job_id
#             }

#         return {
#             "status": "success",
#             "file_path": file_path,
#             "file_name": file_name,
#             "job_id": job_id
#         }

#     except Exception as ex:
#         logging.exception(
#             f"Error in get_plan_file_response for job_id={job_id}"
#         )

#         return {
#             "status": "error",
#             "message": str(ex),
#             "job_id": job_id,
#             "traceback": traceback.format_exc()
#         }

import azure.functions as func
import logging
import traceback
import json

from helpers import get_plan_file_response


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Plan file endpoint called.")

    try:
        logging.info(f"Route params: {req.route_params}")

        job_id = req.route_params.get("job_id")
        logging.info(f"job_id: {job_id}")

        if not job_id:
            logging.warning("job_id is required")

            return func.HttpResponse(
                body=json.dumps({
                    "status": "error",
                    "message": "job_id is required"
                }),
                mimetype="application/json",
                status_code=400
            )

        result = get_plan_file_response(job_id)

        logging.info(f"Helper response: {result}")

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
                "Content-Disposition": f'attachment; filename="{result['file_name']}"',
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