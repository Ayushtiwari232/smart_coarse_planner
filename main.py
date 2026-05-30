from typing import Optional
import traceback

from fastapi import FastAPI, Query

app = FastAPI(title="Smart Course Planner")


@app.get("/health")
def health():
    return {
        "status": "ok",
        "message": "FastAPI is running inside Azure Functions a"
    }


@app.get("/plan")
def plan(
    input: Optional[str] = Query(None, description="SRL and WL identifier"),
):
    try:
        # Import only when /plan is called, not when app starts
        from services.filter_service import apply_filters
        from services.session_calculator_service import calculate_sessions
        from services.planner_service import plan_courses

        filter_result = apply_filters(input=input)

        session_result = calculate_sessions(
            filtered_file=filter_result["output_file"]
        )

        plan_result = plan_courses(
            filtered_file=filter_result["output_file"]
        )

        return {
            "status": "success",
            "filter": filter_result,
            "sessions": session_result,
            "plan": plan_result,
        }

    except Exception as e:
        return {
            "status": "error",
            "error": str(e),
            "traceback": traceback.format_exc()
        }


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)