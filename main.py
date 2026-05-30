from typing import Optional

from fastapi import FastAPI, Query
from services.filter_service import apply_filters
from services.session_calculator_service import calculate_sessions
from services.planner_service import plan_courses

app = FastAPI(title="Smart Coarse Planner")

@app.get("/health")
def health():
    return {
        "status": "ok",
        "message": "FastAPI is running inside Azure Functions"
    }

@app.get("/plan")
def plan(
    input: Optional[str] = Query(None, description="SRL and WL identifier"),
):
    # Step 1: Filter the SRL data
    filter_result = apply_filters(input=input)

    # Step 2: Calculate session requirements from the filtered output
    session_result = calculate_sessions(filtered_file=filter_result["output_file"])

    # Step 3: Plan courses with LLM (trainer assignment + dates)
    plan_result = plan_courses(filtered_file=filter_result["output_file"])

    return {
        "filter": filter_result,
        "sessions": session_result,
        "plan": plan_result,
    }


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
