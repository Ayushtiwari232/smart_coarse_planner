"""Compatibility entry point; prefer `planner` for local runs."""

from course_planner.api import app

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8011)
