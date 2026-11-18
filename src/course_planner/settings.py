import os
from pathlib import Path

from course_planner.models import Settings

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    if value.lower() not in {"true", "false"}:
        raise ValueError(f"{name} must be true or false")
    return value.lower() == "true"


def get_settings() -> Settings:
    return Settings(
        data_root=Path(os.getenv("PLANNER_DATA_ROOT", str(PROJECT_ROOT / "data"))),
        runs_root=Path(os.getenv("PLANNER_RUNS_ROOT", str(PROJECT_ROOT / "runs"))),
        sharepoint_enabled=env_bool("PLANNER_SHAREPOINT_ENABLED"),
        allow_mock_data=env_bool("PLANNER_ALLOW_MOCK_DATA", True),
    )
