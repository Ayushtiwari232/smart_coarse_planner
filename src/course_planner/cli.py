import argparse
import json
import logging
import shutil
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from course_planner.inputs import INPUT_FILES, read_demand
from course_planner.models import PlanRequest, quarter_dates
from course_planner.pipeline import run_plan
from course_planner.settings import get_settings


def init_quarter(location: str, quarter: str, source_quarter: str) -> Path:
    quarter_dates(quarter)
    quarter_dates(source_quarter)
    settings = get_settings()
    source = settings.data_root / "quarters" / source_quarter / location
    target = settings.data_root / "quarters" / quarter / location
    if target.exists():
        raise FileExistsError(f"Quarter/location already exists: {target}")
    missing = [name for name in INPUT_FILES if not (source / name).is_file()]
    if missing:
        raise ValueError(f"Template is missing files: {missing}")
    template_demand = read_demand(source.parent / "shared" / "demand.xlsx")
    target.mkdir(parents=True)
    for name in INPUT_FILES:
        shutil.copyfile(source / name, target / name)
    for name in ("availability.xlsx", "bookings.xlsx"):
        pd.read_excel(target / name).iloc[:0].to_excel(target / name, index=False)
    rules = {
        "notes": [],
        "skip_first_week": False,
        "excluded_dates": [],
        "mock_data": True,
        "assumptions": [
            f"Template copied from {source_quarter}; confirm courses, qualifications and priorities.",
            "Availability and bookings are empty; supply and confirm quarterly data before planning.",
        ],
    }
    (target / "rules.json").write_text(json.dumps(rules, indent=2), encoding="utf-8")
    shared = target.parent / "shared"
    shared.mkdir(exist_ok=True)
    demand = shared / "demand.xlsx"
    if not demand.exists():
        template_demand.iloc[:0].to_excel(demand, index=False)
    return target


def main(argv: list[str] | None = None) -> int:
    load_dotenv()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s %(message)s")
    parser = argparse.ArgumentParser(description="Quarterly course planner")
    parser.add_argument("command", choices=["validate", "plan", "init-quarter"])
    parser.add_argument("--location", choices=["best", "cleveland", "singapore"], required=True)
    parser.add_argument("--quarter", required=True, help="YYYY-Q1 through YYYY-Q4")
    parser.add_argument("--mode", choices=["ai-first", "deterministic"], default="ai-first")
    parser.add_argument("--input-file", type=Path)
    parser.add_argument("--from-quarter", default="2027-Q1")
    args = parser.parse_args(argv)
    try:
        if args.command == "init-quarter":
            print(init_quarter(args.location, args.quarter, args.from_quarter))
            return 0
        request = PlanRequest(location=args.location, quarter=args.quarter, mode=args.mode)
        result = run_plan(
            request,
            get_settings(),
            validate_only=args.command == "validate",
            input_file=args.input_file,
        )
        print(json.dumps(result, indent=2, default=str))
        return 0
    except (ValueError, OSError, RuntimeError) as error:
        logging.getLogger(__name__).error("%s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
