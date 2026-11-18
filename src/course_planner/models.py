"""Validated contracts shared by local, API and storage adapters."""

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

Location = Literal["best", "cleveland", "singapore"]


def quarter_dates(quarter: str) -> tuple[date, date]:
    if not re.fullmatch(r"\d{4}-Q[1-4]", quarter):
        raise ValueError("Quarter must have the form YYYY-Q1 through YYYY-Q4")
    year, number = int(quarter[:4]), int(quarter[-1])
    month = 3 * number - 2
    return date(year, month, 1), date(year, month + 2, calendar.monthrange(year, month + 2)[1])


class Rules(BaseModel):
    model_config = ConfigDict(extra="forbid")
    notes: list[str] = Field(default_factory=list)
    skip_first_week: bool = False
    excluded_dates: list[date] = Field(default_factory=list)
    mock_data: bool = False
    assumptions: list[str] = Field(default_factory=list)


class TableRow(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="ignore")


class TrainerRow(TableRow):
    trainer_id: str = Field(min_length=1)
    trainer: str = Field(min_length=1)


class QualificationRow(TableRow):
    course_code: str = Field(min_length=1)
    trainer_id: str = Field(min_length=1)


class PriorityRow(QualificationRow):
    rank: int = Field(ge=1, le=5)


class AvailabilityRow(TableRow):
    trainer_id: str = Field(min_length=1)
    date: date
    available_fraction: float = Field(ge=0, le=1, allow_inf_nan=False)


class CompatibilityRow(TableRow):
    course_code: str = Field(min_length=1)
    lab: str = Field(min_length=1)


class BookingRow(CompatibilityRow):
    start_date: date
    end_date: date

    @model_validator(mode="after")
    def valid_dates(self):
        if self.end_date < self.start_date:
            raise ValueError("Booking end must not precede start")
        return self


class PlanRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    location: Location = "best"
    quarter: str = "2027-Q1"
    mode: Literal["ai-first", "deterministic"] = "ai-first"
    file_name: str | None = None
    file_content_base64: str | None = None
    input: str | None = None

    @model_validator(mode="after")
    def valid_quarter(self):
        quarter_dates(self.quarter)
        return self


class Settings(BaseModel):
    data_root: Path
    runs_root: Path
    sharepoint_enabled: bool = False
    allow_mock_data: bool = True
