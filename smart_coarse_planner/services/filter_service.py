import os
import tempfile
import traceback
import pandas as pd

INPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "input")

# Use persistent HOME directory on Azure, fall back to /tmp for local dev
_HOME_BASE = os.path.join(os.environ.get("HOME", ""), "smart_planner", "output") if os.environ.get("HOME") else None
_DEFAULT_OUTPUT_DIR = _HOME_BASE or "/tmp/smart_course_planner/output"

OUTPUT_DIR = (
    os.environ.get("SMART_PLANNER_OUTPUT_DIR")
    or _DEFAULT_OUTPUT_DIR
)

os.makedirs(OUTPUT_DIR, exist_ok=True)


INPUT_FILE = "srl_and_wl.xlsx"

MODALITY_VALUES = ["MODALITY IXR"]
SITE_CD_VALUES = ["PHC", "CL", "SLC"]


def _read_srl_wl_data(input_path: str) -> pd.DataFrame:
    required_columns = {"Comp Short Ttl", "Req Site Cd", "Local Course Code", "Course Title"}
    workbook = pd.ExcelFile(input_path, engine="openpyxl")

    for sheet_name in workbook.sheet_names:
        preview = pd.read_excel(workbook, sheet_name=sheet_name, header=None)
        for header_row, row in preview.iterrows():
            columns = {str(value).strip() for value in row.dropna()}
            if required_columns.issubset(columns):
                return pd.read_excel(
                    workbook,
                    sheet_name=sheet_name,
                    header=header_row,
                )

    raise ValueError(
        "Could not find SRL/WL columns: " + ", ".join(sorted(required_columns))
    )


def apply_filters(input: str = None, input_file: str = None) -> dict:
    """Read the SRL/WL Excel file, apply filters on Comp Short Ttl (MODALITY)
    and Req Site Cd, then save the filtered result."""
    try:
        input_path = input_file if input_file else os.path.join(INPUT_DIR, INPUT_FILE)
        print(f"[FILTER] Reading file: {input_path}")
        df = _read_srl_wl_data(input_path)
        print(f"[FILTER] Loaded {len(df)} rows, columns: {list(df.columns)}")

        filtered_df = df[
            (df["Comp Short Ttl"].isin(MODALITY_VALUES))
            & (df["Req Site Cd"].isin(SITE_CD_VALUES))
        ]
        print(f"[FILTER] Filtered to {len(filtered_df)} rows")

        output_filename = f"filtered_{input}_v1.xlsx" if input else "filtered_output_v1.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        filtered_df.to_excel(output_path, index=False)
        print(f"[FILTER] Saved to {output_path}")

        return {
            "input": input,
            "total_rows_before_filter": len(df),
            "total_rows_after_filter": len(filtered_df),
            "output_file": output_path,
        }
    except Exception as e:
        print(f"[FILTER] ERROR: {e}")
        print(f"[FILTER] Traceback: {traceback.format_exc()}")
        raise
