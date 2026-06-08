import os
import traceback
import pandas as pd

INPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "input")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "output")

INPUT_FILE = "srl_and_wl.xlsx"

MODALITY_VALUES = ["MODALITY IXR"]
SITE_CD_VALUES = ["PHC", "CL", "SLC"]


def apply_filters(input: str = None, input_file: str = None) -> dict:
    """Read the SRL/WL Excel file, apply filters on Comp Short Ttl (MODALITY)
    and Req Site Cd, then save the filtered result."""
    try:
        input_path = input_file if input_file else os.path.join(INPUT_DIR, INPUT_FILE)
        print(f"[FILTER] Reading file: {input_path}")
        df = pd.read_excel(input_path, engine="openpyxl")
        print(f"[FILTER] Loaded {len(df)} rows, columns: {list(df.columns)}")

        filtered_df = df[
            (df["Comp Short Ttl"].isin(MODALITY_VALUES))
            & (df["Req Site Cd"].isin(SITE_CD_VALUES))
        ]
        print(f"[FILTER] Filtered to {len(filtered_df)} rows")

        output_filename = f"filtered_{input}.xlsx" if input else "filtered_output.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        filtered_df.to_excel(output_path, index=False)
        print(f"[FILTER] Saved to {output_path}")

        return {
            "input": input,
            "total_rows_before_filter": len(df),
            "total_rows_after_filter": len(filtered_df),
            "output_file": output_filename,
        }
    except Exception as e:
        print(f"[FILTER] ERROR: {e}")
        print(f"[FILTER] Traceback: {traceback.format_exc()}")
        raise
