import os
import pandas as pd

INPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "input")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "output")

INPUT_FILE = "srl_and_wl.xlsx"

MODALITY_VALUES = ["MODALITY IXR"]
SITE_CD_VALUES = ["BEST", "PHC", "CL", "SLC", "VC"]


def apply_filters(input: str = None) -> dict:
    """Read the SRL/WL Excel file, apply filters on Comp Short Ttl (MODALITY)
    and Req Site Cd, then save the filtered result."""

    input_path = os.path.join(INPUT_DIR, INPUT_FILE)
    df = pd.read_excel(input_path)

    filtered_df = df[
        (df["Comp Short Ttl"].isin(MODALITY_VALUES))
        & (df["Req Site Cd"].isin(SITE_CD_VALUES))
    ]

    output_filename = f"filtered_{input}.xlsx" if input else "filtered_output.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filtered_df.to_excel(output_path, index=False)

    return {
        "input": input,
        "total_rows_before_filter": len(df),
        "total_rows_after_filter": len(filtered_df),
        "output_file": output_filename,
    }
