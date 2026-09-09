"""
reports.py
----------
Exports pipeline results to CSV files.
One output file per plate with two sections:
  - targeted_well_level  : one row per well per compound
  - targeted_summary     : one row per compound per condition
  - untargeted_ions      : ranked unknown ions from spectra
"""

import pandas as pd
from pathlib import Path


def export_results(
    targeted_well_df: pd.DataFrame,
    targeted_summary_df: pd.DataFrame,
    untargeted_df: pd.DataFrame,
    plate_id: str,
    output_dir: Path,
) -> Path:
    """
    Writes all results to a single Excel file with three sheets.
    Named: PLATE_XX_DESI_results.xlsx

    Returns the path to the output file.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{plate_id}_DESI_results.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # Sheet 1 — well-level targeted results
        if not targeted_well_df.empty:
            targeted_well_df.to_excel(
                writer,
                sheet_name="Targeted_WellLevel",
                index=False,
            )
            print(f"  Written: Targeted_WellLevel ({len(targeted_well_df)} rows)")

        # Sheet 2 — compound-condition summary
        if not targeted_summary_df.empty:
            targeted_summary_df.to_excel(
                writer,
                sheet_name="Targeted_Summary",
                index=False,
            )
            print(f"  Written: Targeted_Summary ({len(targeted_summary_df)} rows)")

        # Sheet 3 — untargeted ions
        if not untargeted_df.empty:
            untargeted_df.to_excel(
                writer,
                sheet_name="Untargeted_Ions",
                index=False,
            )
            print(f"  Written: Untargeted_Ions ({len(untargeted_df)} rows)")

    print(f"\n  Output saved to: {output_path}")
    return output_path