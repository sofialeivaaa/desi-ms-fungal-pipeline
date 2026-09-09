"""
plate_layout.py
---------------
Joins heatmap intensity data with plate sample metadata.
Produces one flat DataFrame with intensities + all well metadata.
"""

import pandas as pd
from pathlib import Path


def join_heatmap_with_metadata(
    heatmap_df: pd.DataFrame,
    metadata_df: pd.DataFrame,
    heatmap_meta: dict,
) -> pd.DataFrame:
    """
    Joins heatmap intensity values (one row per well) with
    sample metadata (one row per well) on the well column.

    Also attaches compound, ion_mode, reaction, and source_file
    from the parsed heatmap filename metadata.

    Parameters
    ----------
    heatmap_df   : DataFrame with columns [well, intensity]
    metadata_df  : DataFrame with columns [plate_id, well, sample_id,
                   organism, condition, replicate, filename, notes, ...]
    heatmap_meta : dict from parse_heatmap_filename + confirm_heatmap_metadata

    Returns
    -------
    DataFrame with one row per well and all columns merged.
    """
    # Merge on well
    merged = pd.merge(heatmap_df, metadata_df, on="well", how="left")

    # Attach heatmap-level metadata to every row
    merged["compound"]    = heatmap_meta["compound"]
    merged["ion_mode"]    = heatmap_meta["ion_mode"]
    merged["reaction"]    = heatmap_meta["reaction"]
    merged["source_file"] = heatmap_meta["filename"]

    # Add signal flag based on intensity
    merged["signal_flag"] = merged["intensity"].apply(classify_signal)

    # Reorder columns for readability
    priority_cols = [
        "compound", "well", "plate_id", "sample_id",
        "condition", "replicate", "notes",
        "intensity", "signal_flag",
        "source_file", "ion_mode", "reaction",
        "organism",
    ]
    remaining = [c for c in merged.columns if c not in priority_cols]
    merged = merged[priority_cols + remaining]

    return merged


def classify_signal(intensity: float) -> str:
    """
    Simple signal classification based on intensity value.
    Thresholds are relative — adjust after seeing real data distribution.

    strong  : intensity >= 500
    moderate: intensity >= 100
    low     : intensity >= 10
    absent  : intensity < 10
    """
    if intensity >= 500:
        return "strong"
    elif intensity >= 100:
        return "moderate"
    elif intensity >= 10:
        return "low"
    else:
        return "absent"


def find_plate_metadata_file(plate_id: str, plates_dir: Path) -> Path:
    """
    Given a plate ID (e.g. 'PLATE_02'), finds the corresponding
    PLATE_XX.xlsx file in the plates directory.

    If plate_id is None or no match found, prompts the user to select.
    """
    if plate_id:
        candidate = plates_dir / f"{plate_id}.xlsx"
        if candidate.exists():
            return candidate

    # Fallback: list available plate files and ask user
    available = sorted(plates_dir.glob("PLATE_*.xlsx"))
    available = [f for f in available if not f.name.startswith("~")]

    if not available:
        raise FileNotFoundError(f"No PLATE_XX.xlsx files found in {plates_dir}")

    if len(available) == 1:
        print(f"  Using plate metadata file: {available[0].name}")
        return available[0]

    print("\n  Available plate metadata files:")
    for i, f in enumerate(available):
        print(f"    [{i+1}] {f.name}")
    choice = input("  Select plate file number: ").strip()
    try:
        return available[int(choice) - 1]
    except (ValueError, IndexError):
        raise ValueError("Invalid selection")