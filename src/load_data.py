"""
load_data.py
------------
All file reading functions for the DESI-MS pipeline.
Handles heatmap files, plate metadata, spectra, and chemical library.
Each function returns a clean pandas DataFrame.
"""

import re
import pandas as pd
import openpyxl
from pathlib import Path


# ── Constants ────────────────────────────────────────────────────────────────

ROWS = list("ABCDEFGHIJKLMNOP")   # 16 rows in a 384-well plate
COLS = list(range(1, 25))         # 24 columns

# Known reaction types that may appear in filenames
KNOWN_REACTIONS = ["FMPB", "DNPH", "GirT", "Na", "K", "NH4", "FA"]

# Known ion mode keywords in filenames
NEG_KEYWORDS = ["neg", "negative", "Neg", "NEG"]
POS_KEYWORDS = ["pos", "positive", "Pos", "POS"]


# ── Filename parser ───────────────────────────────────────────────────────────

def parse_heatmap_filename(filepath: Path) -> dict:
    """
    Parses a heatmap filename to extract plate ID, compound name,
    ion mode, and reaction type.

    Filename convention expected:
        Heatmap_<plate_id>_<instrument>_<plate_type>_<compound>_<reaction>.xlsx
    Example:
        Heatmap_DESI1_girardt1_FullP_Arabitol_FMPB.xlsx
        Heatmap_DESI1_girardt1_FullP_Mannitol_Na.xlsx

    Returns a dict with keys:
        filename, plate_id, compound, ion_mode, reaction, confirmed
    """
    name = filepath.stem  # filename without extension

    parsed = {
        "filename":  filepath.name,
        "plate_id":  None,
        "compound":  None,
        "ion_mode":  None,
        "reaction":  None,
        "confirmed": False,
    }

    parts = name.split("_")

    # Try to find compound and reaction from the last two meaningful tokens
    # Convention: ..._<compound>_<reaction>
    if len(parts) >= 2:
        parsed["reaction"] = parts[-1]
        parsed["compound"] = parts[-2]

    # Try to detect ion mode from filename
    name_lower = name.lower()
    if any(k.lower() in name_lower for k in NEG_KEYWORDS):
        parsed["ion_mode"] = "Negative"
    elif any(k.lower() in name_lower for k in POS_KEYWORDS):
        parsed["ion_mode"] = "Positive"
    else:
        parsed["ion_mode"] = "Unknown"

    # Try to extract plate ID — look for a token starting with "PLATE" or "DESI"
    # First try filename
    for part in parts:
        if part.upper().startswith("PLATE"):
            parsed["plate_id"] = part
        break

    # If not found in filename, try the parent folder name
    if not parsed["plate_id"]:
        parent = filepath.parent.name
        if parent.upper().startswith("PLATE"):
            parsed["plate_id"] = parent    
    

    return parsed


def confirm_heatmap_metadata(parsed: dict) -> dict:
    """
    Pauses and asks the user to confirm or correct what was parsed
    from the heatmap filename. Returns the confirmed metadata dict.
    """
    print("\n" + "="*60)
    print(f"  File: {parsed['filename']}")
    print("="*60)
    print(f"  Compound  : {parsed['compound']}")
    print(f"  Ion mode  : {parsed['ion_mode']}")
    print(f"  Reaction  : {parsed['reaction']}")
    print(f"  Plate ID  : {parsed['plate_id']}")
    print("-"*60)

    confirm = input("  Is this correct? (y to confirm, or type correction) [y]: ").strip()

    if confirm.lower() in ("", "y", "yes"):
        parsed["confirmed"] = True
        return parsed

    # Allow field-by-field correction
    print("  Enter new values (press Enter to keep current):")

    new_compound = input(f"  Compound [{parsed['compound']}]: ").strip()
    if new_compound:
        parsed["compound"] = new_compound

    new_mode = input(f"  Ion mode [{parsed['ion_mode']}] (Negative/Positive): ").strip()
    if new_mode:
        parsed["ion_mode"] = new_mode

    new_reaction = input(f"  Reaction [{parsed['reaction']}]: ").strip()
    if new_reaction:
        parsed["reaction"] = new_reaction

    new_plate = input(f"  Plate ID [{parsed['plate_id']}]: ").strip()
    if new_plate:
        parsed["plate_id"] = new_plate

    parsed["confirmed"] = True
    print(f"\n  Confirmed: {parsed['compound']} | {parsed['ion_mode']} | {parsed['reaction']}")
    return parsed


# ── Heatmap reader ────────────────────────────────────────────────────────────

def load_heatmap(filepath: Path) -> pd.DataFrame:
    """
    Reads a 384-well heatmap Excel file (16 rows x 24 cols).
    Row labels are letters (A-P), column labels are numbers (1-24).

    Returns a long-format DataFrame with columns:
        well, intensity
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    records = []

    # Find the start row (where row labels A-P begin)
    start_row = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val == "A":
            start_row = r
            break

    if start_row is None:
        raise ValueError(f"Could not find row labels (A-P) in {filepath.name}")

    # Read 16 rows x 24 columns
    for row_idx, row_letter in enumerate(ROWS):
        r = start_row + row_idx
        for col_idx, col_num in enumerate(COLS):
            c = col_idx + 2   # column 1 in file is the row label
            val = ws.cell(row=r, column=c).value
            intensity = float(val) if val is not None else 0.0
            well = f"{row_letter}{col_num}"
            records.append({"well": well, "intensity": intensity})

    return pd.DataFrame(records)


# ── Plate metadata reader ─────────────────────────────────────────────────────

def load_plate_metadata(plate_file: Path) -> pd.DataFrame:
    """
    Reads the Sample_Metadata sheet from a PLATE_XX.xlsx file.
    Returns a DataFrame with one row per well and all metadata columns.
    """
    wb = openpyxl.load_workbook(plate_file, data_only=True)

    if "Sample_Metadata" not in wb.sheetnames:
        raise ValueError(f"No 'Sample_Metadata' sheet found in {plate_file.name}")

    ws = wb["Sample_Metadata"]

    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    headers = [h for h in headers if h is not None]

    records = []
    for r in range(2, ws.max_row + 1):
        row = {}
        for i, h in enumerate(headers):
            row[h] = ws.cell(row=r, column=i + 1).value
        if any(v is not None for v in row.values()):
            records.append(row)

    df = pd.DataFrame(records)
    print(f"  Loaded plate metadata: {len(df)} wells from {plate_file.name}")
    return df


# ── Chemical library reader ───────────────────────────────────────────────────

def load_chemical_library(csv_path: Path) -> pd.DataFrame:
    """
    Reads the chemical_library.csv file.
    Returns a DataFrame with one row per target compound.
    """
    df = pd.read_csv(csv_path)
    required = ["molecule_name", "target_mz"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"chemical_library.csv is missing required column: {col}")
    print(f"  Loaded chemical library: {len(df)} compounds")
    return df


# ── Spectra reader ────────────────────────────────────────────────────────────

def load_spectra_filtcent(spectra_file: Path) -> dict:
    """
    Reads the filtered, centroided sheets from Spectra.xlsx.
    Returns a dict with keys 'Negative' and 'Positive', each a DataFrame with:
        m/z, Cont, Ex, Dif, YN
    where YN flags which condition the ion is enriched in.
    """
    wb = openpyxl.load_workbook(spectra_file, data_only=True)

    result = {}

    sheet_map = {
        "Negative": "Neg_FiltCent",
        "Positive": "Pos_FiltCent",
    }

    for mode, sheet_name in sheet_map.items():
        if sheet_name not in wb.sheetnames:
            print(f"  Warning: sheet '{sheet_name}' not found in {spectra_file.name}")
            continue

        ws = wb[sheet_name]
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        headers = [h.strip() if isinstance(h, str) else h for h in headers]

        records = []
        for r in range(2, ws.max_row + 1):
            row = {}
            for i, h in enumerate(headers):
                if h is not None:
                    row[h] = ws.cell(row=r, column=i + 1).value
            if any(v is not None for v in row.values()):
                records.append(row)

        df = pd.DataFrame(records)
        df = df.rename(columns={"m/z ": "mz", "m/z": "mz", "Y/N": "YN"})
        df["ion_mode"] = mode
        result[mode] = df
        print(f"  Loaded {mode} FiltCent spectra: {len(df)} m/z values")

    return result