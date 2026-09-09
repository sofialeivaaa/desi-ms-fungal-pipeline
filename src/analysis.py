"""
analysis.py
-----------
Targeted and untargeted analysis functions.

Targeted : checks whether known compounds from chemical_library.csv
           are present in the heatmap data and how intense they are.

Untargeted: takes the FiltCent spectra and ranks intense unknown ions,
            flagging which condition they are enriched in.
"""

import pandas as pd


# ── Targeted analysis ─────────────────────────────────────────────────────────

def run_targeted_analysis(
    joined_df: pd.DataFrame,
    library_df: pd.DataFrame,
    mz_tolerance_ppm: float = 10.0,
) -> pd.DataFrame:
    """
    For each compound in the chemical library, checks whether the
    heatmap compound matches (by name) and annotates with library info.

    Since heatmap files are already compound-specific (one file per compound),
    the match is done by compound name rather than m/z search.

    Returns the joined_df with library columns added where a match is found.
    Adds a 'meeting_recommendation' column summarizing what to tell Dr. Morato.
    """
    # Normalize compound names for matching
    lib = library_df.copy()
    lib["compound_norm"] = lib["molecule_name"].str.strip().str.lower()

    df = joined_df.copy()
    df["compound_norm"] = df["compound"].str.strip().str.lower()

    # Merge library metadata onto the data
    df = df.merge(
        lib[["compound_norm", "exact_mass", "target_mz", "type", "adduct_type"]],
        on="compound_norm",
        how="left",
    )
    df = df.drop(columns=["compound_norm"])

    # Summarize per compound for meeting recommendation
    df["meeting_recommendation"] = df.apply(_recommend, axis=1)

    return df


def _recommend(row) -> str:
    """
    Generates a simple recommendation for each well-compound observation.
    """
    flag = row.get("signal_flag", "absent")
    if flag == "strong":
        return "TARGETED — confirm and quantify"
    elif flag == "moderate":
        return "TARGETED — worth quantifying"
    elif flag == "low":
        return "BORDERLINE — discuss with Dr. Morato"
    else:
        return "LOW/ABSENT — consider untargeted if expected"


def summarize_targeted(targeted_df: pd.DataFrame) -> pd.DataFrame:
    """
    Creates a compound-level summary across all wells:
    mean intensity, max intensity, n wells detected, and
    overall recommendation per condition per compound.
    """
    summary = (
        targeted_df
        .groupby(["compound", "condition", "ion_mode", "reaction"])
        .agg(
            mean_intensity=("intensity", "mean"),
            max_intensity=("intensity", "max"),
            n_wells_total=("well", "count"),
            n_wells_detected=("signal_flag", lambda x: (x != "absent").sum()),
        )
        .round(2)
        .reset_index()
    )

    # Overall recommendation per compound-condition
    summary["overall_recommendation"] = summary.apply(
        lambda r: "TARGETED — strong signal" if r["mean_intensity"] >= 500
        else "TARGETED — moderate signal" if r["mean_intensity"] >= 100
        else "LOW — discuss with Dr. Morato" if r["mean_intensity"] >= 10
        else "ABSENT — consider untargeted",
        axis=1,
    )

    return summary


# ── Untargeted analysis ───────────────────────────────────────────────────────

def run_untargeted_analysis(
    spectra_dict: dict,
    top_n: int = 30,
) -> pd.DataFrame:
    """
    Takes the FiltCent spectra (Negative and Positive) and returns
    a ranked list of intense unknown ions, sorted by absolute intensity
    difference between Cont and Ex conditions.

    Parameters
    ----------
    spectra_dict : dict with keys 'Negative' and 'Positive',
                   each a DataFrame from load_spectra_filtcent()
    top_n        : number of top ions to return per mode

    Returns
    -------
    DataFrame with columns:
        mz, ion_mode, intensity_Cont, intensity_Ex,
        difference, enriched_in, recommendation
    """
    frames = []

    for mode, df in spectra_dict.items():
        if df is None or df.empty:
            continue

        df = df.copy()
        df = df.rename(columns={"Cont": "intensity_Cont", "Ex": "intensity_Ex"})

        # Calculate absolute difference and max intensity for ranking
        df["abs_difference"] = df["Dif"].abs()
        df["max_intensity"]  = df[["intensity_Cont", "intensity_Ex"]].max(axis=1)

        # Rename YN to enriched_in for clarity
        df = df.rename(columns={"YN": "enriched_in"})

        # Add recommendation
        df["recommendation"] = df.apply(_untargeted_recommend, axis=1)

        # Keep and reorder columns
        keep = ["mz", "ion_mode", "intensity_Cont", "intensity_Ex",
                "abs_difference", "enriched_in", "recommendation"]
        df = df[[c for c in keep if c in df.columns]]

        # Sort by absolute difference descending and take top N
        df = df.sort_values("abs_difference", ascending=False).head(top_n)

        frames.append(df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True).sort_values(
        "abs_difference", ascending=False
    )


def _untargeted_recommend(row) -> str:
    """
    Recommendation for each unknown ion in the untargeted analysis.
    """
    yn = row.get("enriched_in")
    diff = abs(row.get("Dif", row.get("abs_difference", 0)))

    if pd.isna(yn) or yn is None:
        return "PRESENT IN BOTH — not condition-specific"
    elif diff > 0.5:
        return f"UNTARGETED — strong enrichment in {yn}, high priority"
    elif diff > 0.1:
        return f"UNTARGETED — moderate enrichment in {yn}"
    else:
        return f"UNTARGETED — weak enrichment in {yn}"