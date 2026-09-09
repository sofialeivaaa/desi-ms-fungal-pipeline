import os
import pandas as pd
import pymzml

def load_all_plates(plates_folder):
    """
    Scans the data/plates folder and loads all available PLATE_0X.xlsx files.
    """
    print(f"📋 Scanning for plate metadata in: {plates_folder}")
    plate_files = [f for f in os.listdir(plates_folder) if f.startswith('PLATE_') and f.endswith('.xlsx')]
    
    if not plate_files:
        print("⚠️ No plate metadata files found! Make sure your Excel files are in data/plates/")
        return None
        
    all_plates_data = []
    for file in plate_files:
        file_path = os.path.join(plates_folder, file)
        print(f" -> Loading {file}...")
        
        # 🎯 CAMBIO AQUÍ: Especificamos la pestaña correcta (cambia 'sample metadata' si se llama distinto)
        df = pd.read_excel(file_path, sheet_name='Sample_Metadata')
        
        # Limpia espacios en blanco ocultos en los nombres de las columnas
        df.columns = df.columns.astype(str).str.strip().str.lower()
        
        all_plates_data.append(df)
        
    combined_metadata = pd.concat(all_plates_data, ignore_index=True)
    print(f"Successfully loaded a total of {len(combined_metadata)} rows/wells from {len(plate_files)} plate(s).")
    return combined_metadata

def load_chemical_library(library_path):
    """
    Loads the reference database with the target molecules we expect from the fungi.
    """
    print(f"Loading chemical reference library from: {library_path}")
    return pd.read_csv(library_path)

def calculate_ppm_error(experimental_mz, theoretical_mz):
    """
    Calculates the mass accuracy error in parts per million (ppm).
    """
    return abs(experimental_mz - theoretical_mz) / theoretical_mz * 1e6

def identify_molecules(well_position, experimental_peaks, chemical_library, metadata_df, ppm_tolerance=10.0):
    """
    Core Engine: Compares peaks against library AND bridges them with well metadata.
    """
    identifications = []
    
    # Buscar la información del hongo para este pocillo específico
    well_info = metadata_df[metadata_df['well'] == well_position]
    
    if well_info.empty:
        organism = "Unknown"
        condition = "Unknown"
        sample_id = "Unknown"
    else:
        organism = well_info.iloc[0]['organism']
        condition = well_info.iloc[0]['condition']
        sample_id = well_info.iloc[0]['sample_id']

    for peak_mz, intensity in experimental_peaks:
        for idx, row in chemical_library.iterrows():
            error = calculate_ppm_error(peak_mz, row['target_mz'])
            
            if error <= ppm_tolerance:
                print(f"   ✨ MATCH in Well {well_position} ({organism}): {row['molecule_name']} Found! (Error: {error:.2f} ppm)")
                identifications.append({
                    "well": well_position,
                    "sample_id": sample_id,
                    "organism": organism,
                    "condition": condition,
                    "experimental_mz": peak_mz,
                    "intensity": intensity,
                    "matched_molecule": row['molecule_name'],
                    "type": row['type'],
                    "ppm_error": error
                })
                
    return identifications

def load_raw_spectra(raw_folder, metadata_df):
    """
    Scans for .mzML files and stands ready to match them with the loaded plates.
    """
    print(f"\nScanning for spectrometry files in: {raw_folder}")
    files = [f for f in os.listdir(raw_folder) if f.endswith('.mzML')]
    
    if not files:
        print("⚠️ No .mzML files found in data/raw yet. Standing by for Morato's files!")
    else:
        print(f"Found {len(files)} file(s) ready to process.")
    return files

if __name__ == "__main__":
    PLATES_PATH = os.path.join("data", "plates")
    RAW_PATH = os.path.join("data", "raw")
    LIBRARY_PATH = os.path.join("data", "chemical_library.csv")
    PROCESSED_PATH = os.path.join("data", "processed")

    try:
        # 1. Load plate metadata and chemical library
        metadata = load_all_plates(PLATES_PATH)

        if metadata is not None:

            chem_lib = load_chemical_library(LIBRARY_PATH)
            sample_list = load_raw_spectra(RAW_PATH, metadata)

            

            # 3. Heatmap intensity analysis (pre-processed results from Dr. Morato)
            print("\n" + "="*50)
            print("HEATMAP INTENSITY ANALYSIS")
            print("="*50)

            import sys
            from pathlib import Path
            sys.path.insert(0, os.path.dirname(__file__))

            from load_data import (parse_heatmap_filename, confirm_heatmap_metadata,
                                   load_heatmap, load_plate_metadata, load_spectra_filtcent)
            from plate_layout import join_heatmap_with_metadata, find_plate_metadata_file
            from analysis import run_targeted_analysis, summarize_targeted, run_untargeted_analysis
            from reports import export_results

            PLATES_PATH_P    = Path("data/plates")
            SPECTRA_PATH_P   = Path("data/spectra")
            HEATMAPS_PATH_P  = Path("data/raw/Heatmaps")
            PROCESSED_PATH_P = Path("data/processed")

            chem_lib_p = load_chemical_library(LIBRARY_PATH)

            heatmap_files = sorted([
                f for f in HEATMAPS_PATH_P.rglob("Heatmap_*.xlsx")
                if not f.name.startswith("~")
            ])

            if not heatmap_files:
                print("⚠️  No heatmap files found in data/plates/")
            else:
                all_targeted_well    = []
                all_targeted_summary = []
                plate_id_used        = None

                for heatmap_path in heatmap_files:
                    parsed    = parse_heatmap_filename(heatmap_path)
                    confirmed = confirm_heatmap_metadata(parsed)

                    heatmap_df  = load_heatmap(heatmap_path)
                    plate_file  = find_plate_metadata_file(confirmed["plate_id"], PLATES_PATH_P)
                    metadata_df = load_plate_metadata(plate_file)

                    if not confirmed["plate_id"] and "plate_id" in metadata_df.columns:
                        confirmed["plate_id"] = metadata_df["plate_id"].iloc[0]
                    plate_id_used = confirmed["plate_id"]

                    joined_df        = join_heatmap_with_metadata(heatmap_df, metadata_df, confirmed)
                    targeted_well    = run_targeted_analysis(joined_df, chem_lib_p)
                    targeted_summary = summarize_targeted(targeted_well)

                    all_targeted_well.append(targeted_well)
                    all_targeted_summary.append(targeted_summary)

                spectra_files = sorted([
                    f for f in SPECTRA_PATH_P.glob("Spectra*.xlsx")
                    if not f.name.startswith("~")
                ])
                untargeted_df = pd.DataFrame()
                if spectra_files:
                    spectra_dict  = load_spectra_filtcent(spectra_files[0])
                    untargeted_df = run_untargeted_analysis(spectra_dict, top_n=30)

                plate_id_used = plate_id_used or "PLATE_UNKNOWN"
                export_results(
                    targeted_well_df    = pd.concat(all_targeted_well,    ignore_index=True),
                    targeted_summary_df = pd.concat(all_targeted_summary, ignore_index=True),
                    untargeted_df       = untargeted_df,
                    plate_id            = plate_id_used,
                    output_dir          = PROCESSED_PATH_P,
                )

    except FileNotFoundError as e:
        print(f"Directory error: {e}. Please check your files in data/")