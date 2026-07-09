import os
import pandas as pd
import pymzml

def load_all_plates(plates_folder):
    """
    Scans the data/plates folder and loads all available PLATE_0X.xlsx files.
    """
    print(f"📋 Scanning for plate metadata in: {plates_folder}")
    
    # Buscamos todos los archivos que empiecen con 'PLATE_' y terminen en '.xlsx'
    plate_files = [f for f in os.listdir(plates_folder) if f.startswith('PLATE_') and f.endswith('.xlsx')]
    
    if not plate_files:
        print("⚠️ No plate metadata files found! Make sure your Excel files are in data/plates/")
        return None
        
    all_plates_data = []
    for file in plate_files:
        file_path = os.path.join(plates_folder, file)
        print(f" -> Loading {file}...")
        df = pd.read_excel(file_path)
        all_plates_data.append(df)
        
    # Combinamos todas las placas en una sola gran tabla de control
    combined_metadata = pd.concat(all_plates_data, ignore_index=True)
    print(f"✨ Successfully loaded a total of {len(combined_metadata)} rows/wells from {len(plate_files)} plate(s).")
    
    return combined_metadata

def load_chemical_library(library_path):
    """
    Loads the reference database with the target molecules we expect from the fungi.
    """
    print(f"🔬 Loading chemical reference library from: {library_path}")
    return pd.read_csv(library_path)

def calculate_ppm_error(experimental_mz, theoretical_mz):
    """
    Calculates the mass accuracy error in parts per million (ppm).
    Formula: (|Exp - Theo| / Theo) * 10^6
    """
    return abs(experimental_mz - theoretical_mz) / theoretical_mz * 1e6

def identify_molecules(experimental_peaks, chemical_library, ppm_tolerance=5.0):
    """
    Compares detected m/z peaks against our local database using ppm tolerance.
    """
    print(f"\n🧠 Running molecular identification engine (Tolerance: {ppm_tolerance} ppm)...")
    identifications = []
    
    for peak_mz, intensity in experimental_peaks:
        for idx, row in chemical_library.iterrows():
            error = calculate_ppm_error(peak_mz, row['target_mz'])
            
            if error <= ppm_tolerance:
                print(f"   ✨ MATCH FOUND! m/z {peak_mz:.4f} identified as {row['molecule_name']} (Error: {error:.2f} ppm)")
                identifications.append({
                    "experimental_mz": peak_mz,
                    "intensity": intensity,
                    "matched_molecule": row['molecule_name'],
                    "formula": row['formula'],
                    "ppm_error": error
                })
                
    return pd.DataFrame(identifications)

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
    # Define relative paths matching your exact structure
    PLATES_PATH = os.path.join("data", "plates")
    RAW_PATH = os.path.join("data", "raw")
    LIBRARY_PATH = os.path.join("data", "chemical_library.csv")
    PROCESSED_PATH = os.path.join("data", "processed")
    
    # Execute workflow
    try:
        # 1. Cargar control de pocillos
        metadata = load_all_plates(PLATES_PATH)
        
        # 2. Cargar base de datos química de referencia
        if metadata is not None:
            chem_lib = load_chemical_library(LIBRARY_PATH)
            
            # 3. Escanear espectros crudos
            sample_list = load_raw_spectra(RAW_PATH, metadata)
            
            # 4. SIMULACIÓN: Datos de prueba mientras llegan los de Morato
            # Simulamos que el espectrómetro leyó un pico en 405.2650 m/z
            mock_experimental_peaks = [
                (396.6512, 2500000),  # Pico muy cercano a Ergosterol
                (152.1505, 1800000),  # Pico muy cercano a Arabitol
                (999.1234, 120000)    # Ruido de fondo (no tendrá match)
             ]
            
            # Ejecutar emparejamiento automático con un margen de 10 ppm
            results_df = identify_molecules(mock_experimental_peaks, chem_lib, ppm_tolerance=10.0)
            
    except FileNotFoundError as e:
        print(f"⚠️ Directory error: {e}. Please check your files in data/")