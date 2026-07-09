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
    """
    return abs(experimental_mz - theoretical_mz) / theoretical_mz * 1e6

def identify_molecules(well_position, experimental_peaks, chemical_library, metadata_df, ppm_tolerance=10.0):
    """
    🧠 Core Engine: Compares peaks against library AND bridges them with well metadata.
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
        # 1. Cargar control de placas e historial químico
        metadata = load_all_plates(PLATES_PATH)
        
        if metadata is not None:
            
            chem_lib = load_chemical_library(LIBRARY_PATH)
            sample_list = load_raw_spectra(RAW_PATH, metadata)
            
            # 2. SIMULACIÓN: Ejecutamos el motor cruzando el mapa de tu placa
            print(f"\n🧠 Running molecular identification engine with sample mapping...")
            
            # Simulamos que Morato nos dio datos del pocillo A1
            mock_well = "A1" 
            mock_experimental_peaks = [
                (396.6512, 2500000),  # Ergosterol
                (152.1505, 1800000)   # Arabitol
            ]
            
            # El motor corre y genera el reporte cruzado
            matches = identify_molecules(mock_well, mock_experimental_peaks, chem_lib, metadata, ppm_tolerance=10.0)
            
            # 3. Guardar el resultado final si hubo hallazgos
            if matches:
                results_df = pd.DataFrame(matches)
                output_file = os.path.join(PROCESSED_PATH, "identified_metabolites.csv")
                results_df.to_csv(output_file, index=False)
                print(f"\n💾 Results successfully bridged and saved to: {output_file}")
            
    except FileNotFoundError as e:
        print(f"⚠️ Directory error: {e}. Please check your files in data/")