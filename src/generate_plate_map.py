"""
Genera un libro de Excel (.xlsx) para el seguimiento de una microplaca de 384 pocillos
(Eppendorf Twin.tec 384) usada en el pipeline DESI-MS de hongos.

Pestañas:
  1. Plate_Map        -> cuadrícula visual de la placa (16 filas A-P x 24 columnas)
  2. Sample_Metadata   -> tabla plana (384 filas) lista para pandas

Requiere: openpyxl

Uso:
    python src/generate_plate_map.py PLATE_01
    python src/generate_plate_map.py PLATE_02
"""
import sys
import string
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

FONT_NAME = "Calibri"
PLATE_ID = sys.argv[1] if len(sys.argv) > 1 else "PLATE_01"

# Carpeta de salida: data/plates/ relativa a la raíz del repo
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "data" / "plates"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ROWS = list(string.ascii_uppercase[:16])   # A ... P
COLS = list(range(1, 25))                  # 1 ... 24

# ---------------------------------------------------------------------------
# Ejemplos ficticios para las primeras posiciones (sirven de plantilla visual)
# well -> (sample_id, organism, condition, replicate, notes)
# ---------------------------------------------------------------------------
EXAMPLES = {
    "A1": ("SAM_001", "Aspergillus niger",        "Control",               1, "Enviado por Morato"),
    "A2": ("SAM_002", "Aspergillus niger",        "Control",               2, "Enviado por Morato"),
    "A3": ("SAM_003", "Aspergillus niger",        "Tratado con compuesto X", 1, "Mutante bio-ingeniería"),
    "A4": ("SAM_004", "Aspergillus niger",        "Tratado con compuesto X", 2, "Mutante bio-ingeniería"),
    "B1": ("SAM_005", "Penicillium chrysogenum",  "Control",               1, "Cultivo de 5 días"),
    "B2": ("SAM_006", "Penicillium chrysogenum",  "Control",               2, "Cultivo de 5 días"),
    "B3": ("SAM_007", "Penicillium chrysogenum",  "Tratado con compuesto X", 1, "Cultivo de 5 días"),
    "B4": ("SAM_008", "Penicillium chrysogenum",  "Tratado con compuesto X", 2, "Cultivo de 5 días"),
    "C1": ("SAM_009", "Blank",                    "Control",               1, "Blanco / medio sin inocular"),
}

PLATE_LABEL = {well: data[0] if well not in ("C1",) else "Control" for well, data in EXAMPLES.items()}
# nicer visual labels for the plate map (short, readable in a small cell)
PLATE_LABEL = {
    "A1": "Hongo_A_Rep1", "A2": "Hongo_A_Rep2", "A3": "Hongo_A_TratX_R1", "A4": "Hongo_A_TratX_R2",
    "B1": "Hongo_B_Rep1", "B2": "Hongo_B_Rep2", "B3": "Hongo_B_TratX_R1", "B4": "Hongo_B_TratX_R2",
    "C1": "Control",
}

wb = Workbook()

# ===========================================================================
# 1) PLATE_MAP
# ===========================================================================
ws_map = wb.active
ws_map.title = "Plate_Map"

title_font   = Font(name=FONT_NAME, size=14, bold=True, color="1F1F1F")
axis_font    = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
well_font    = Font(name=FONT_NAME, size=8, color="1F1F1F")
example_font = Font(name=FONT_NAME, size=8, bold=True, color="1F1F1F")

axis_fill    = PatternFill("solid", fgColor="404040")
empty_fill   = PatternFill("solid", fgColor="FFFFFF")
sample_fill  = PatternFill("solid", fgColor="BDD7EE")
control_fill = PatternFill("solid", fgColor="F8CBAD")

thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

ws_map.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS) + 1)
ws_map.cell(row=1, column=1, value=f"{PLATE_ID} — Mapa de Placa (384 pocillos, 16x24)").font = title_font
ws_map.row_dimensions[1].height = 22

header_row = 3
first_data_row = 4

# column headers (1-24)
ws_map.cell(row=header_row, column=1, value="").fill = axis_fill
ws_map.cell(row=header_row, column=1).border = border
for j, col_num in enumerate(COLS, start=2):
    c = ws_map.cell(row=header_row, column=j, value=col_num)
    c.font = axis_font
    c.fill = axis_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = border

# row headers (A-P) + wells
for i, row_letter in enumerate(ROWS, start=0):
    r = first_data_row + i
    rh = ws_map.cell(row=r, column=1, value=row_letter)
    rh.font = axis_font
    rh.fill = axis_fill
    rh.alignment = Alignment(horizontal="center", vertical="center")
    rh.border = border

    for j, col_num in enumerate(COLS, start=2):
        well = f"{row_letter}{col_num}"
        c = ws_map.cell(row=r, column=j)
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.font = example_font if well in PLATE_LABEL else well_font
        if well in PLATE_LABEL:
            c.value = PLATE_LABEL[well]
            c.fill = control_fill if PLATE_LABEL[well] == "Control" else sample_fill
        else:
            c.fill = empty_fill

ws_map.column_dimensions["A"].width = 4
for j in range(2, len(COLS) + 2):
    ws_map.column_dimensions[get_column_letter(j)].width = 11
for i in range(len(ROWS)):
    ws_map.row_dimensions[first_data_row + i].height = 26

ws_map.freeze_panes = ws_map.cell(row=first_data_row, column=2)

# ===========================================================================
# 2) SAMPLE_METADATA
# ===========================================================================
ws_meta = wb.create_sheet("Sample_Metadata")

columns = ["plate_id", "well", "sample_id", "organism", "condition", "replicate", "filename", "notes"]
widths  = [12, 8, 12, 24, 24, 10, 26, 30]

header_font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="404040")
cell_font   = Font(name=FONT_NAME, size=11, color="1F1F1F")

for i, h in enumerate(columns, start=1):
    c = ws_meta.cell(row=1, column=i, value=h)
    c.font = header_font
    c.fill = header_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = border
    ws_meta.column_dimensions[get_column_letter(i)].width = widths[i-1]

r = 2
for row_letter in ROWS:
    for col_num in COLS:
        well = f"{row_letter}{col_num}"
        filename = f"{PLATE_ID}_{well}.mzML"
        if well in EXAMPLES:
            sample_id, organism, condition, replicate, notes = EXAMPLES[well]
        else:
            sample_id, organism, condition, replicate, notes = ("", "", "", "", "")

        rowdata = [PLATE_ID, well, sample_id, organism, condition, replicate, filename, notes]
        for i, val in enumerate(rowdata, start=1):
            c = ws_meta.cell(row=r, column=i, value=val)
            c.font = cell_font
            c.border = border
            c.alignment = Alignment(horizontal="left" if i not in (2, 6) else "center", vertical="center")
        r += 1

last_row = r - 1
last_col_letter = get_column_letter(len(columns))
table_ref = f"A1:{last_col_letter}{last_row}"
tab = Table(displayName="SampleMetadata384", ref=table_ref)
tab.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium2", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=False
)
ws_meta.add_table(tab)
ws_meta.freeze_panes = "A2"

output_path = OUTPUT_DIR / f"{PLATE_ID}.xlsx"
wb.save(output_path)
print(f"Listo: {last_row - 1} pocillos mapeados (A1 -> P24).")
print(f"Guardado en: {output_path}")