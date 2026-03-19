"""
make_test_data.py  –  Generate a synthetic CFX Maestro .xlsx for pipeline testing.

Run from the project root:
    python scripts/make_test_data.py

Writes:  data/raw/test_run_001.xlsx

Expected pipeline outputs (ground truth)
-----------------------------------------
Sample types & positivity:
  Auto_Neg    4 wells  0 positive   0.0 %
  Auto_Pos    4 wells  4 positive 100.0 %   mean Target Cq ≈ 27.15  (one Borderline well)
  Manual_Neg  4 wells  0 positive   0.0 %
  Manual_Pos  4 wells  4 positive 100.0 %   mean Target Cq ≈ 22.95

QC flags expected (qc_flags.csv):
  C04  Auto_Pos   Borderline: Target Cq > 35       (Cq = 36.1)
  D04  Auto_Neg   IC extraction failure             (I.C. Cq = 37.8, plate mean ≈ 28.1, SD ≈ 0.2)

Figures expected:
  01  Positivity rate bar  – Auto_Pos 100%, Manual_Pos 100%, both Neg 0%
  02  Target Cq boxplot    – two groups (Auto_Pos, Manual_Pos); Auto_Pos has one high outlier
  03  I.C. Cq boxplot      – tight cluster ~28 for all sample types
  04  Cq scatter           – positives bottom-left, negatives missing target Cq
  05  Plate result heatmap – green A01-A04, C01-C04; red B01-B04, D01-D04; blue controls
  06  Plate Cq heatmap     – low Cq (dark green) A row, one bright cell at C04
  08  I.C. Cq control chart – flat line ~28 with single spike at D04

Excel report expected:
  Per-Run Detail  – C04 Target Cq cell yellow; D04 I.C. Cq cell light red
  QC Flags sheet  – 2 rows, both highlighted red
"""

from pathlib import Path
import openpyxl
from openpyxl import Workbook


# ---------------------------------------------------------------------------
# Well data
# ---------------------------------------------------------------------------
# Columns: Well, Content, Sample, Cq, I.C. Cq, SQ, Result
WELLS = [
    # --- Manual Positives (4 wells, all Positive, tight Cq ~22-23) ---
    ("A01", "Unkn", "Manual_Pos",  22.5,  28.1,  "",  "Positive"),
    ("A02", "Unkn", "Manual_Pos",  23.1,  27.9,  "",  "Positive"),
    ("A03", "Unkn", "Manual_Pos",  22.8,  28.3,  "",  "Positive"),
    ("A04", "Unkn", "Manual_Pos",  23.4,  28.0,  "",  "Positive"),

    # --- Manual Negatives (4 wells, all Negative, no target Cq) ---
    ("B01", "Unkn", "Manual_Neg",  "",    28.2,  "",  "Negative"),
    ("B02", "Unkn", "Manual_Neg",  "",    27.8,  "",  "Negative"),
    ("B03", "Unkn", "Manual_Neg",  "",    28.4,  "",  "Negative"),
    ("B04", "Unkn", "Manual_Neg",  "",    28.1,  "",  "Negative"),

    # --- Auto Positives (3 normal + 1 BORDERLINE Cq > 35) ---
    ("C01", "Unkn", "Auto_Pos",    24.2,  28.0,  "",  "Positive"),
    ("C02", "Unkn", "Auto_Pos",    23.8,  27.9,  "",  "Positive"),
    ("C03", "Unkn", "Auto_Pos",    24.5,  28.2,  "",  "Positive"),
    ("C04", "Unkn", "Auto_Pos",    36.1,  28.1,  "",  "Positive"),  # ← BORDERLINE

    # --- Auto Negatives (3 normal + 1 IC EXTRACTION FAILURE) ---
    ("D01", "Unkn", "Auto_Neg",    "",    28.3,  "",  "Negative"),
    ("D02", "Unkn", "Auto_Neg",    "",    27.7,  "",  "Negative"),
    ("D03", "Unkn", "Auto_Neg",    "",    28.0,  "",  "Negative"),
    ("D04", "Unkn", "Auto_Neg",    "",    37.8,  "",  "Negative"),  # ← IC FAILURE

    # --- Plate controls ---
    ("E01", "Pos Ctrl", "Pos Control",  20.1,  28.0,  "",  "Valid Ctrl"),
    ("E02", "Neg Ctrl", "Neg Control",  "",    28.1,  "",  "Valid Ctrl"),
]

# ---------------------------------------------------------------------------
# Run Information (key-value pairs for Sheet 1)
# ---------------------------------------------------------------------------
RUN_INFO = [
    ("Run Started",          "2026-03-13 09:00:00"),
    ("Assay Name",           "TBV_CellQuant_Test"),
    ("Lot Number",           "LOT-TEST-001"),
    ("Extraction Lot Number","EXT-TEST-001"),
    ("Operator",             "Test User"),
    ("Instrument Serial",    "TEST-SN-12345"),
]


# ---------------------------------------------------------------------------
# Build workbook
# ---------------------------------------------------------------------------
def make_test_xlsx(out_path: Path):
    wb = Workbook()

    # ---- Sheet 0: Results table ----------------------------------------
    ws_results = wb.active
    ws_results.title = "IDE Summary Results"

    headers = ["Well", "Content", "Sample", "Cq", "I.C. Cq", "SQ", "Result"]
    ws_results.append(headers)

    for row in WELLS:
        ws_results.append(list(row))

    # ---- Sheet 1: Run Information ---------------------------------------
    ws_info = wb.create_sheet("Run Information")
    for key, val in RUN_INFO:
        ws_info.append([key, val])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    root    = Path(__file__).resolve().parent.parent
    out     = root / "data" / "raw" / "test_run_001.xlsx"
    made    = make_test_xlsx(out)

    print(f"\n  [test data] Written → {made}")
    print(f"  Wells: {len(WELLS)}  ({len([w for w in WELLS if w[1] not in ('Pos Ctrl','Neg Ctrl')])} samples + 2 controls)")

    print("""
Expected pipeline results
=========================
summary_statistics.csv
  Sample Type   N  Positive  Rate     Mean Target Cq
  Auto_Neg      4     0      0.0 %    —
  Auto_Pos      4     4    100.0 %    27.150
  Manual_Neg    4     0      0.0 %    —
  Manual_Pos    4     4    100.0 %    22.950

qc_flags.csv  (2 rows expected)
  run_id           well  sample_type  flag_reason
  test_run_001     C04   Auto_Pos     Borderline: Target Cq > 35
  test_run_001     D04   Auto_Neg     IC extraction failure: I.C. Cq=37.80 ...

Excel report
  Per-Run Detail row C04 → Target Cq cell YELLOW
  Per-Run Detail row D04 → I.C. Cq cell LIGHT RED
  QC Flags sheet       → 2 rows, all highlighted red
""")


if __name__ == "__main__":
    main()
