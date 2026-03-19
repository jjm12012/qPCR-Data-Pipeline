"""
parse_raw.py  –  Step 1 of the qPCR pipeline
Reads a CFX Maestro summary .xlsx export, normalises columns,
extracts run metadata, classifies sample types, and writes a
clean CSV suitable for downstream analysis.

Expected sheet layouts (auto-detected):
  Layout A: Sheet 0 = results table, Sheet 1 = Run Information
  Layout B: Sheet 0 = Run Information, Sheet 1 = results table

Results table columns: Well, Content, Sample, Cq, I.C. Cq, SQ, Result
  (blank first-column spacers are handled automatically)

Output CSV columns:
  run_id, run_date, assay_name, lot_number, extraction_lot,
  well, content, sample_label, sample_type, replicate,
  target_cq, ic_cq, sq, result, qc_flag
"""

import argparse
import re
import sys
import zipfile
from pathlib import Path

import pandas as pd


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _fix_xlsx(src: Path, tmp: Path) -> None:
    """
    Some CFX Maestro exports use lowercase zip-entry names or backslash
    separators. Rewrite the archive with standard names so openpyxl can read it.
    """
    remap = {}
    with zipfile.ZipFile(src) as z:
        for name in z.namelist():
            data = z.read(name)
            new_name = (
                name.replace("\\", "/")
                    .replace("[content_types]", "[Content_Types]")
                    .replace("sharedstrings", "sharedStrings")
            )
            remap[new_name] = data

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in remap.items():
            zout.writestr(name, data)


def _detect_sheets(xl: pd.ExcelFile) -> tuple:
    """
    Auto-detect which sheet holds the results table and which holds Run Information.
    Scans up to the first 2 sheets for a row containing 'well'.
    Returns (results_sheet_idx, run_info_sheet_idx).
    Falls back to (0, 1) if detection fails.
    """
    n_sheets = len(xl.sheet_names)
    for sheet_idx in range(min(2, n_sheets)):
        try:
            peek = pd.read_excel(xl, sheet_name=sheet_idx, header=None, nrows=30)
            for _, row in peek.iterrows():
                if any(str(v).strip().lower() == "well" for v in row):
                    run_info_idx = 1 - sheet_idx if n_sheets >= 2 else None
                    return sheet_idx, run_info_idx
        except Exception:
            continue
    return 0, 1 if n_sheets >= 2 else None


def _read_run_info(xl: pd.ExcelFile, sheet_idx=1) -> dict:
    """Parse the Run Information sheet into a dict."""
    info = {}
    if sheet_idx is None:
        return info
    try:
        df = pd.read_excel(xl, sheet_name=sheet_idx, header=None)
        for _, row in df.iterrows():
            if pd.notna(row[0]) and pd.notna(row[1]):
                info[str(row[0]).strip()] = str(row[1]).strip()
    except Exception:
        pass
    return info


def _parse_sample_type(label) -> tuple:
    """
    Derive a normalised sample_type from a raw sample label.
    Returns (sample_type_key, label_clean).
    """
    if pd.isna(label) or str(label).strip() == "":
        return ("Unknown", "")

    label = str(label).strip()
    label_clean = re.sub(r"\s+", " ", label)
    label_clean = re.sub(r"(\d+\.?\d*)\s*mL", r"\1mL", label_clean, flags=re.IGNORECASE)
    sample_type = label_clean.replace(" ", "_")
    return sample_type, label_clean


def _read_results_sheet(xl: pd.ExcelFile, sheet_idx: int = 0) -> pd.DataFrame:
    """
    Read the primary results sheet.
    Uses positional indexing to handle duplicate column names (e.g. multiple 'Cq').
    Blank/spacer columns (empty header) are skipped automatically.
    Returns a DataFrame with clean, standardised column names.
    """
    raw = pd.read_excel(xl, sheet_name=sheet_idx, header=None)

    # Drop fully-empty rows and columns
    raw.dropna(how="all", inplace=True)
    raw.dropna(axis=1, how="all", inplace=True)
    raw.reset_index(drop=True, inplace=True)

    # Locate header row (contains 'Well')
    header_row_idx = None
    for i, row in raw.iterrows():
        if any(str(v).strip().lower() == "well" for v in row):
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("Could not locate header row containing 'Well'.")

    # Extract raw header labels as a list (preserves duplicate names)
    raw_headers = [str(v).strip() for v in raw.iloc[header_row_idx]]

    # Data rows below the header, reset to integer column positions
    data = raw.iloc[header_row_idx + 1:].copy()
    data.reset_index(drop=True, inplace=True)
    data.columns = list(range(len(raw_headers)))

    # Skip blank-header columns (spacers like an empty column A)
    raw_headers_filtered = [
        (pos, name) for pos, name in enumerate(raw_headers)
        if name.strip() not in ("", "nan", "NaN")
    ]

    # Map positions: first match wins for each target column
    # I.C. Cq must be checked before plain Cq to avoid mis-assignment
    col_pos = {}
    for pos, name in raw_headers_filtered:
        lo = name.lower()
        if ("i.c" in lo or "ic" in lo) and "cq" in lo:
            if "ic_cq" not in col_pos:
                col_pos["ic_cq"] = pos
        elif lo == "cq":
            if "target_cq" not in col_pos:
                col_pos["target_cq"] = pos
        elif lo == "well":
            if "well" not in col_pos:
                col_pos["well"] = pos
        elif lo == "content":
            if "content" not in col_pos:
                col_pos["content"] = pos
        elif lo == "sample":
            if "sample_label" not in col_pos:
                col_pos["sample_label"] = pos
        elif lo == "sq":
            if "sq" not in col_pos:
                col_pos["sq"] = pos
        elif lo == "result":
            if "result" not in col_pos:
                col_pos["result"] = pos

    # Build output DataFrame column by column
    final_cols = ["well", "content", "sample_label", "target_cq", "ic_cq", "sq", "result"]
    out = {}
    for dest in final_cols:
        if dest in col_pos:
            out[dest] = data[col_pos[dest]].values
        else:
            out[dest] = pd.NA

    df = pd.DataFrame(out)

    # Cast numeric columns
    for col in ["target_cq", "ic_cq", "sq"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Drop rows where 'well' is NaN (blank trailing rows)
    df = df[df["well"].notna()].copy()
    df.reset_index(drop=True, inplace=True)

    return df, col_pos


def _validate_structure(xl: pd.ExcelFile, results_idx: int,
                        run_info_idx, col_pos: dict) -> list:
    """
    Return a list of issue strings found in the xlsx structure.
    Used by --dry-run mode.
    """
    issues = []
    n = len(xl.sheet_names)

    if n < 1:
        issues.append("CRITICAL: Workbook has no sheets.")
        return issues

    print(f"  Sheets found ({n}): {xl.sheet_names}")
    print(f"  Results sheet  → index {results_idx} ({xl.sheet_names[results_idx]})")
    if run_info_idx is not None and run_info_idx < n:
        print(f"  Run Info sheet → index {run_info_idx} ({xl.sheet_names[run_info_idx]})")
    else:
        issues.append("WARNING: Run Information sheet not found; metadata will be blank.")

    required = ["well", "target_cq", "ic_cq", "result"]
    for req in required:
        if req not in col_pos:
            issues.append(f"CRITICAL: Expected column '{req}' not found in results sheet.")

    found_cols = sorted(col_pos.keys())
    print(f"  Columns mapped: {found_cols}")
    return issues


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Parse raw qPCR .xlsx export to clean CSV.")
    parser.add_argument("--input",    required=True, help="Path to raw .xlsx file")
    parser.add_argument("--output",   required=True, help="Path for output .csv file")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Validate xlsx structure and report issues; do not write output.")
    args = parser.parse_args()

    src = Path(args.input)
    out = Path(args.output)

    # Fix non-standard zip structure if needed
    tmp = src.parent / (src.stem + "_tmp_fixed.xlsx")
    try:
        _fix_xlsx(src, tmp)
        xl = pd.ExcelFile(tmp)
    except Exception as e:
        print(f"  [parse] Warning: xlsx fix failed ({e}), trying direct read.")
        xl = pd.ExcelFile(src)

    # Auto-detect sheet layout
    results_idx, run_info_idx = _detect_sheets(xl)
    print(f"  [parse] Detected layout: results=sheet[{results_idx}], "
          f"run_info=sheet[{run_info_idx}]")

    # --- Run metadata ---
    run_info      = _read_run_info(xl, sheet_idx=run_info_idx)
    run_id        = src.stem
    run_date      = run_info.get("Run Started", "")
    assay_name    = run_info.get("Assay Name", "")
    lot_number    = run_info.get("Lot Number", "")
    extraction_lot = run_info.get("Extraction Lot Number", "")

    # --- Results data ---
    try:
        df, col_pos = _read_results_sheet(xl, sheet_idx=results_idx)
    except ValueError as e:
        print(f"  [parse] ERROR reading results sheet: {e}")
        xl.close()
        if tmp.exists():
            tmp.unlink()
        sys.exit(1)

    # --- Dry-run: validate and exit without writing ---
    if args.dry_run:
        print(f"\n  [dry-run] Validating: {src.name}")
        issues = _validate_structure(xl, results_idx, run_info_idx, col_pos)
        print(f"  Data rows found: {len(df)}")
        if issues:
            print("\n  Issues found:")
            for issue in issues:
                print(f"    • {issue}")
            print("\n  [dry-run] FAILED — see issues above.")
            xl.close()
            if tmp.exists():
                tmp.unlink()
            sys.exit(1)
        else:
            print("\n  [dry-run] OK — no structural issues found. No CSV written.")
            xl.close()
            if tmp.exists():
                tmp.unlink()
            sys.exit(0)

    # --- Derive sample type + replicate (graceful row skipping) ---
    sample_types  = []
    labels_clean  = []
    skipped_rows  = 0
    for idx, row in df.iterrows():
        try:
            st, lc = _parse_sample_type(row["sample_label"])
        except Exception as e:
            print(f"  [parse] Warning: skipping row {idx} – {e}")
            st, lc = "Unknown", ""
            skipped_rows += 1
        sample_types.append(st)
        labels_clean.append(lc)

    if skipped_rows:
        print(f"  [parse] Warning: {skipped_rows} malformed row(s) skipped.")

    df["sample_type"]        = sample_types
    df["sample_label_clean"] = labels_clean
    df["replicate"] = df.groupby("sample_type").cumcount() + 1

    # --- Attach run metadata ---
    df.insert(0, "extraction_lot", extraction_lot)
    df.insert(0, "lot_number",     lot_number)
    df.insert(0, "assay_name",     assay_name)
    df.insert(0, "run_date",       run_date)
    df.insert(0, "run_id",         run_id)

    # --- Flag controls ---
    df["is_control"] = df["content"].astype(str).str.strip().str.lower().isin(
        ["pos ctrl", "neg ctrl"]
    )

    # --- Write output ---
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out, index=False)

    # Cleanup temp file
    xl.close()
    if tmp.exists():
        tmp.unlink()

    n_total = len(df)
    n_ctrl  = df["is_control"].sum()
    n_types = df.loc[~df["is_control"], "sample_type"].nunique()
    print(f"  [parse] {src.name}: {n_total} wells → {n_types} sample types, {n_ctrl} controls")
    print(f"          → {out}")


if __name__ == "__main__":
    main()
