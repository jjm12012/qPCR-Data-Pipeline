# qPCR Pipeline Web App

A Streamlit web app that wraps the 4-step qPCR automation pipeline into a drag-and-drop interface with a live results dashboard.

## What it does

Upload one or more **CFX Maestro `.xlsx` exports** and the app will:

1. **Parse** — normalise raw exports, extract run metadata, classify sample types
2. **Analyze** — compute per-sample-type statistics, positivity rates, QC flags, and inter-run CV%
3. **Visualise** — generate plate heatmaps, positivity bar charts, Cq scatter plots, and more
4. **Report** — compile an Excel workbook and a self-contained HTML report

Results appear immediately in the browser across five tabs: **Summary**, **Plate Maps**, **Charts**, **QC Flags**, and **Downloads**.

---

## Files

```
app.py              ← Streamlit entry point
requirements.txt    ← Python dependencies
.streamlit/
  config.toml       ← Theme and upload size settings
parse_raw.py        ← Step 1: parse raw .xlsx exports
analyze.py          ← Step 2: statistics and QC flagging
visualize.py        ← Step 3: matplotlib figures
report.py           ← Step 4: Excel + HTML report generation
make_test_data.py   ← Generates a synthetic test .xlsx for development
```

---

## Run locally

```bash
pip install streamlit pandas numpy matplotlib openpyxl

streamlit run app.py
```

Then open [http://localhost:8501](http://localhost:8501) in your browser.

---

## Deploy free on Streamlit Community Cloud

1. Push this folder to a GitHub repository (all files at the repo root)
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
3. Click **New app** → select your repo → set main file to `app.py`
4. Click **Deploy** — your app will be live at a public URL in ~2 minutes

> The app is stateless: no data is stored between sessions.

---

## Input format

The app expects **CFX Maestro `- IDE Summary` `.xlsx` exports** — the summary file exported directly from CFX Maestro with "- IDE Summary" in the filename.

Each file should contain two sheets:
- A results table with columns: `Well`, `Content`, `Sample`, `Cq`, `I.C. Cq`, `SQ`, `Result`
- A Run Information sheet with metadata (assay name, lot numbers, run date)

Sheet order is auto-detected. Multiple files uploaded together are treated as separate runs and analyzed jointly for inter-run statistics.

---

## QC flags

| Flag | Condition |
|---|---|
| Borderline (yellow) | Target Cq > 35 |
| IC extraction failure (red) | I.C. Cq > 2 SD from plate mean |
