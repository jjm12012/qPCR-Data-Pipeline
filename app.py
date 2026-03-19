"""
app.py — Streamlit web app for the qPCR automated pipeline.

Wraps parse_raw, analyze, visualize, and report modules to provide
a drag-and-drop interface with a live dashboard and downloadable outputs.

Deploy free at: https://share.streamlit.io
"""

import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook

# ── Make the scripts importable from the same directory ──────────────────────
sys.path.insert(0, str(Path(__file__).parent))
import parse_raw
import analyze
import visualize
import report as rpt

# ─────────────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="qPCR Pipeline",
    page_icon="🧬",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────────────────────
# Figure caption map
# ─────────────────────────────────────────────────────────────────────────────
_CAPTIONS = {
    "01_positivity_rate_bar":    "Positivity Rate by Sample Type",
    "04_cq_scatter_by_well":     "Target Cq vs. I.C. Cq Scatter",
    "05_plate_heatmap_result":   "Plate Result Map",
    "06_plate_heatmap_cq":       "Plate Target Cq Heatmap",
    "07_run_positivity_heatmap": "Run × Sample-Type Positivity Heatmap",
    "08_ic_cq_control_chart":    "Cq Spatial & Sample-Type Overview",
    "09_interrun_cv_bar":        "Inter-run CV% by Sample Type",
}


def _caption(fname: str) -> str:
    stem = Path(fname).stem
    for prefix, caption in _CAPTIONS.items():
        if stem.startswith(prefix):
            return caption
    return stem.replace("_", " ").title()


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline runner
# ─────────────────────────────────────────────────────────────────────────────
def run_pipeline(uploaded_files, progress_bar, status_text: st.empty):
    """
    Run the full 4-step pipeline on the uploaded files.
    Returns a dict of all outputs held in memory so the temp dir can be cleaned up.
    """
    n_files = len(uploaded_files)

    with tempfile.TemporaryDirectory() as _tmpdir:
        tmpdir   = Path(_tmpdir)
        proc_dir = tmpdir / "processed"
        fig_dir  = tmpdir / "figures"
        rep_dir  = tmpdir / "reports"
        for d in [proc_dir, fig_dir, rep_dir]:
            d.mkdir()

        # ── Step 1: Parse each uploaded file ─────────────────────────────────
        for i, uf in enumerate(uploaded_files):
            status_text.text(f"📂 Parsing {uf.name}  ({i + 1} / {n_files})…")

            raw_path  = tmpdir / uf.name
            raw_path.write_bytes(uf.read())

            tmp_fixed = tmpdir / (raw_path.stem + "_fixed.xlsx")
            try:
                parse_raw._fix_xlsx(raw_path, tmp_fixed)
                xl = pd.ExcelFile(tmp_fixed)
            except Exception:
                xl = pd.ExcelFile(raw_path)

            results_idx, run_info_idx = parse_raw._detect_sheets(xl)
            run_info = parse_raw._read_run_info(xl, run_info_idx)
            df, _    = parse_raw._read_results_sheet(xl, results_idx)

            # Derive sample types
            stypes, labels = [], []
            for _, row in df.iterrows():
                st_val, lc = parse_raw._parse_sample_type(row["sample_label"])
                stypes.append(st_val)
                labels.append(lc)

            df["sample_type"]        = stypes
            df["sample_label_clean"] = labels
            df["replicate"]          = df.groupby("sample_type").cumcount() + 1

            df.insert(0, "extraction_lot", run_info.get("Extraction Lot Number", ""))
            df.insert(0, "lot_number",     run_info.get("Lot Number", ""))
            df.insert(0, "assay_name",     run_info.get("Assay Name", ""))
            df.insert(0, "run_date",       run_info.get("Run Started", ""))
            df.insert(0, "run_id",         raw_path.stem)

            df["is_control"] = (
                df["content"].astype(str).str.strip().str.lower()
                .isin(["pos ctrl", "neg ctrl"])
            )

            df.to_csv(proc_dir / (raw_path.stem + "_parsed.csv"), index=False)

            xl.close()
            if tmp_fixed.exists():
                tmp_fixed.unlink()

            progress_bar.progress((i + 1) / (n_files * 4))

        # ── Step 2: Analyze ───────────────────────────────────────────────────
        status_text.text("📊 Running analysis…")

        combined   = analyze.load_all_parsed(proc_dir)
        flags_high = analyze.flag_high_cq(combined)
        flags_ic   = analyze.flag_ic_extraction(combined)
        qc_flags   = pd.concat([flags_high, flags_ic], ignore_index=True)

        combined   = analyze._annotate_qc_flags(combined, qc_flags)
        stats      = analyze.compute_stats(combined)
        ctrl_stats = analyze.compute_control_stats(combined)
        run_cmp    = analyze.compare_runs(combined)
        cv         = analyze.compute_interrun_cv(combined)

        # Write CSVs (needed by report functions that reference the dir)
        qc_flags.to_csv(rep_dir / "qc_flags.csv", index=False)
        stats.to_csv(rep_dir / "summary_statistics.csv", index=False)
        if not ctrl_stats.empty:
            ctrl_stats.to_csv(rep_dir / "control_statistics.csv", index=False)
        run_cmp.to_csv(rep_dir / "run_comparison.csv", index=False)
        if not cv.empty:
            cv.to_csv(rep_dir / "interrun_cv.csv", index=False)
        combined.to_csv(rep_dir / "all_runs_combined.csv", index=False)

        progress_bar.progress(0.50)

        # ── Step 3: Visualize ─────────────────────────────────────────────────
        status_text.text("📈 Generating figures…")

        visualize.fig_positivity_bar(stats, fig_dir)
        visualize.fig_cq_scatter(combined, fig_dir)
        visualize.fig_plate_heatmap_result(combined, fig_dir)
        visualize.fig_plate_heatmap_cq(combined, fig_dir)
        visualize.fig_run_heatmap(combined, fig_dir)
        visualize.fig_ic_control_chart(combined, fig_dir)
        if not cv.empty:
            visualize.fig_interrun_cv_bar(cv, fig_dir)

        progress_bar.progress(0.75)

        # ── Step 4: Build reports ─────────────────────────────────────────────
        status_text.text("📝 Building reports…")

        wb = Workbook()
        wb.remove(wb.active)
        rpt._sheet_summary(wb, stats, combined,
                           ctrl_stats if not ctrl_stats.empty else None)
        rpt._sheet_detail(wb, combined, qc_flags=qc_flags)
        rpt._sheet_plate_layout(wb, combined)
        rpt._sheet_qc_flags(wb, qc_flags)
        rpt._sheet_stats(wb, stats)
        rpt._sheet_run_comparison(wb, run_cmp)
        rpt._sheet_controls(wb, ctrl_stats if not ctrl_stats.empty else None)
        rpt._sheet_notes(wb)

        excel_path = rep_dir / "qpcr_summary_report.xlsx"
        wb.save(excel_path)

        timestamp = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
        html_path = rpt._write_html_report(
            fig_dir=fig_dir,
            out_dir=rep_dir,
            stats=stats,
            combined=combined,
            qc_flags=qc_flags,
            timestamp=timestamp,
        )

        progress_bar.progress(1.0)

        # ── Collect everything into memory before the temp dir is cleaned up ──
        figures = {
            p.name: p.read_bytes()
            for p in sorted(fig_dir.glob("*.png"))
        }
        parsed_csvs = {
            p.name: p.read_bytes()
            for p in sorted(proc_dir.glob("*_parsed.csv"))
        }

        return {
            "combined":       combined,
            "stats":          stats,
            "ctrl_stats":     ctrl_stats,
            "qc_flags":       qc_flags,
            "run_comparison": run_cmp,
            "cv":             cv,
            "figures":        figures,
            "excel_bytes":    excel_path.read_bytes(),
            "html_bytes":     html_path.read_bytes(),
            "combined_csv":   (rep_dir / "all_runs_combined.csv").read_bytes(),
            "parsed_csvs":    parsed_csvs,
        }


# ─────────────────────────────────────────────────────────────────────────────
# UI — Header
# ─────────────────────────────────────────────────────────────────────────────
st.title("🧬 qPCR Pipeline")
st.markdown(
    "Upload one or more CFX Maestro **`- IDE Summary`** `.xlsx` files to run the full qPCR analysis pipeline. "
    "The app parses your run data, computes positivity rates and QC flags, generates figures, "
    "and compiles a ready-to-download Excel and HTML report — all in one click."
)

# ─────────────────────────────────────────────────────────────────────────────
# UI — File upload
# ─────────────────────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload .xlsx files",
    type=["xlsx"],
    accept_multiple_files=True,
    help="CFX Maestro summary exports (.xlsx). Multiple files are treated as separate runs.",
)

if uploaded_files:
    names = [f.name for f in uploaded_files]
    st.caption(f"**{len(names)} file(s) ready:** {', '.join(names)}")

    if st.button("▶ Run Pipeline", type="primary"):
        prog   = st.progress(0)
        status = st.empty()
        try:
            results = run_pipeline(uploaded_files, prog, status)
            st.session_state["results"]   = results
            st.session_state["run_names"] = names
        except Exception as e:
            st.error(f"**Pipeline error:** {e}")
            st.exception(e)
        finally:
            prog.empty()
            status.empty()

# ─────────────────────────────────────────────────────────────────────────────
# UI — Results dashboard
# ─────────────────────────────────────────────────────────────────────────────
if "results" in st.session_state:
    res      = st.session_state["results"]
    combined = res["combined"]
    stats    = res["stats"]
    qc_flags = res["qc_flags"]

    n_runs    = combined["run_id"].nunique()
    n_samples = len(combined[~combined["is_control"]])
    n_pos     = int(
        (combined[~combined["is_control"]]["result"]
         .str.strip().str.lower() == "positive").sum()
    )
    rate    = round(n_pos / n_samples * 100, 1) if n_samples else 0
    n_flags = len(qc_flags)

    # ── Metric cards ──────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Runs Processed", n_runs)
    c2.metric("Sample Wells",   n_samples)
    c3.metric("Positivity Rate", f"{rate}%")
    c4.metric("QC Flags", f"{n_flags} ⚠️" if n_flags else "0 ✅")

    st.divider()

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab_summary, tab_figures, tab_qc, tab_dl = st.tabs([
        "📊 Summary", "📈 Charts & Figures", "⚠️ QC Flags", "📥 Downloads",
    ])

    # ── Summary tab ───────────────────────────────────────────────────────────
    with tab_summary:
        st.subheader("Sample Type Statistics")
        display_stats = stats.rename(columns={
            "sample_type":         "Sample Type",
            "n_wells":             "N Wells",
            "n_positive":          "N Positive",
            "n_negative":          "N Negative",
            "positivity_rate_pct": "Positivity %",
            "mean_target_cq":      "Mean Target Cq",
            "sd_target_cq":        "SD Target Cq",
            "mean_ic_cq":          "Mean I.C. Cq",
            "sd_ic_cq":            "SD I.C. Cq",
        })
        st.dataframe(display_stats, use_container_width=True, hide_index=True)

        if not res["ctrl_stats"].empty:
            st.subheader("Control Statistics")
            st.dataframe(res["ctrl_stats"], use_container_width=True, hide_index=True)

        if not res["cv"].empty:
            st.subheader("Inter-run CV%")
            st.dataframe(res["cv"], use_container_width=True, hide_index=True)

    # ── Charts & Figures tab (all figures in one place) ───────────────────────
    with tab_figures:
        # Figure display order: summary charts → analytical charts → plate maps
        prefix_order = ["01_", "04_", "07_", "08_", "09_", "05_", "06_"]

        def _sort_key(fname):
            stem = Path(fname).stem
            for i, prefix in enumerate(prefix_order):
                if stem.startswith(prefix):
                    return (i, fname)
            return (len(prefix_order), fname)

        sorted_figs = sorted(res["figures"].items(), key=lambda x: _sort_key(x[0]))

        if sorted_figs:
            for name, img_bytes in sorted_figs:
                st.subheader(_caption(name))
                st.image(img_bytes, use_container_width=True)
                st.divider()
        else:
            st.info("No figures were generated.")

    # ── QC Flags tab ──────────────────────────────────────────────────────────
    with tab_qc:
        if qc_flags.empty:
            st.success("✅ No QC flags — all wells passed.")
        else:
            st.warning(f"⚠️ {n_flags} well(s) flagged across all runs.")
            display_qc = qc_flags.rename(columns={
                "run_id":      "Run ID",
                "well":        "Well",
                "sample_type": "Sample Type",
                "flag_reason": "Flag Reason",
            })
            st.dataframe(display_qc, use_container_width=True, hide_index=True)

    # ── Downloads tab ─────────────────────────────────────────────────────────
    with tab_dl:
        st.subheader("Reports")
        col_a, col_b, col_c = st.columns(3)
        col_a.download_button(
            "📥 Excel Report (.xlsx)",
            data=res["excel_bytes"],
            file_name="qpcr_summary_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        col_b.download_button(
            "📥 HTML Report (.html)",
            data=res["html_bytes"],
            file_name="qpcr_figures_report.html",
            mime="text/html",
            use_container_width=True,
        )
        col_c.download_button(
            "📥 Combined Data (.csv)",
            data=res["combined_csv"],
            file_name="all_runs_combined.csv",
            mime="text/csv",
            use_container_width=True,
        )

        st.subheader("Individual Figures")
        fig_cols = st.columns(3)
        for i, (name, img_bytes) in enumerate(sorted(res["figures"].items())):
            fig_cols[i % 3].download_button(
                f"📥 {_caption(name)}",
                data=img_bytes,
                file_name=name,
                mime="image/png",
                key=f"dl_fig_{name}",
                use_container_width=True,
            )

        if res["parsed_csvs"]:
            st.subheader("Per-Run Parsed CSVs")
            csv_cols = st.columns(min(len(res["parsed_csvs"]), 3))
            for i, (name, csv_bytes) in enumerate(sorted(res["parsed_csvs"].items())):
                csv_cols[i % len(csv_cols)].download_button(
                    f"📥 {name}",
                    data=csv_bytes,
                    file_name=name,
                    mime="text/csv",
                    key=f"dl_csv_{name}",
                    use_container_width=True,
                )
