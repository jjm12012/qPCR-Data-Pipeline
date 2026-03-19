"""
visualize.py  –  Step 3 of the qPCR pipeline
Generates publication-ready figures comparing sample types.

Figures produced:
  01_positivity_rate_bar.png         – Positivity rate (%) per sample type
  04_cq_scatter_by_well.png          – Target Cq vs. I.C. Cq scatter, coloured by result
  05_plate_heatmap_result.png        – Plate layout heatmap (Result)
  06_plate_heatmap_cq.png            – Plate layout heatmap (Target Cq)
  07_run_positivity_heatmap.png      – Run × sample-type positivity heatmap (multi-run)
  08_ic_cq_control_chart.png         – Cq Spatial & Sample-Type Overview (plate row distribution + sample-type distribution)
  09_interrun_cv_bar.png             – Inter-run CV% per sample type (multi-run only)
"""

import argparse
import re
import textwrap
import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Style constants
# ---------------------------------------------------------------------------
PALETTE = {
    "Positive":   "#2ecc71",
    "Negative":   "#e74c3c",
    "Valid Ctrl": "#3498db",
    "Unknown":    "#95a5a6",
}
SAMPLE_PALETTE = [
    "#4e79a7", "#f28e2b", "#e15759", "#76b7b2",
    "#59a14f", "#edc948", "#b07aa1", "#ff9da7",
]
FIG_DPI = 300
FONT_FAMILY = "DejaVu Sans"


def _savefig(fig, path: Path, tight: bool = True):
    if tight:
        fig.tight_layout()
    fig.savefig(path, dpi=FIG_DPI, bbox_inches="tight")
    plt.close(fig)
    print(f"  [viz] → {path.name}")


# ---------------------------------------------------------------------------
# Figure 1: Positivity rate bar chart
# ---------------------------------------------------------------------------
def fig_positivity_bar(stats: pd.DataFrame, out_dir: Path):
    sample_types = stats["sample_type"].tolist()
    rates        = stats["positivity_rate_pct"].tolist()
    n_wells      = stats["n_wells"].tolist()
    n_positive   = stats["n_positive"].tolist()

    # Assign colours cyclically
    colours = [SAMPLE_PALETTE[i % len(SAMPLE_PALETTE)] for i in range(len(sample_types))]

    fig, ax = plt.subplots(figsize=(max(6, len(sample_types) * 1.2), 5))
    bars = ax.bar(sample_types, rates, color=colours, edgecolor="white", linewidth=0.8)

    # Annotate bars with positive count and rate
    for bar, rate, n_pos, n_tot in zip(bars, rates, n_positive, n_wells):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 1.5,
            f"{n_pos}/{n_tot}\n({rate:.1f}%)",
            ha="center", va="bottom", fontsize=8.5,
        )

    ax.set_ylim(0, 115)
    ax.set_ylabel("Positivity Rate (%)", fontsize=11)
    ax.set_xlabel("Sample Type", fontsize=11)
    ax.set_title("Positivity Rate by Sample Type", fontsize=13, fontweight="bold")
    ax.set_xticklabels(sample_types, rotation=30, ha="right", fontsize=9)
    ax.spines[["top", "right"]].set_visible(False)

    _savefig(fig, out_dir / "01_positivity_rate_bar.png")




# ---------------------------------------------------------------------------
# Figure 4: Target Cq vs I.C. Cq scatter
# ---------------------------------------------------------------------------
def fig_cq_scatter(df: pd.DataFrame, out_dir: Path):
    scatter_df = df[
        (~df["is_control"]) &
        df["target_cq"].notna() &
        df["ic_cq"].notna()
    ].copy()

    if scatter_df.empty:
        print("  [viz] Skipping scatter – insufficient data.")
        return

    fig, ax = plt.subplots(figsize=(7, 5))

    for result, colour in PALETTE.items():
        subset = scatter_df[scatter_df["result"].str.strip() == result]
        if subset.empty:
            continue
        ax.scatter(
            subset["target_cq"], subset["ic_cq"],
            color=colour, alpha=0.7, s=45, label=result, edgecolors="white", lw=0.4,
        )

    ax.set_xlabel("Cq", fontsize=11)
    ax.set_ylabel("I.C. Cq", fontsize=11)
    ax.set_title("Cq vs. Internal Control Cq", fontsize=13, fontweight="bold")
    ax.legend(title="Result", fontsize=9, title_fontsize=9)
    ax.spines[["top", "right"]].set_visible(False)

    _savefig(fig, out_dir / "04_cq_scatter_by_well.png")


# ---------------------------------------------------------------------------
# Figure 5 & 6: Plate layout heatmaps
# ---------------------------------------------------------------------------
def _well_to_rowcol(well: str):
    """Convert 'A01' → (0, 0), 'H12' → (7, 11)."""
    match = re.match(r"([A-H])(\d+)", str(well).strip())
    if not match:
        return None, None
    row = ord(match.group(1)) - ord("A")
    col = int(match.group(2)) - 1
    return row, col


def _well_label(row: pd.Series) -> str:
    """Return the best display label for a well: sample_label if present, else content."""
    label = row.get("sample_label", "")
    if pd.isna(label) or str(label).strip() in ("", "nan"):
        label = row.get("content", "")
    return str(label).strip() if pd.notna(label) else ""


def _wrap_well_text(text: str, width: int = 10) -> str:
    """Wrap a sample label to fit inside a well cell."""
    return textwrap.fill(text, width=width) if text else ""


def fig_plate_heatmap_result(df: pd.DataFrame, out_dir: Path):
    # Numeric encoding for each result category
    result_map = {
        "positive":    -1,
        "negative":     0,
        "valid ctrl":   1,
        "inhibition":   2,
    }
    # Colours in the same order: Positive, Negative, Valid Ctrl, Inhibition
    result_colors = {
        "Positive":   "#e74c3c",
        "Negative":   "#2ecc71",
        "Valid Ctrl": "#3498db",
        "Inhibition": "#f39c12",
    }
    cmap   = matplotlib.colors.ListedColormap(list(result_colors.values()))
    bounds = [-1.5, -0.5, 0.5, 1.5, 2.5]
    norm   = matplotlib.colors.BoundaryNorm(bounds, cmap.N)
    cmap.set_bad("lightgrey")   # truly empty / no-data wells

    runs = df["run_id"].unique()
    for run in runs:
        sub = df[df["run_id"] == run].copy()
        grid       = np.full((8, 12), np.nan)
        label_grid = np.full((8, 12), "", dtype=object)

        for _, row in sub.iterrows():
            r, c = _well_to_rowcol(row["well"])
            if r is None:
                continue
            val = result_map.get(str(row["result"]).strip().lower(), np.nan)
            grid[r, c]       = val
            label_grid[r, c] = _wrap_well_text(_well_label(row))

        fig, ax = plt.subplots(figsize=(14, 7))
        im = ax.imshow(grid, cmap=cmap, norm=norm, aspect="auto")

        # Sample ID annotations
        for r in range(8):
            for c in range(12):
                lbl = label_grid[r, c]
                if lbl:
                    ax.text(c, r, lbl, ha="center", va="center",
                            fontsize=5, color="black", linespacing=1.3,
                            multialignment="center")

        # Grid lines
        ax.set_xticks(np.arange(-0.5, 12, 1), minor=True)
        ax.set_yticks(np.arange(-0.5, 8, 1), minor=True)
        ax.grid(which="minor", color="white", linewidth=1.5)

        ax.set_xticks(range(12))
        ax.set_xticklabels([str(i + 1).zfill(2) for i in range(12)])
        ax.set_yticks(range(8))
        ax.set_yticklabels(list("ABCDEFGH"))
        ax.set_title(f"Plate Result Map  |  Run: {run}", fontsize=12, fontweight="bold")

        patches = [mpatches.Patch(color=c, label=lbl)
                   for lbl, c in result_colors.items()]
        patches.append(mpatches.Patch(color="lightgrey", label="Empty"))
        ax.legend(handles=patches, bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=9)

        fname = f"05_plate_heatmap_result_{run}.png"
        _savefig(fig, out_dir / fname)


def fig_plate_heatmap_cq(df: pd.DataFrame, out_dir: Path):
    runs = df["run_id"].unique()
    for run in runs:
        sub_all = df[df["run_id"] == run].copy()
        sub     = sub_all[sub_all["target_cq"].notna()].copy()
        if sub.empty:
            continue

        grid        = np.full((8, 12), np.nan)
        label_grid  = np.full((8, 12), "", dtype=object)
        result_grid = np.full((8, 12), False, dtype=bool)

        # Populate Cq grid from wells that have a Cq value
        for _, row in sub.iterrows():
            r, c = _well_to_rowcol(row["well"])
            if r is not None:
                grid[r, c]       = row["target_cq"]
                label_grid[r, c] = _well_label(row)

        # Track which wells were called Positive (all wells, not just those with Cq)
        for _, row in sub_all.iterrows():
            r, c = _well_to_rowcol(row["well"])
            if r is not None and str(row.get("result", "")).strip().lower() == "positive":
                result_grid[r, c] = True

        fig, ax = plt.subplots(figsize=(14, 7))
        cmap = plt.cm.RdYlGn.copy()
        cmap.set_bad("lightgrey")
        vmin = np.nanmin(grid) - 1
        vmax = np.nanmax(grid) + 1

        im = ax.imshow(grid, cmap=cmap, aspect="auto", vmin=vmin, vmax=vmax)
        cbar = fig.colorbar(im, ax=ax, fraction=0.03, pad=0.04)
        cbar.set_label("Cq", fontsize=9)

        # Annotate cells: sample ID (small, top) + Cq value (larger, bottom)
        for r in range(8):
            for c in range(12):
                val = grid[r, c]
                if not np.isnan(val):
                    lbl = _wrap_well_text(label_grid[r, c], width=10)
                    if lbl:
                        ax.text(c, r - 0.18, lbl, ha="center", va="center",
                                fontsize=4.5, color="black", linespacing=1.2,
                                multialignment="center")
                    ax.text(c, r + 0.18, f"{val:.1f}", ha="center", va="center",
                            fontsize=6, color="black", fontweight="bold")
                    if result_grid[r, c]:
                        ax.text(c, r + 0.38, "POS", ha="center", va="center",
                                fontsize=5.5, color="black", fontweight="bold")

        # Border overlay for Positive-called wells
        for r in range(8):
            for c in range(12):
                if result_grid[r, c]:
                    ax.add_patch(mpatches.Rectangle(
                        (c - 0.5, r - 0.5), 1, 1,
                        linewidth=2.5, edgecolor="black", facecolor="none", zorder=4,
                    ))

        ax.set_xticks(np.arange(-0.5, 12, 1), minor=True)
        ax.set_yticks(np.arange(-0.5, 8, 1), minor=True)
        ax.grid(which="minor", color="white", linewidth=1.5)
        ax.set_xticks(range(12))
        ax.set_xticklabels([str(i + 1).zfill(2) for i in range(12)])
        ax.set_yticks(range(8))
        ax.set_yticklabels(list("ABCDEFGH"))
        ax.set_title(f"Plate Cq Heatmap  |  Run: {run}", fontsize=12, fontweight="bold")

        pos_patch = mpatches.Patch(facecolor="none", edgecolor="black", linewidth=2.5,
                                   label="Positive result")
        ax.legend(handles=[pos_patch], bbox_to_anchor=(1.02, 0), loc="lower left", fontsize=9)

        fname = f"06_plate_heatmap_cq_{run}.png"
        _savefig(fig, out_dir / fname)


# ---------------------------------------------------------------------------
# Figure 7: Multi-run positivity heatmap
# ---------------------------------------------------------------------------
def fig_run_heatmap(df: pd.DataFrame, out_dir: Path):
    if df["run_id"].nunique() < 2:
        print("  [viz] Skipping run heatmap – only one run present.")
        return

    unknowns = df[~df["is_control"]].copy()
    unknowns["is_positive"] = unknowns["result"].str.strip().str.lower() == "positive"

    pivot = unknowns.pivot_table(
        index="run_id", columns="sample_type",
        values="is_positive", aggfunc="mean"
    ) * 100

    fig, ax = plt.subplots(figsize=(max(8, pivot.shape[1] * 1.2),
                                    max(4, pivot.shape[0] * 0.7)))
    im = ax.imshow(pivot.values, cmap="RdYlGn", aspect="auto", vmin=0, vmax=100)
    cbar = fig.colorbar(im, ax=ax, fraction=0.03, pad=0.04)
    cbar.set_label("Positivity Rate (%)", fontsize=9)

    for r in range(pivot.shape[0]):
        for c in range(pivot.shape[1]):
            val = pivot.values[r, c]
            if not np.isnan(val):
                ax.text(c, r, f"{val:.0f}%", ha="center", va="center",
                        fontsize=8, color="black")

    ax.set_xticks(range(pivot.shape[1]))
    ax.set_xticklabels(pivot.columns.tolist(), rotation=30, ha="right", fontsize=9)
    ax.set_yticks(range(pivot.shape[0]))
    ax.set_yticklabels(pivot.index.tolist(), fontsize=9)
    ax.set_title("Positivity Rate (%) — Run × Sample Type", fontsize=13, fontweight="bold")

    _savefig(fig, out_dir / "07_run_positivity_heatmap.png")


# ---------------------------------------------------------------------------
# Figure 8: I.C. Extraction QC Overview (2-panel)
# ---------------------------------------------------------------------------
def _well_sort_key(well: str) -> tuple:
    """Sort key: ('A', 1), ('A', 2), … ('H', 12) — row-major plate order."""
    m = re.match(r"([A-Ha-h])(\d+)", str(well).strip())
    if not m:
        return ("Z", 999)
    return (m.group(1).upper(), int(m.group(2)))


def fig_ic_control_chart(df: pd.DataFrame, out_dir: Path):
    """
    2-panel I.C. Extraction QC Overview.

    Panel A — Cq by plate row: spatial view of Cq across rows A–H, column
    position used as x-spread, flagged (Borderline) wells annotated.

    Panel B — Cq by sample type: stripplot + boxplot, colored by run_id,
    with grand mean ±2 SD reference.
    """
    data_df = df[
        (~df["is_control"]) &
        df["target_cq"].notna()
    ].copy()

    if data_df.empty:
        print("  [viz] Skipping Cq spatial chart – no data.")
        return

    runs         = sorted(data_df["run_id"].unique())
    sample_types = sorted(data_df["sample_type"].unique())
    n_runs       = len(runs)

    # Color maps
    st_color  = {st: SAMPLE_PALETTE[i % len(SAMPLE_PALETTE)]
                 for i, st in enumerate(sample_types)}
    run_color = {rid: SAMPLE_PALETTE[i % len(SAMPLE_PALETTE)]
                 for i, rid in enumerate(runs)}

    # Grand stats
    grand_mean = data_df["target_cq"].mean()
    grand_sd   = data_df["target_cq"].std()

    fig, (ax_lj, ax_strip) = plt.subplots(
        2, 1, figsize=(max(10, n_runs * 3), 10)
    )

    # ------------------------------------------------------------------
    # Panel A — Levey-Jennings
    # ------------------------------------------------------------------
    plate_rows  = list("ABCDEFGH")
    row_idx_map = {r: i for i, r in enumerate(plate_rows)}

    def _well_row(well):
        m = re.match(r"([A-Ha-h])(\d+)", str(well).strip())
        return m.group(1).upper() if m else None

    def _well_col(well):
        m = re.match(r"([A-Ha-h])(\d+)", str(well).strip())
        return int(m.group(2)) if m else None

    data_df["_row"] = data_df["well"].apply(_well_row)
    data_df["_col"] = data_df["well"].apply(_well_col)

    # Check whether qc_flag column is present
    has_flags = "qc_flag" in data_df.columns

    # ±2 SD reference band
    ax_lj.fill_between(
        [-0.5, len(plate_rows) - 0.5],
        grand_mean - 2 * grand_sd,
        grand_mean + 2 * grand_sd,
        color="grey", alpha=0.10, zorder=1
    )
    ax_lj.hlines([grand_mean - 2 * grand_sd, grand_mean + 2 * grand_sd],
                 -0.5, len(plate_rows) - 0.5,
                 colors="grey", lw=0.7, linestyles=":")

    # Plot each well: x = row group + column-based spread, color = run_id
    for _, row in data_df.iterrows():
        xi = row_idx_map.get(row["_row"])
        if xi is None or pd.isna(row["_col"]):
            continue
        # Columns 1-12 spread across ±0.38 within each row group
        x_pos = xi + (row["_col"] - 6.5) / 12.0 * 0.76
        ax_lj.scatter(x_pos, row["target_cq"],
                      color=run_color[row["run_id"]], s=22, zorder=3, alpha=0.80)

    # Flagged wells: red ring + well label
    if has_flags:
        flagged = data_df[data_df["qc_flag"].str.contains(
            "Borderline", case=False, na=False
        )]
        for _, row in flagged.iterrows():
            xi = row_idx_map.get(row["_row"])
            if xi is None or pd.isna(row["_col"]):
                continue
            x_pos = xi + (row["_col"] - 6.5) / 12.0 * 0.76
            ax_lj.scatter(x_pos, row["target_cq"], s=80,
                          facecolors="none", edgecolors="red",
                          linewidths=1.5, zorder=4)
            ax_lj.annotate(
                row["well"],
                (x_pos, row["target_cq"]),
                textcoords="offset points", xytext=(4, 4),
                fontsize=6.5, color="red"
            )

    ax_lj.set_xticks(range(len(plate_rows)))
    ax_lj.set_xticklabels(plate_rows, fontsize=10)
    ax_lj.set_xlim(-0.6, len(plate_rows) - 0.4)
    ax_lj.set_ylabel("Cq", fontsize=11)
    ax_lj.set_xlabel("Plate Row  (left→right = column 1→12)", fontsize=10)
    ax_lj.set_title(
        "Panel A — Cq by Plate Row",
        fontsize=12, fontweight="bold"
    )
    ax_lj.spines[["top", "right"]].set_visible(False)

    run_patches_lj = [mpatches.Patch(color=run_color[rid], label=rid)
                      for rid in runs]
    if has_flags:
        flag_marker = plt.Line2D(
            [0], [0], marker="o", color="w", markerfacecolor="none",
            markeredgecolor="red", markersize=8, label="Borderline"
        )
        run_patches_lj.append(flag_marker)
    ax_lj.legend(handles=run_patches_lj, fontsize=8, loc="upper right",
                 title="Run" if n_runs > 1 else None, title_fontsize=8)

    # ------------------------------------------------------------------
    # Panel B — Cq by sample type
    # ------------------------------------------------------------------
    # Grand reference band
    ax_strip.fill_between(
        [-0.5, len(sample_types) - 0.5],
        grand_mean - 2 * grand_sd,
        grand_mean + 2 * grand_sd,
        color="grey", alpha=0.10, zorder=1
    )
    ax_strip.axhline(grand_mean + 2 * grand_sd, color="grey",
                     lw=0.7, linestyle=":", zorder=1)
    ax_strip.axhline(grand_mean - 2 * grand_sd, color="grey",
                     lw=0.7, linestyle=":", zorder=1)

    rng = np.random.default_rng(seed=42)

    for st_idx, st in enumerate(sample_types):
        grp = data_df[data_df["sample_type"] == st]["target_cq"].dropna().values

        # Boxplot (drawn first, so points sit on top)
        if len(grp) >= 2:
            ax_strip.boxplot(
                grp,
                positions=[st_idx],
                widths=0.35,
                patch_artist=False,
                manage_ticks=False,
                flierprops=dict(marker=""),
                medianprops=dict(color="dimgrey", lw=1.5),
                whiskerprops=dict(color="grey", lw=0.8),
                capprops=dict(color="grey", lw=0.8),
                boxprops=dict(color="grey", lw=0.8),
                zorder=2,
            )

        # Stripplot (jittered)
        grp_df = data_df[data_df["sample_type"] == st].copy()
        jitter = rng.uniform(-0.12, 0.12, size=len(grp_df))
        for j, (_, row) in enumerate(grp_df.iterrows()):
            colour = run_color[row["run_id"]]
            ax_strip.scatter(st_idx + jitter[j], row["target_cq"],
                             color=colour, s=18, alpha=0.75, zorder=3)

        # n= annotation below the lowest data point
        ax_strip.text(st_idx, data_df["target_cq"].min() - grand_sd * 0.5,
                      f"n={len(grp)}", ha="center", va="top",
                      fontsize=7.5, color="dimgrey")

    ax_strip.set_xticks(range(len(sample_types)))
    ax_strip.set_xticklabels(
        [textwrap.fill(st, 12) for st in sample_types],
        fontsize=9
    )
    ax_strip.set_xlim(-0.6, len(sample_types) - 0.4)
    ax_strip.set_ylabel("Cq", fontsize=11)
    ax_strip.set_xlabel("Sample Type", fontsize=11)
    ax_strip.set_title(
        "Panel B — Cq by Sample Type",
        fontsize=12, fontweight="bold"
    )
    ax_strip.spines[["top", "right"]].set_visible(False)

    if n_runs > 1:
        run_patches = [mpatches.Patch(color=run_color[rid], label=rid)
                       for rid in runs]
        ax_strip.legend(handles=run_patches, fontsize=8, loc="upper right",
                        title="Run", title_fontsize=8)

    fig.suptitle(
        "Cq Spatial & Sample-Type Overview",
        fontsize=13, fontweight="bold", y=1.01
    )

    _savefig(fig, out_dir / "08_ic_cq_control_chart.png")


# ---------------------------------------------------------------------------
# Figure 9: Inter-run CV% bar chart
# ---------------------------------------------------------------------------
def fig_interrun_cv_bar(cv_df: pd.DataFrame, out_dir: Path):
    """
    Bar chart of inter-run CV% per sample type.
    Reference line at 20% (common acceptability threshold).
    """
    if cv_df.empty:
        print("  [viz] Skipping inter-run CV% chart – no data.")
        return

    sample_types = cv_df["sample_type"].tolist()
    cv_vals      = cv_df["cv_pct"].tolist()
    colours      = [SAMPLE_PALETTE[i % len(SAMPLE_PALETTE)] for i in range(len(sample_types))]

    fig, ax = plt.subplots(figsize=(max(6, len(sample_types) * 1.3), 5))
    bars = ax.bar(sample_types, cv_vals, color=colours, edgecolor="white", linewidth=0.8)

    for bar, val in zip(bars, cv_vals):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.5,
            f"{val:.1f}%",
            ha="center", va="bottom", fontsize=8.5,
        )

    # Reference line at 20%
    ax.axhline(20, color="#e74c3c", lw=1.2, linestyle="--", alpha=0.8,
               label="20% acceptability threshold")

    ax.set_ylabel("Inter-run CV (%)", fontsize=11)
    ax.set_xlabel("Sample Type", fontsize=11)
    ax.set_title("Inter-run CV% of Target Cq by Sample Type", fontsize=13, fontweight="bold")
    ax.set_xticklabels(sample_types, rotation=30, ha="right", fontsize=9)
    ax.legend(fontsize=8)
    ax.spines[["top", "right"]].set_visible(False)

    _savefig(fig, out_dir / "09_interrun_cv_bar.png")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Generate qPCR visualizations.")
    parser.add_argument("--input_dir",  required=True)
    parser.add_argument("--stats",      required=True)
    parser.add_argument("--output_dir", required=True)
    args = parser.parse_args()

    proc_dir   = Path(args.input_dir)
    out_dir    = Path(args.output_dir)
    report_dir = Path(args.stats).parent   # derive reports/ from --stats path
    out_dir.mkdir(parents=True, exist_ok=True)

    # Load combined data
    files = sorted(proc_dir.glob("*_parsed.csv"))
    df    = pd.concat([pd.read_csv(f) for f in files], ignore_index=True)
    stats = pd.read_csv(args.stats)

    # Generate all figures
    fig_positivity_bar(stats, out_dir)
    fig_cq_scatter(df, out_dir)
    fig_plate_heatmap_result(df, out_dir)
    fig_plate_heatmap_cq(df, out_dir)
    fig_run_heatmap(df, out_dir)
    fig_ic_control_chart(df, out_dir)

    cv_path = report_dir / "interrun_cv.csv"
    if cv_path.exists():
        cv_df = pd.read_csv(cv_path)
        if not cv_df.empty:
            fig_interrun_cv_bar(cv_df, out_dir)
        else:
            print("  [viz] Skipping inter-run CV% chart – empty data file.")
    else:
        print("  [viz] Skipping inter-run CV% chart – interrun_cv.csv not found (< 2 runs).")

    print(f"  [viz] All figures saved to {out_dir}")


if __name__ == "__main__":
    main()
