"""
analyze.py  –  Step 2 of the qPCR pipeline
Aggregates all parsed CSVs, computes per-sample-type statistics,
flags QC issues, and writes:
  all_runs_combined.csv      – merged, annotated dataset (includes qc_flag column)
  summary_statistics.csv     – mean / SD / n / positivity rate per sample type
  control_statistics.csv     – plate control summary
  run_comparison.csv         – per-run × per-sample-type positivity pivot
  qc_flags.csv               – all flagged wells with run_id, well, sample_type, flag_reason
  interrun_cv.csv            – inter-run CV% per sample type (only when 2+ runs present)
"""

import argparse
from pathlib import Path

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_all_parsed(proc_dir: Path) -> pd.DataFrame:
    files = sorted(proc_dir.glob("*_parsed.csv"))
    if not files:
        raise FileNotFoundError(f"No *_parsed.csv files found in {proc_dir}")
    frames = [pd.read_csv(f) for f in files]
    df = pd.concat(frames, ignore_index=True)
    print(f"  [analyze] Loaded {len(files)} run(s), {len(df)} total wells.")
    return df


def compute_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    Per-sample-type summary statistics, across all runs.
    Excludes plate controls (Pos Ctrl / Neg Ctrl).
    """
    unknowns = df[~df["is_control"]].copy()

    # Positivity: 'Positive' result (case-insensitive)
    unknowns["is_positive"] = (
        unknowns["result"].str.strip().str.lower() == "positive"
    )
    unknowns["has_target_cq"] = unknowns["target_cq"].notna()

    groups = unknowns.groupby("sample_type")

    stats = groups.agg(
        n_wells            = ("well", "count"),
        n_positive         = ("is_positive", "sum"),
        n_negative         = ("is_positive", lambda x: (~x).sum()),
        n_target_cq        = ("has_target_cq", "sum"),
        mean_target_cq     = ("target_cq", "mean"),
        sd_target_cq       = ("target_cq", "std"),
        min_target_cq      = ("target_cq", "min"),
        max_target_cq      = ("target_cq", "max"),
        mean_ic_cq         = ("ic_cq", "mean"),
        sd_ic_cq           = ("ic_cq", "std"),
    ).reset_index()

    stats["positivity_rate_pct"] = (
        stats["n_positive"] / stats["n_wells"] * 100
    ).round(1)

    # Round Cq stats
    for col in ["mean_target_cq", "sd_target_cq", "min_target_cq",
                "max_target_cq", "mean_ic_cq", "sd_ic_cq"]:
        stats[col] = stats[col].round(3)

    # Sort by sample type for readability
    stats.sort_values("sample_type", inplace=True)
    stats.reset_index(drop=True, inplace=True)
    return stats


def compute_control_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Summary of plate controls across all runs."""
    controls = df[df["is_control"]].copy()
    if controls.empty:
        return pd.DataFrame()

    controls["ctrl_type"] = controls["content"].str.strip()
    summary = controls.groupby("ctrl_type").agg(
        n_controls     = ("well", "count"),
        n_valid        = ("result", lambda x: (x.str.strip().str.lower() == "valid ctrl").sum()),
        mean_target_cq = ("target_cq", "mean"),
        sd_target_cq   = ("target_cq", "std"),
        mean_ic_cq     = ("ic_cq", "mean"),
        sd_ic_cq       = ("ic_cq", "std"),
    ).reset_index()

    for col in ["mean_target_cq", "sd_target_cq", "mean_ic_cq", "sd_ic_cq"]:
        summary[col] = summary[col].round(3)

    return summary


def compare_runs(df: pd.DataFrame) -> pd.DataFrame:
    """
    Per-run × per-sample-type positivity table:
    Useful for spotting inter-run variation.
    """
    unknowns = df[~df["is_control"]].copy()
    unknowns["is_positive"] = (
        unknowns["result"].str.strip().str.lower() == "positive"
    )

    pivot = unknowns.pivot_table(
        index="run_id",
        columns="sample_type",
        values="is_positive",
        aggfunc=["sum", "count"],
    )
    pivot.columns = ["_".join(c).strip() for c in pivot.columns]
    pivot.reset_index(inplace=True)
    return pivot


# ---------------------------------------------------------------------------
# QC flagging
# ---------------------------------------------------------------------------

def flag_high_cq(df: pd.DataFrame) -> pd.DataFrame:
    """Flag non-control wells where Target Cq > 35 as Borderline."""
    mask = (~df["is_control"]) & (df["target_cq"] > 35) & df["target_cq"].notna()
    flagged = df.loc[mask, ["run_id", "well", "sample_type"]].copy()
    flagged["flag_reason"] = "Borderline: Target Cq > 35"
    return flagged.reset_index(drop=True)


def flag_ic_extraction(df: pd.DataFrame) -> pd.DataFrame:
    """
    Flag non-control wells where I.C. Cq is >2 SD from the plate (run) mean.
    Skips runs with fewer than 3 valid IC Cq values.
    """
    flags = []
    unknowns = df[~df["is_control"]].copy()

    for run_id, group in unknowns.groupby("run_id"):
        valid_ic = group["ic_cq"].dropna()
        if len(valid_ic) < 3:
            continue
        mean_ic = valid_ic.mean()
        sd_ic   = valid_ic.std()
        if pd.isna(sd_ic) or sd_ic == 0:
            continue

        outliers = group[
            group["ic_cq"].notna() &
            (abs(group["ic_cq"] - mean_ic) > 2 * sd_ic)
        ]
        for _, row in outliers.iterrows():
            flags.append({
                "run_id":      run_id,
                "well":        row["well"],
                "sample_type": row["sample_type"],
                "flag_reason": (
                    f"IC extraction failure: I.C. Cq={row['ic_cq']:.2f} "
                    f"(plate mean={mean_ic:.2f}, SD={sd_ic:.2f})"
                ),
            })

    return pd.DataFrame(flags) if flags else pd.DataFrame(
        columns=["run_id", "well", "sample_type", "flag_reason"]
    )


def compute_interrun_cv(df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute inter-run CV% for Target Cq per sample type.
    Returns empty DataFrame when fewer than 2 runs are present.
    Only includes sample types that appear in 2+ runs.
    """
    if df["run_id"].nunique() < 2:
        return pd.DataFrame()

    unknowns = df[~df["is_control"]].copy()
    # Per-run mean Cq per sample type
    per_run = (
        unknowns.groupby(["run_id", "sample_type"])["target_cq"]
        .mean()
        .reset_index()
        .rename(columns={"target_cq": "run_mean_cq"})
    )

    # Across runs: CV%
    cv = per_run.groupby("sample_type")["run_mean_cq"].agg(
        n_runs="count",
        mean_cq="mean",
        sd_cq="std",
    ).reset_index()

    cv = cv[cv["n_runs"] >= 2].copy()
    cv["cv_pct"]  = (cv["sd_cq"] / cv["mean_cq"] * 100).round(1)
    cv["mean_cq"] = cv["mean_cq"].round(3)
    cv["sd_cq"]   = cv["sd_cq"].round(3)
    return cv.reset_index(drop=True)


def _annotate_qc_flags(df: pd.DataFrame, qc_flags: pd.DataFrame) -> pd.DataFrame:
    """Add a qc_flag column to the combined dataframe."""
    df = df.copy()
    df["qc_flag"] = ""

    if qc_flags.empty:
        return df

    for _, flag_row in qc_flags.iterrows():
        mask = (df["run_id"] == flag_row["run_id"]) & (df["well"] == flag_row["well"])
        existing = df.loc[mask, "qc_flag"]
        df.loc[mask, "qc_flag"] = existing.apply(
            lambda v: (v + "; " + flag_row["flag_reason"]) if v else flag_row["flag_reason"]
        )
    return df


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Analyze parsed qPCR CSVs.")
    parser.add_argument("--input_dir",  required=True)
    parser.add_argument("--output_dir", required=True)
    args = parser.parse_args()

    proc_dir = Path(args.input_dir)
    out_dir  = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = load_all_parsed(proc_dir)

    # --- QC flagging ---
    flags_high_cq = flag_high_cq(df)
    flags_ic      = flag_ic_extraction(df)
    qc_flags      = pd.concat([flags_high_cq, flags_ic], ignore_index=True)

    qc_path = out_dir / "qc_flags.csv"
    qc_flags.to_csv(qc_path, index=False)
    n_flags = len(qc_flags)
    print(f"  [analyze] QC flags: {n_flags} well(s) flagged → {qc_path}")
    if n_flags:
        for reason, grp in qc_flags.groupby("flag_reason", sort=False):
            print(f"            {len(grp):3d}× {reason[:80]}")

    # Annotate combined dataset with qc_flag column
    df = _annotate_qc_flags(df, qc_flags)

    # Combined dataset
    combined_path = out_dir / "all_runs_combined.csv"
    df.to_csv(combined_path, index=False)
    print(f"  [analyze] Combined dataset → {combined_path}")

    # Per-sample-type stats
    stats = compute_stats(df)
    stats_path = out_dir / "summary_statistics.csv"
    stats.to_csv(stats_path, index=False)
    print(f"  [analyze] Summary statistics → {stats_path}")
    print(stats[["sample_type", "n_wells", "n_positive", "positivity_rate_pct",
                 "mean_target_cq", "sd_target_cq"]].to_string(index=False))

    # Control stats
    ctrl_stats = compute_control_stats(df)
    if not ctrl_stats.empty:
        ctrl_path = out_dir / "control_statistics.csv"
        ctrl_stats.to_csv(ctrl_path, index=False)
        print(f"  [analyze] Control statistics → {ctrl_path}")

    # Inter-run comparison
    run_comparison = compare_runs(df)
    run_path = out_dir / "run_comparison.csv"
    run_comparison.to_csv(run_path, index=False)
    print(f"  [analyze] Run comparison → {run_path}")

    # Inter-run CV%
    cv = compute_interrun_cv(df)
    if not cv.empty:
        cv_path = out_dir / "interrun_cv.csv"
        cv.to_csv(cv_path, index=False)
        print(f"  [analyze] Inter-run CV% → {cv_path}")
        print(cv[["sample_type", "n_runs", "mean_cq", "cv_pct"]].to_string(index=False))
    else:
        print("  [analyze] Inter-run CV%: skipped (< 2 runs).")


if __name__ == "__main__":
    main()
