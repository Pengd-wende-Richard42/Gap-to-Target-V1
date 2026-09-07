"""Publication figures for the Paris Gap-to-Target analysis.

This script replaces the Paris-related graphical routines from PolicyBrief.py.
It is deliberately limited to the revised Paris framework:

* Gap-to-Target indicators (NDC1 and NDC2);
* target-consistent sectoral emissions;
* country-level medians (not means and no winsorisation);
* uncertainty across five alternative scenarios on a common sample;
* genuinely paired conditional/unconditional targets;
* target type, sectoral coverage, income and region heterogeneity;
* ambition versus pathway alignment;
* within-country transitions in ambition and alignment across NDC cycles;
* appendix distributions, target architecture and world maps.

The obsolete implementation/adjustment-pace indicators are never used.

Run on Windows (default project folders):
    python Paris_GapToTarget_Figures_Final.py

Run with explicit files (useful for testing):
    python Paris_GapToTarget_Figures_Final.py \
      --final-data ".../Paris_GapToTarget_Final.dta" \
      --targets-data ".../Paris_NDC_Targets.dta" \
      --icr-data ".../ICR.xlsx" \
      --output-dir ".../Figures"
"""

from __future__ import annotations

import argparse
import json
import math
import sys
import warnings
import re
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.colors import TwoSlopeNorm
from matplotlib.lines import Line2D
from matplotlib.patches import Patch


# -----------------------------------------------------------------------------
# Paths and constants
# -----------------------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parents[2]
RAW_DATA_DIR = PROJECT_ROOT / "data" / "raw" / "paris"
PROCESSED_DATA_DIR = PROJECT_ROOT / "data" / "processed" / "paris"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs" / "figures" / "paris"
DEFAULT_WORLD_SHAPEFILE = (
    PROJECT_ROOT / "data" / "raw" / "geospatial" /
    "ne_110m_admin_0_countries.zip"
)

DEFAULT_FINAL_DATA = PROCESSED_DATA_DIR / "Paris_GapToTarget_Final.dta"
DEFAULT_TARGETS_DATA = PROCESSED_DATA_DIR / "Paris_NDC_Targets.dta"
DEFAULT_ICR_DATA = RAW_DATA_DIR / "ICR.xlsx"

NATURAL_EARTH_URL = (
    "https://naturalearth.s3.amazonaws.com/110m_cultural/"
    "ne_110m_admin_0_countries.zip"
)

CYCLES = ("n1", "n2")
CYCLE_LABEL = {"n1": "First NDC", "n2": "Second NDC"}
CYCLE_SHORT = {"n1": "NDC1", "n2": "NDC2"}
CYCLE_COLOR = {"n1": "#2878B5", "n2": "#D95F02"}

SCENARIOS = ("rcp45", "rcp85", "oecd", "pik", "iiasa")
SCENARIO_LABEL = {
    "main": "Main estimate",
    "rcp45": "RCP4.5",
    "rcp85": "RCP8.5",
    "oecd": "OECD",
    "pik": "PIK",
    "iiasa": "IIASA",
}

TARGET_TYPE_ORDER = (
    "BASE_YEAR_REDUCTION",
    "BAU_REDUCTION",
    "INTENSITY_TARGET",
    "FIXED_LEVEL_TARGET",
    "PER_CAPITA_COMBINED",
    "TRAJECTORY_TARGET",
    "OTHER_OR_UNKNOWN",
)
TARGET_TYPE_LABEL = {
    "BASE_YEAR_REDUCTION": "Base year",
    "BAU_REDUCTION": "BAU",
    "INTENSITY_TARGET": "Intensity",
    "FIXED_LEVEL_TARGET": "Fixed level",
    "PER_CAPITA_COMBINED": "Per-capita target",
    "TRAJECTORY_TARGET": "Cumulative trajectory",
    "OTHER_OR_UNKNOWN": "No usable quantified target",
}

COVERAGE_ORDER = ("ECONOMY_WIDE_OR_BROAD", "MULTISECTORAL", "SECTORAL")
COVERAGE_LABEL = {
    "ECONOMY_WIDE_OR_BROAD": "Economy-wide or broad",
    "MULTISECTORAL": "Multisectoral",
    "SECTORAL": "Sectoral",
}

INCOME_ORDER = (
    "Low income",
    "Lower middle income",
    "Upper middle income",
    "High income",
)

REGION_ORDER = (
    "East Asia & Pacific",
    "Europe & Central Asia",
    "Latin America & Caribbean",
    "Middle East & North Africa",
    "North America",
    "South Asia",
    "Sub-Saharan Africa",
)

# Additional geographic breakdown requested for the regional robustness figure.
# The original seven-region classification is retained unchanged in its own
# Appendix figure. This second classification separates Europe from Central
# Asia and the Middle East from North Africa. A small number of countries are
# reassigned from broad source groups to their geographic subregion.
DETAILED_REGION_ORDER = (
    "East Asia & Pacific",
    "Europe",
    "Central Asia",
    "Latin America & Caribbean",
    "Middle East",
    "North Africa",
    "North America",
    "South Asia",
    "Sub-Saharan Africa",
)

CENTRAL_ASIA_ISO3 = {"KAZ", "KGZ", "TJK", "TKM", "UZB"}
NORTH_AFRICA_ISO3 = {"DZA", "EGY", "LBY", "MAR", "TUN"}
WESTERN_ASIA_ISO3 = {"ARM", "AZE", "CYP", "GEO", "TUR"}
SOUTH_ASIA_OVERRIDES = {"AFG", "PAK"}
SUB_SAHARAN_AFRICA_OVERRIDES = {"DJI"}
EUROPE_OVERRIDES = {"MLT"}

CONDITIONALITY_ORDER = (
    "UNCONDITIONAL_ONLY",
    "CONDITIONAL_ONLY",
    "MIXED_CONDITIONALITY",
    "PARTIALLY_CONDITIONAL",
    "UNKNOWN",
)
CONDITIONALITY_LABEL = {
    "UNCONDITIONAL_ONLY": "Unconditional only",
    "CONDITIONAL_ONLY": "Conditional only",
    "MIXED_CONDITIONALITY": "Conditional and unconditional",
    "PARTIALLY_CONDITIONAL": "Partially conditional",
    "UNKNOWN": "Unknown",
}

CONDITIONALITY_GRAPH_ORDER = (
    "UNCONDITIONAL_ONLY",
    "CONDITIONAL_ONLY",
    "BOTH",
)
CONDITIONALITY_GRAPH_LABEL = {
    "UNCONDITIONAL_ONLY": "Unconditional only",
    "CONDITIONAL_ONLY": "Conditional only",
    "BOTH": "Both",
}

RANDOM_SEED = 20260824

# Slightly thicker interval strokes; cap widths remain unchanged.
INTERVAL_LINEWIDTH = 1.8
INTERVAL_CAPTHICK = 1.5
# Only BAU-referenced targets are plotted in the country sensitivity figures.
# Full reference classifications remain available in the diagnostic tables.
REFERENCE_LABEL = {"BAU": "BAU reference"}
SCENARIO_MARKERS = dict(zip(SCENARIOS, ("s", "^", "v", "D", "P")))
SCENARIO_COLORS = dict(zip(SCENARIOS, ("#6A51A3", "#A6761D", "#238B45", "#C43C39", "#637A89")))

PAIR_AUDIT_COLUMNS = (
    "pair_first_second_available",
    "pair_target_comparable_flag",
    "target_change_first_to_second_pp",
    "ambition_change_direction",
    "metric_changed_first_to_second",
    "reference_changed_first_to_secon",
    "scope_changed_first_to_second",
    "target_year_changed_first_to_sec",
)


@dataclass(frozen=True)
class OutputFolders:
    root: Path
    main: Path
    appendix: Path
    diagnostics: Path


# -----------------------------------------------------------------------------
# General utilities
# -----------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate revised Paris Gap-to-Target publication figures."
    )
    parser.add_argument("--final-data", type=Path, default=DEFAULT_FINAL_DATA)
    parser.add_argument("--targets-data", type=Path, default=DEFAULT_TARGETS_DATA)
    parser.add_argument("--icr-data", type=Path, default=DEFAULT_ICR_DATA)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument(
        "--no-maps", action="store_true", help="Skip optional appendix maps."
    )
    parser.add_argument(
        "--world-shapefile",
        type=Path,
        default=DEFAULT_WORLD_SHAPEFILE,
        help=(
            "Optional Natural Earth country-boundary .shp or .zip file. When omitted, "
            "the official 110m Admin 0 countries archive is downloaded and cached."
        ),
    )
    parser.add_argument(
        "--show", action="store_true", help="Display figures while running."
    )
    parser.add_argument(
        "--country-plot-style", choices=("range", "dumbbell", "both"), default="range",
        help="BAU country plots only: min-max range, farthest-scenario dumbbell, or both.",
    )
    parser.add_argument(
        "--countries-per-page", type=int, default=32,
        help="Maximum countries per image (all countries are retained across pages).",
    )
    return parser.parse_args()


def configure_style() -> None:
    # Matplotlib-only styling keeps the script independent of seaborn/SciPy.
    # This avoids binary compatibility failures when a recent NumPy version is
    # installed alongside an older compiled SciPy wheel.
    plt.style.use("seaborn-v0_8-whitegrid")
    mpl.rcParams.update(
        {
            "figure.dpi": 110,
            "savefig.dpi": 400,
            "font.family": "DejaVu Sans",
            "font.size": 9.5,
            "axes.titlesize": 11,
            "axes.labelsize": 10,
            "axes.titleweight": "semibold",
            "legend.fontsize": 8.5,
            "xtick.labelsize": 8.5,
            "ytick.labelsize": 8.5,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "axes.grid": True,
            "grid.alpha": 0.20,
            "grid.linewidth": 0.65,
            "pdf.fonttype": 42,
            "ps.fonttype": 42,
        }
    )


def create_output_folders(root: Path) -> OutputFolders:
    folders = OutputFolders(
        root=root,
        main=root / "Main",
        appendix=root / "Appendix",
        diagnostics=root / "Diagnostics",
    )
    for folder in (folders.root, folders.main, folders.appendix, folders.diagnostics):
        folder.mkdir(parents=True, exist_ok=True)
    # Remove only obsolete outputs previously produced by this script. This
    # prevents a stale main-text figure from being mistaken for a current one.
    for stem in ("Figure_8_Ambition_and_Alignment", "Figure_6_Paired_Conditionality"):
        for extension in ("png", "pdf"):
            obsolete = folders.main / f"{stem}.{extension}"
            if obsolete.exists():
                obsolete.unlink()
    return folders


def save_figure(
    fig: mpl.figure.Figure,
    stem: str,
    folder: Path,
    show: bool,
    *,
    formats: Sequence[str] = ("png",),
) -> None:
    fig.tight_layout()
    for extension in formats:
        # Atomic replacement prevents a partially written image from being left
        # behind if a renderer or synchronized folder interrupts a write.
        output = folder / f"{stem}.{extension}"
        temporary = folder / f".{stem}.{extension}.tmp"
        fig.savefig(temporary, format=extension, bbox_inches="tight")
        temporary.replace(output)
    if show:
        plt.show()
    plt.close(fig)


def clean_iso(series: pd.Series) -> pd.Series:
    return series.astype("string").str.strip().str.upper()


def require_columns(data: pd.DataFrame, columns: Iterable[str], source: str) -> None:
    missing = sorted(set(columns).difference(data.columns))
    if missing:
        raise ValueError(f"Missing required columns in {source}: {missing}")


def numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").replace([np.inf, -np.inf], np.nan)


def add_zero_line(ax: mpl.axes.Axes) -> None:
    ax.axhline(0, color="#333333", lw=0.9, ls="--", alpha=0.85, zorder=1)


def set_category_ticks(ax: mpl.axes.Axes, labels: Sequence[str], rotation: int = 0) -> None:
    ax.set_xticks(np.arange(len(labels)))
    ax.set_xticklabels(labels, rotation=rotation, ha="right" if rotation else "center")


def annotate_n(ax: mpl.axes.Axes, x: float, y: float, n: int, offset: float) -> None:
    va = "bottom" if y >= 0 else "top"
    dy = offset if y >= 0 else -offset
    ax.text(x, y + dy, f"N={n}", ha="center", va=va, fontsize=7.4, color="#555555")


# -----------------------------------------------------------------------------
# Loading, harmonisation and quality controls
# -----------------------------------------------------------------------------

def load_inputs(
    final_path: Path, targets_path: Path, icr_path: Path
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    for path in (final_path, targets_path, icr_path):
        if not path.exists():
            raise FileNotFoundError(f"Input file not found: {path}")

    final = pd.read_stata(final_path, convert_categoricals=False)
    targets = pd.read_stata(targets_path, convert_categoricals=False)
    icr = pd.read_excel(icr_path, sheet_name="Country - Metadata")

    required_final = ["iso3", "year"]
    for cycle in CYCLES:
        required_final.extend(
            [
                f"gap_{cycle}",
                f"gap_{cycle}_uncond",
                f"gap_{cycle}_cond",
                f"ambition_{cycle}",
                *[f"gap_{cycle}_{scenario}" for scenario in SCENARIOS],
            ]
        )
    require_columns(final, required_final, str(final_path))
    require_columns(
        targets,
        [
            "iso3",
            "Country",
            "NDCs_number",
            "conditionality_std",
            "target_metric_std",
            "target_scope_std",
            "target_status",
            "target_year_std",
        ],
        str(targets_path),
    )
    require_columns(icr, ["iso3", "Incomegroup", "Region"], str(icr_path))

    final = final.copy()
    targets = targets.copy()
    icr = icr.copy()
    for data in (final, targets, icr):
        data["iso3"] = clean_iso(data["iso3"])

    final["year"] = numeric(final["year"])
    if final.duplicated(["iso3", "year"]).any():
        examples = final.loc[final.duplicated(["iso3", "year"], keep=False), ["iso3", "year"]]
        raise ValueError(f"Duplicate iso3-year rows in final data:\n{examples.head(20)}")
    if targets.duplicated(["iso3", "NDCs_number"]).any():
        raise ValueError("Duplicate iso3-NDCs_number rows in target data.")
    if icr.duplicated("iso3").any():
        raise ValueError("Duplicate iso3 rows in ICR.xlsx.")

    # Keep source data immutable and handle known classification gaps in code.
    # COK is absent from the supplied ICR file; ETH and VEN have no income value.
    # Only region is filled for COK because it is unambiguous geographically.
    if "COK" not in set(icr["iso3"]):
        icr = pd.concat(
            [
                icr,
                pd.DataFrame(
                    [{"iso3": "COK", "Incomegroup": pd.NA, "Region": "East Asia & Pacific"}]
                ),
            ],
            ignore_index=True,
        )

    return final, targets, icr


def targets_wide(targets: pd.DataFrame) -> pd.DataFrame:
    keep = [
        "iso3",
        "Country",
        "NDCs_number",
        "conditionality_std",
        "target_metric_std",
        "target_scope_std",
        "target_status",
        "target_year_std",
    ]
    # Reference is not the same as metric: intensity can refer to BAU or a base year.
    if "target_reference_std" in targets.columns:
        keep.append("target_reference_std")
    parts = []
    for cycle, ndc_name in (("n1", "First_NDC"), ("n2", "Second_NDC")):
        part = targets.loc[targets["NDCs_number"].eq(ndc_name), keep].copy()
        rename = {
            column: f"{column}_{cycle}"
            for column in keep
            if column not in {"iso3", "Country"}
        }
        part = part.rename(columns=rename)
        if cycle == "n2":
            part = part.drop(columns="Country")
        parts.append(part)
    wide = parts[0].merge(parts[1], on="iso3", how="outer", validate="one_to_one")
    names = targets.dropna(subset=["Country"]).drop_duplicates("iso3").set_index("iso3")["Country"]
    wide["Country"] = wide["Country"].fillna(wide["iso3"].map(names))
    return wide


def enrich_panel(
    final: pd.DataFrame, targets: pd.DataFrame, icr: pd.DataFrame
) -> pd.DataFrame:
    metadata = targets_wide(targets)
    enriched = final.merge(metadata, on="iso3", how="left", validate="many_to_one")
    available_audit_columns = [
        column for column in PAIR_AUDIT_COLUMNS if column in targets.columns
    ]
    if available_audit_columns:
        pair_audit = (
            targets[["iso3", *available_audit_columns]]
            .drop_duplicates()
            .drop_duplicates("iso3")
        )
        enriched = enriched.merge(
            pair_audit, on="iso3", how="left", validate="many_to_one"
        )
    enriched = enriched.merge(icr, on="iso3", how="left", validate="many_to_one")
    return enriched


def country_level(panel: pd.DataFrame) -> pd.DataFrame:
    value_columns: list[str] = []
    for cycle in CYCLES:
        value_columns.extend(
            [
                f"gap_{cycle}",
                f"gap_{cycle}_uncond",
                f"gap_{cycle}_cond",
                f"ambition_{cycle}",
                *[f"gap_{cycle}_{scenario}" for scenario in SCENARIOS],
            ]
        )
    aggregations: dict[str, str] = {column: "median" for column in value_columns}
    for column in [
        "Country",
        "Incomegroup",
        "Region",
        *[f"conditionality_std_{c}" for c in CYCLES],
        *[f"target_metric_std_{c}" for c in CYCLES],
        *[f"target_scope_std_{c}" for c in CYCLES],
        *[f"target_status_{c}" for c in CYCLES],
        *[f"target_year_std_{c}" for c in CYCLES],
        *PAIR_AUDIT_COLUMNS,
    ]:
        if column in panel.columns:
            aggregations[column] = "first"
    return panel.groupby("iso3", as_index=False, dropna=False).agg(aggregations)


def common_scenario_columns(cycle: str) -> list[str]:
    return [f"gap_{cycle}", *[f"gap_{cycle}_{s}" for s in SCENARIOS]]


def common_scenario_sample(data: pd.DataFrame, cycle: str) -> pd.DataFrame:
    columns = common_scenario_columns(cycle)
    return data.dropna(subset=columns).copy()


def detailed_region(iso3: object, broad_region: object) -> object:
    """Split broad source regions into nine transparent geographic groups."""
    code = str(iso3).strip().upper() if pd.notna(iso3) else ""
    if code in CENTRAL_ASIA_ISO3:
        return "Central Asia"
    if code in NORTH_AFRICA_ISO3:
        return "North Africa"
    if code in SOUTH_ASIA_OVERRIDES:
        return "South Asia"
    if code in SUB_SAHARAN_AFRICA_OVERRIDES:
        return "Sub-Saharan Africa"
    if code in EUROPE_OVERRIDES:
        return "Europe"
    if code in WESTERN_ASIA_ISO3:
        return "Middle East"
    if broad_region == "Europe & Central Asia":
        return "Europe"
    if broad_region == "Middle East & North Africa":
        return "Middle East"
    return broad_region


def add_graph_categories(data: pd.DataFrame) -> pd.DataFrame:
    """Create presentation-only groupings without altering harmonised fields."""
    out = data.copy()
    out["Detailed_region"] = [
        detailed_region(iso3, region)
        for iso3, region in zip(out["iso3"], out["Region"])
    ]
    for cycle in CYCLES:
        metric = f"target_metric_std_{cycle}"
        if metric in out:
            graph_metric = f"target_metric_graph_{cycle}"
            out[graph_metric] = out[metric].replace(
                {
                    "PER_CAPITA_TARGET": "PER_CAPITA_COMBINED",
                    "BAU_PER_CAPITA_TARGET": "PER_CAPITA_COMBINED",
                }
            )
        conditionality = f"conditionality_std_{cycle}"
        if conditionality in out:
            graph_conditionality = f"conditionality_graph_{cycle}"
            out[graph_conditionality] = out[conditionality].replace(
                {
                    "MIXED_CONDITIONALITY": "BOTH",
                    "PARTIALLY_CONDITIONAL": "BOTH",
                }
            )
    return out


def write_diagnostics(
    panel: pd.DataFrame,
    countries: pd.DataFrame,
    folders: OutputFolders,
) -> None:
    report: dict[str, object] = {
        "panel_rows": int(len(panel)),
        "countries": int(panel["iso3"].nunique()),
        "year_min": int(panel["year"].min()),
        "year_max": int(panel["year"].max()),
        "classification_missing": {
            "income_iso3": sorted(countries.loc[countries["Incomegroup"].isna(), "iso3"].tolist()),
            "region_iso3": sorted(countries.loc[countries["Region"].isna(), "iso3"].tolist()),
        },
        "cycles": {},
    }
    annual_rows = []
    for cycle in CYCLES:
        gap = f"gap_{cycle}"
        active = countries[countries[gap].notna()]
        scenario_common = common_scenario_sample(countries, cycle)
        annual_n = panel.groupby("year")[gap].count().astype(int)
        report["cycles"][cycle] = {
            "countries_main": int(active["iso3"].nunique()),
            "countries_all_scenarios_common": int(scenario_common["iso3"].nunique()),
            "country_median_gap": float(active[gap].median()),
            "annual_n": {str(int(k)): int(v) for k, v in annual_n.items()},
        }
        for year, n in annual_n.items():
            annual_rows.append(
                {
                    "cycle": CYCLE_SHORT[cycle],
                    "year": int(year),
                    "N": int(n),
                }
            )

    with (folders.diagnostics / "figure_input_audit.json").open("w", encoding="utf-8") as stream:
        json.dump(report, stream, indent=2, ensure_ascii=False)
    pd.DataFrame(annual_rows).to_csv(
        folders.diagnostics / "annual_sample_sizes.csv", index=False
    )
    countries[["iso3", "Country", "Region", "Detailed_region"]].sort_values("iso3").to_csv(
        folders.diagnostics / "detailed_region_classification.csv", index=False
    )
    detailed_summaries = [
        categorical_summary(
            countries, cycle, "Detailed_region", DETAILED_REGION_ORDER
        )
        for cycle in CYCLES
    ]
    pd.concat(detailed_summaries, ignore_index=True).assign(
        cycle=lambda frame: frame["cycle"].map(CYCLE_SHORT)
    ).to_csv(
        folders.diagnostics / "detailed_region_summary.csv", index=False
    )


# -----------------------------------------------------------------------------
# Figure 2: annual median gaps and scenario uncertainty
# -----------------------------------------------------------------------------

def annual_scenario_summary(panel: pd.DataFrame, cycle: str) -> pd.DataFrame:
    columns = common_scenario_columns(cycle)
    subset = panel.dropna(subset=columns).copy()
    rows = []
    for year, group in subset.groupby("year", sort=True):
        scenario_medians = {column: group[column].median() for column in columns}
        rows.append(
            {
                "year": int(year),
                "N": int(group["iso3"].nunique()),
                "main": scenario_medians[f"gap_{cycle}"],
                "scenario_low": min(scenario_medians.values()),
                "scenario_high": max(scenario_medians.values()),
                **scenario_medians,
            }
        )
    return pd.DataFrame(rows).sort_values("year")


def plot_annual_median_uncertainty(
    panel: pd.DataFrame,
    folders: OutputFolders,
    show: bool,
) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(10.2, 4.55), sharex=True, sharey=True)

    for ax, cycle in zip(axes, CYCLES):
        summary = annual_scenario_summary(panel, cycle)
        if summary.empty:
            continue
        color = CYCLE_COLOR[cycle]
        ax.fill_between(
            summary["year"],
            summary["scenario_low"],
            summary["scenario_high"],
            color=color,
            alpha=0.16,
            linewidth=0,
            label="Alternative-scenario range",
        )
        ax.plot(
            summary["year"],
            summary["main"],
            color=color,
            lw=2.4,
            marker="o",
            ms=4.2,
            label="Main estimate",
        )
        add_zero_line(ax)
        ax.set_title(CYCLE_SHORT[cycle])
        ax.set_xlabel("Year")
        ax.legend(frameon=False, loc="upper left")
    axes[0].set_ylabel("Median Gap-to-Target (%)")
    save_figure(fig, "Figure_2_Annual_Median_Gaps_Uncertainty", folders.main, show)


# -----------------------------------------------------------------------------
# Figure 3: global cycle medians and distributions
# -----------------------------------------------------------------------------

def plot_cycle_medians_and_distributions(
    countries: pd.DataFrame,
    folders: OutputFolders,
    show: bool,
) -> None:
    rows = []
    long_rows = []
    for cycle in CYCLES:
        gap = f"gap_{cycle}"
        scenario_sample = common_scenario_sample(countries, cycle)
        values = numeric(scenario_sample[gap]).dropna()
        scenario_medians = [
            scenario_sample[column].median()
            for column in common_scenario_columns(cycle)
        ]
        rows.append(
            {
                "cycle": CYCLE_SHORT[cycle],
                "median": values.median(),
                "lo": min(scenario_medians),
                "hi": max(scenario_medians),
                "N": values.size,
                "color": CYCLE_COLOR[cycle],
            }
        )
        long_rows.extend(
            {"cycle": CYCLE_SHORT[cycle], "gap": value} for value in values.to_numpy()
        )
    summary = pd.DataFrame(rows)
    long = pd.DataFrame(long_rows)

    fig, axes = plt.subplots(1, 2, figsize=(9.2, 4.25), gridspec_kw={"width_ratios": [0.9, 1.45]})
    ax = axes[0]
    x = np.arange(len(summary))
    ax.errorbar(
        x,
        summary["median"],
        yerr=[summary["median"] - summary["lo"], summary["hi"] - summary["median"]],
        fmt="none",
        ecolor="#333333",
        elinewidth=INTERVAL_LINEWIDTH,
        capthick=INTERVAL_CAPTHICK,
        capsize=5,
        zorder=2,
    )
    ax.scatter(x, summary["median"], s=72, c=summary["color"], zorder=3)
    span = max(1.0, float(summary["median"].abs().max()) * 0.10)
    for i, row in summary.iterrows():
        ax.text(
            i,
            row["hi"] + span,
            f"N={int(row['N'])}",
            ha="center",
            va="bottom",
            fontsize=8,
            color="#555555",
        )
    set_category_ticks(ax, summary["cycle"].tolist())
    add_zero_line(ax)
    ax.set_ylabel("Country median Gap-to-Target (%)")
    ax.set_title("(a) Cycle medians and scenario ranges")
    y_low, y_high = ax.get_ylim()
    ax.set_ylim(y_low, y_high + max(0.8, (y_high - y_low) * 0.10))

    ax = axes[1]
    box_values = [
        long.loc[long["cycle"].eq(cycle), "gap"].dropna().to_numpy()
        for cycle in ("NDC1", "NDC2")
    ]
    box = ax.boxplot(
        box_values,
        positions=[0, 1],
        widths=0.52,
        whis=(5, 95),
        showfliers=False,
        patch_artist=True,
        medianprops={"color": "#222222", "linewidth": 1.7},
        whiskerprops={"color": "#555555", "linewidth": 1.2},
        capprops={"color": "#555555", "linewidth": 1.2},
    )
    for patch, color in zip(box["boxes"], [CYCLE_COLOR["n1"], CYCLE_COLOR["n2"]]):
        patch.set_facecolor(color)
        patch.set_alpha(0.80)
    jitter_rng = np.random.default_rng(RANDOM_SEED)
    for position, values in enumerate(box_values):
        jitter = jitter_rng.uniform(-0.18, 0.18, size=len(values))
        ax.scatter(
            position + jitter,
            values,
            color="#222222",
            alpha=0.25,
            s=7,
            linewidths=0,
            zorder=2,
        )
    add_zero_line(ax)
    ax.set_xlabel("")
    set_category_ticks(ax, ["NDC1", "NDC2"])
    ax.set_ylabel("Country median Gap-to-Target (%)")
    ax.set_title("(b) Country distribution")
    save_figure(fig, "Figure_3_Global_Medians_Distributions", folders.main, show)


# -----------------------------------------------------------------------------
# Alignment position and transitions around the zero benchmark
# -----------------------------------------------------------------------------

ALIGNMENT_TOLERANCE = 1e-9
ALIGNMENT_ORDER = ("Below", "Aligned", "Above")
ALIGNMENT_COLOR = {
    "Below": "#2878B5",
    "Aligned": "#A7A9AC",
    "Above": "#D95F02",
}


def alignment_position(values: pd.Series) -> pd.Series:
    """Classify country medians relative to zero without rounding the data."""
    values = numeric(values)
    result = pd.Series(pd.NA, index=values.index, dtype="string")
    result.loc[values < -ALIGNMENT_TOLERANCE] = "Below"
    result.loc[values > ALIGNMENT_TOLERANCE] = "Above"
    result.loc[values.abs() <= ALIGNMENT_TOLERANCE] = "Aligned"
    return result


def plot_alignment_positions_and_transitions(
    countries: pd.DataFrame, folders: OutputFolders, show: bool
) -> None:
    """Summarise positions around zero and within-country cycle transitions."""
    data = countries[["iso3", "gap_n1", "gap_n2"]].copy()
    data["position_n1"] = alignment_position(data["gap_n1"])
    data["position_n2"] = alignment_position(data["gap_n2"])

    share_rows = []
    for cycle, position_column in (("NDC1", "position_n1"), ("NDC2", "position_n2")):
        observed = data[position_column].dropna()
        counts = observed.value_counts().reindex(ALIGNMENT_ORDER, fill_value=0)
        for position in ALIGNMENT_ORDER:
            share_rows.append(
                {
                    "cycle": cycle,
                    "position": position,
                    "N": int(counts[position]),
                    "share": 100.0 * counts[position] / len(observed),
                    "denominator": int(len(observed)),
                }
            )
    shares = pd.DataFrame(share_rows)

    paired = data.dropna(subset=["position_n1", "position_n2"]).copy()
    transition_order = (
        "Remained below",
        "Below to above",
        "Above to below",
        "Remained above",
        "At zero in at least one cycle",
    )
    transition_color = {
        "Remained below": "#2878B5",
        "Below to above": "#E6A141",
        "Above to below": "#756BB1",
        "Remained above": "#D95F02",
        "At zero in at least one cycle": "#A7A9AC",
    }

    def transition_label(row: pd.Series) -> str:
        n1, n2 = row["position_n1"], row["position_n2"]
        if "Aligned" in (n1, n2):
            return "At zero in at least one cycle"
        if n1 == "Below" and n2 == "Below":
            return "Remained below"
        if n1 == "Below" and n2 == "Above":
            return "Below to above"
        if n1 == "Above" and n2 == "Below":
            return "Above to below"
        return "Remained above"

    paired["transition"] = paired.apply(transition_label, axis=1)
    transition_counts = paired["transition"].value_counts().reindex(
        transition_order, fill_value=0
    )
    transition_shares = 100.0 * transition_counts / len(paired)
    nonzero = paired.loc[
        paired["position_n1"].ne("Aligned") & paired["position_n2"].ne("Aligned")
    ]
    changed_sign = (
        100.0 * (nonzero["position_n1"] != nonzero["position_n2"]).mean()
        if len(nonzero)
        else np.nan
    )

    fig, axes = plt.subplots(
        1, 2, figsize=(10.0, 4.5), gridspec_kw={"width_ratios": [1.0, 1.35]}
    )

    ax = axes[0]
    bottoms = np.zeros(2)
    x = np.arange(2)
    for position in ALIGNMENT_ORDER:
        part = shares.loc[shares["position"].eq(position)].set_index("cycle")
        values = part.loc[["NDC1", "NDC2"], "share"].to_numpy()
        counts = part.loc[["NDC1", "NDC2"], "N"].to_numpy()
        bars = ax.bar(
            x, values, bottom=bottoms, width=0.58,
            color=ALIGNMENT_COLOR[position], label=position,
        )
        for bar, value, count, bottom in zip(bars, values, counts, bottoms):
            if value >= 4.0:
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bottom + value / 2,
                    f"{value:.1f}%\n(N={int(count)})",
                    ha="center", va="center", fontsize=7.6,
                    color="white" if position != "Aligned" else "#333333",
                )
        bottoms += values
    set_category_ticks(ax, ["NDC1", "NDC2"])
    ax.set_ylim(0, 100)
    ax.set_ylabel("Share of countries (%)")
    ax.set_title("(a) Position relative to the annual pathway")
    ax.legend(frameon=False, loc="lower center", bbox_to_anchor=(0.5, -0.25), ncol=3)

    ax = axes[1]
    left = 0.0
    for transition in transition_order:
        value = float(transition_shares[transition])
        if value <= 0:
            continue
        ax.barh(
            [0], [value], left=left, height=0.42,
            color=transition_color[transition], label=transition,
        )
        if value >= 4.0:
            ax.text(
                left + value / 2, 0,
                f"{value:.1f}%\n(N={int(transition_counts[transition])})",
                ha="center", va="center", fontsize=7.6,
                color="white" if transition != "At zero in at least one cycle" else "#333333",
            )
        left += value
    ax.set_xlim(0, 100)
    ax.set_yticks([])
    ax.set_xlabel("Share of countries in the common sample (%)")
    ax.set_title("(b) Change in position between NDC cycles")
    ax.legend(frameon=False, loc="lower center", bbox_to_anchor=(0.5, -0.37), ncol=2)
   

    shares.to_csv(folders.diagnostics / "alignment_position_shares.csv", index=False)
    paired[["iso3", "position_n1", "position_n2", "transition"]].to_csv(
        folders.diagnostics / "alignment_position_transitions.csv", index=False
    )
    save_figure(
        fig,
        "Appendix_Aggregate_Alignment_Positions_and_Transitions",
        folders.appendix,
        show,
    )


# -----------------------------------------------------------------------------
# Generic categorical median plot
# -----------------------------------------------------------------------------

def categorical_summary(
    countries: pd.DataFrame,
    cycle: str,
    category_column: str,
    categories: Sequence[str],
) -> pd.DataFrame:
    gap = f"gap_{cycle}"
    scenario_columns = common_scenario_columns(cycle)
    rows = []
    for category in categories:
        group = countries.loc[countries[category_column].eq(category)].copy()
        group = group.dropna(subset=scenario_columns)
        if group.empty:
            continue
        scenario_medians = [group[column].median() for column in scenario_columns]
        rows.append(
            {
                "cycle": cycle,
                "category": category,
                "median": group[gap].median(),
                "lo": min(scenario_medians),
                "hi": max(scenario_medians),
                "N": group["iso3"].nunique(),
            }
        )
    return pd.DataFrame(rows)


def plot_categorical_medians(
    countries: pd.DataFrame,
    category_base: str,
    categories: Sequence[str],
    labels: dict[str, str],
    title: str,
    stem: str,
    folder: Path,
    show: bool,
    *,
    rotate: int = 0,
    flag_small_samples: bool = False,
) -> None:
    pieces = []
    for cycle in CYCLES:
        column = f"{category_base}_{cycle}" if f"{category_base}_{cycle}" in countries else category_base
        pieces.append(
            categorical_summary(
                countries, cycle, column, categories
            )
        )
    summary = pd.concat(pieces, ignore_index=True)
    visible_categories = [c for c in categories if c in set(summary["category"])]
    x = np.arange(len(visible_categories), dtype=float)
    offsets = {"n1": -0.16, "n2": 0.16}

    fig, ax = plt.subplots(figsize=(max(7.4, 1.15 * len(visible_categories)), 4.75))
    all_bounds = pd.concat([numeric(summary["lo"]), numeric(summary["hi"])])
    offset_y = max(0.75, float(all_bounds.max() - all_bounds.min()) * 0.035)

    for cycle in CYCLES:
        part = summary[summary["cycle"].eq(cycle)].set_index("category")
        available = [c for c in visible_categories if c in part.index]
        xpos = np.array([visible_categories.index(c) for c in available], dtype=float) + offsets[cycle]
        values = part.loc[available, "median"].to_numpy()
        lo = part.loc[available, "lo"].to_numpy()
        hi = part.loc[available, "hi"].to_numpy()
        n = part.loc[available, "N"].astype(int).to_numpy()
        ax.errorbar(
            xpos,
            values,
            yerr=[values - lo, hi - values],
            fmt="o",
            color=CYCLE_COLOR[cycle],
            ecolor=CYCLE_COLOR[cycle],
            elinewidth=INTERVAL_LINEWIDTH,
            capthick=INTERVAL_CAPTHICK,
            capsize=3.5,
            markersize=6.0,
            label=CYCLE_SHORT[cycle],
            zorder=3,
        )
        for xx, upper, nn in zip(xpos, hi, n):
            sample_label = f"N={int(nn)}"
            if flag_small_samples and int(nn) < 5:
                sample_label += "*"
            ax.text(
                xx,
                upper + offset_y,
                sample_label,
                ha="center",
                va="bottom",
                fontsize=7.6,
                color="#555555",
            )

    add_zero_line(ax)
    set_category_ticks(ax, [labels.get(c, c) for c in visible_categories], rotate)
    ax.set_ylabel("Country median Gap-to-Target (%)")
    ax.legend(frameon=False, ncol=2)
    y_low, y_high = ax.get_ylim()
    ax.set_ylim(y_low, y_high + max(0.8, (y_high - y_low) * 0.12))
    ax.text(
        0.995,
        0.015,
        "Bars: range across the main estimate and five alternative scenarios",
        transform=ax.transAxes,
        ha="right",
        va="bottom",
        fontsize=7.5,
        color="#555555",
    )
    if flag_small_samples:
        ax.text(
            0.005,
            0.015,
            "* Fewer than five countries; interpret cautiously",
            transform=ax.transAxes,
            ha="left",
            va="bottom",
            fontsize=7.5,
            color="#555555",
        )
    save_figure(fig, stem, folder, show)


# -----------------------------------------------------------------------------
# Figure 6: genuinely paired conditional versus unconditional targets
# -----------------------------------------------------------------------------

def paired_conditionality_data(countries: pd.DataFrame, cycle: str) -> pd.DataFrame:
    uncond = f"gap_{cycle}_uncond"
    cond = f"gap_{cycle}_cond"
    structure = f"conditionality_std_{cycle}"
    require_columns(countries, [uncond, cond, structure], "country-level data")
    # MIXED_CONDITIONALITY is the auditable case in which both components are
    # actually reported. Replicated single-component targets are excluded.
    return countries.loc[
        countries[structure].eq("MIXED_CONDITIONALITY"),
        ["iso3", "Country", uncond, cond],
    ].dropna(subset=[uncond, cond])


def plot_paired_conditionality(
    countries: pd.DataFrame,
    folders: OutputFolders,
    show: bool,
) -> None:
    rows = []
    for cycle in CYCLES:
        data = paired_conditionality_data(countries, cycle)
        u = f"gap_{cycle}_uncond"
        c = f"gap_{cycle}_cond"
        rows.extend(
            [
                {
                    "cycle": cycle,
                    "commitment": "Unconditional",
                    "median": data[u].median(),
                    "N": len(data),
                },
                {
                    "cycle": cycle,
                    "commitment": "Conditional",
                    "median": data[c].median(),
                    "N": len(data),
                },
            ]
        )
    summary = pd.DataFrame(rows)
    fig, ax = plt.subplots(figsize=(7.6, 4.6))
    x = np.arange(2, dtype=float)
    width = 0.34
    for cycle, offset in (("n1", -width / 2), ("n2", width / 2)):
        part = summary[summary["cycle"].eq(cycle)].set_index("commitment")
        values = part.loc[["Unconditional", "Conditional"], "median"].to_numpy()
        bars = ax.bar(
            x + offset,
            values,
            width,
            color=CYCLE_COLOR[cycle],
            alpha=0.88,
            label=CYCLE_SHORT[cycle],
        )
        for bar, value, n in zip(
            bars,
            values,
            part.loc[["Unconditional", "Conditional"], "N"].astype(int),
        ):
            va = "bottom" if value >= 0 else "top"
            dy = 0.45 if value >= 0 else -0.45
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                value + dy,
                f"N={n}",
                ha="center",
                va=va,
                fontsize=8,
                color="#555555",
            )
    add_zero_line(ax)
    set_category_ticks(ax, ["Unconditional", "Conditional"])
    ax.set_ylabel("Country median Gap-to-Target (%)")
    ax.legend(frameon=False, ncol=2)
    ax.text(
        0.995,
        0.015,
        "Alternative-scenario variants are not available separately for conditional and unconditional targets.",
        transform=ax.transAxes,
        ha="right",
        va="bottom",
        fontsize=7.4,
        color="#555555",
    )
    save_figure(fig, "Appendix_Paired_Target_Components", folders.appendix, show)


# -----------------------------------------------------------------------------
# Ambition and alignment
# -----------------------------------------------------------------------------

def label_all_countries(
    ax: mpl.axes.Axes,
    data: pd.DataFrame,
    x_column: str,
    y_column: str,
    fontsize: float = 4.8,
    label_column: str = "iso3",
) -> None:
    """Label every country and repel labels when adjustText is available."""
    texts = [
        ax.text(
            row[x_column], row[y_column], str(row[label_column]),
            fontsize=fontsize, color="#333333", ha="center", va="center", zorder=6,
        )
        for _, row in data.iterrows()
    ]
    try:
        from adjustText import adjust_text

        adjust_text(
            texts,
            ax=ax,
            x=data[x_column].to_numpy(),
            y=data[y_column].to_numpy(),
            expand=(1.03, 1.08),
            force_text=(0.22, 0.32),
            force_points=(0.08, 0.12),
            arrowprops={"arrowstyle": "-", "color": "#888888", "lw": 0.28},
        )
    except ImportError:
        warnings.warn(
            "adjustText is not installed; ISO3 labels are shown without collision adjustment.",
            RuntimeWarning,
        )

def plot_ambition_alignment(
    countries: pd.DataFrame, folders: OutputFolders, show: bool
) -> None:
    fig, axes = plt.subplots(2, 1, figsize=(9.4, 10.2), sharex=False, sharey=False)
    for ax, cycle in zip(axes, CYCLES):
        ambition = f"ambition_{cycle}"
        gap = f"gap_{cycle}"
        data = countries[["iso3", ambition, gap]].dropna().copy()
        if data.empty:
            ax.text(0.5, 0.5, "No ambition data", transform=ax.transAxes, ha="center")
            continue
        ax.scatter(
            data[ambition], data[gap], s=20, color=CYCLE_COLOR[cycle], alpha=0.48,
            edgecolor="white", linewidth=0.25
        )
        ax.axhline(0, color="#333333", lw=0.8, ls="--")
        ax.axvline(0, color="#333333", lw=0.8, ls="--")
        # Preserve every observation while making the central mass readable.
        # A symmetric-log scale is linear around zero and logarithmic in both
        # tails; it is preferable here to dropping or winsorising extremes.
        ax.set_xscale("symlog", linthresh=10, linscale=1.0, base=10)

        # A LOWESS curve is descriptive only and is explicitly labelled as such.
        try:
            from statsmodels.nonparametric.smoothers_lowess import lowess

            smooth = lowess(data[gap], data[ambition], frac=0.55, return_sorted=True)
            ax.plot(smooth[:, 0], smooth[:, 1], color="#222222", lw=1.7, label="LOWESS fit")
            ax.legend(frameon=False, loc="best")
        except Exception:
            # LOWESS is optional. The scatter plot remains valid if statsmodels
            # or one of its compiled SciPy dependencies is unavailable.
            pass
        label_all_countries(ax, data, ambition, gap, fontsize=4.6)
        ax.set_title(f"{CYCLE_SHORT[cycle]} (N={len(data)})")
        ax.set_xlabel("Target ambition (%)")
        ax.set_ylabel("Country median Gap-to-Target (%)")
    fig.text(
        0.5,
        -0.01,
        "Dashed lines identify zero ambition and exact pathway alignment. The solid line is a locally weighted regression fit and is descriptive only.",
        ha="center",
        fontsize=7.7,
        color="#555555",
    )
    save_figure(fig, "Appendix_Ambition_and_Alignment", folders.appendix, show)


AMBITION_ALIGNMENT_QUADRANT_COLORS = {
    "Ambition < 0; Gap > 0": "#D81B60",
    "Ambition < 0; Gap < 0": "#009E73",
    "Ambition > 0; Gap > 0": "#E69F00",
    "Ambition > 0; Gap < 0": "#0072B2",
}


def ambition_alignment_quadrant(ambition: pd.Series, gap: pd.Series) -> pd.Series:
    """Classify countries using the signs that define ambition and alignment."""
    ambition = numeric(ambition)
    gap = numeric(gap)
    result = pd.Series(pd.NA, index=ambition.index, dtype="string")
    valid = ambition.notna() & gap.notna()
    result.loc[valid & ambition.lt(0) & gap.gt(0)] = "Ambition < 0; Gap > 0"
    result.loc[valid & ambition.lt(0) & gap.lt(0)] = "Ambition < 0; Gap < 0"
    result.loc[valid & ambition.ge(0) & gap.gt(0)] = "Ambition > 0; Gap > 0"
    result.loc[valid & ambition.ge(0) & gap.lt(0)] = "Ambition > 0; Gap < 0"
    result.loc[valid & gap.abs().le(ALIGNMENT_TOLERANCE)] = "At exact alignment"
    return result


def plot_ambition_alignment_transitions(
    countries: pd.DataFrame, folders: OutputFolders, show: bool
) -> None:
    """Show NDC1 quadrants and NDC2 positions for the common country sample."""
    required = ["iso3", "ambition_n1", "ambition_n2", "gap_n1", "gap_n2"]
    audit_metadata = [column for column in PAIR_AUDIT_COLUMNS if column in countries]
    data = countries[required + audit_metadata].copy()
    data["origin_quadrant"] = ambition_alignment_quadrant(
        data["ambition_n1"], data["gap_n1"]
    )
    n1 = data.dropna(subset=["ambition_n1", "gap_n1", "origin_quadrant"]).copy()
    common = data.dropna(
        subset=["ambition_n1", "gap_n1", "ambition_n2", "gap_n2", "origin_quadrant"]
    ).copy()
    if n1.empty or common.empty:
        warnings.warn("Insufficient ambition and Gap-to-Target data for the transition figure.")
        return

    position_n1 = alignment_position(common["gap_n1"])
    position_n2 = alignment_position(common["gap_n2"])
    common["transition"] = np.select(
        [
            position_n1.eq("Below") & position_n2.eq("Above"),
            position_n1.eq("Above") & position_n2.eq("Below"),
            position_n1.eq("Below") & position_n2.eq("Below"),
            position_n1.eq("Above") & position_n2.eq("Above"),
        ],
        ["Below to above", "Above to below", "Remained below", "Remained above"],
        default="At zero in at least one cycle",
    )
    transition_order = (
        "Remained below", "Below to above", "Above to below", "Remained above",
        "At zero in at least one cycle",
    )
    counts = common["transition"].value_counts().reindex(transition_order, fill_value=0)
    shares = 100.0 * counts / len(common)
    pd.DataFrame(
        {
            "transition": transition_order,
            "N": [int(counts[label]) for label in transition_order],
            "share_percent": [float(shares[label]) for label in transition_order],
            "common_sample_N": len(common),
        }
    ).to_csv(
        folders.diagnostics / "ambition_alignment_transition_counts.csv", index=False
    )

    if "pair_target_comparable_flag" in common:
        common["architecture_changed"] = numeric(
            common["pair_target_comparable_flag"]
        ).eq(0)
    else:
        common["architecture_changed"] = False
    common["iso3_label"] = common["iso3"].astype(str) + np.where(
        common["architecture_changed"], "†", ""
    )
    audit_columns = [
        "iso3", "ambition_n1", "gap_n1", "ambition_n2", "gap_n2",
        "origin_quadrant", "transition", "architecture_changed", *audit_metadata,
    ]
    common[audit_columns].to_csv(
        folders.diagnostics / "ambition_alignment_transition_audit.csv", index=False
    )

    fig, axes = plt.subplots(2, 1, figsize=(10.2, 12.0), sharex=False, sharey=False)
    panels = (
        (axes[0], n1, "ambition_n1", "gap_n1", f"(a) NDC1 positions (N={len(n1)})"),
        (axes[1], common, "ambition_n2", "gap_n2", f"(b) NDC2 positions for the common sample (N={len(common)})"),
    )
    marker_by_transition = {
        "Remained below": "o",
        "Remained above": "o",
        "Below to above": "^",
        "Above to below": "v",
        "At zero in at least one cycle": "s",
    }

    for panel_index, (ax, part, x_column, y_column, title) in enumerate(panels):
        ax.axhline(0, color="#333333", lw=0.85, ls="--", zorder=1)
        ax.axvline(0, color="#333333", lw=0.85, ls="--", zorder=1)
        ax.set_xscale("symlog", linthresh=10, linscale=1.0, base=10)
        if panel_index == 0:
            for quadrant, color in AMBITION_ALIGNMENT_QUADRANT_COLORS.items():
                selected = part.loc[part["origin_quadrant"].eq(quadrant)]
                ax.scatter(
                    selected[x_column], selected[y_column], s=20, marker="o",
                    color=color, alpha=0.72, edgecolor="white", linewidth=0.25,
                    zorder=3,
                )
            labels = part.assign(iso3_label=part["iso3"].astype(str))
        else:
            for transition in transition_order:
                transitioned = part.loc[part["transition"].eq(transition)]
                marker = marker_by_transition[transition]
                for quadrant, color in AMBITION_ALIGNMENT_QUADRANT_COLORS.items():
                    selected = transitioned.loc[
                        transitioned["origin_quadrant"].eq(quadrant)
                    ]
                    ax.scatter(
                        selected[x_column], selected[y_column],
                        s=34 if marker != "o" else 22,
                        marker=marker, color=color, alpha=0.84,
                        edgecolor="#222222" if marker != "o" else "white",
                        linewidth=0.7 if marker != "o" else 0.25,
                        zorder=3,
                    )
            labels = part

        try:
            from statsmodels.nonparametric.smoothers_lowess import lowess

            smooth = lowess(
                part[y_column], part[x_column], frac=0.55, return_sorted=True
            )
            ax.plot(smooth[:, 0], smooth[:, 1], color="#222222", lw=1.7, zorder=4)
        except Exception:
            pass
        label_all_countries(
            ax, labels, x_column, y_column,
            fontsize=4.25, label_column="iso3_label",
        )
        ax.set_title(title)
        ax.set_xlabel("Target ambition (%)")
        ax.set_ylabel("Country median Gap-to-Target (%)")

    quadrant_handles = [
        Line2D(
            [0], [0], marker="o", linestyle="none", markersize=5.5,
            markerfacecolor=color, markeredgecolor="none", label=label,
        )
        for label, color in AMBITION_ALIGNMENT_QUADRANT_COLORS.items()
    ]
    transition_handles = [
        Line2D(
            [0], [0], marker=marker_by_transition[label], linestyle="none",
            markersize=6.0, markerfacecolor="#777777", markeredgecolor="#222222",
            label=f"{label}: N={int(counts[label])} ({shares[label]:.1f}%)",
        )
        for label in transition_order if counts[label] > 0
    ]
    lowess_handle = Line2D([0], [0], color="#222222", lw=1.7, label="LOWESS fit")
    fig.legend(
        handles=quadrant_handles + transition_handles + [lowess_handle],
        frameon=False, loc="lower center", bbox_to_anchor=(0.5, 0.047),
        ncol=3, fontsize=7.2,
    )
    architecture_count = int(common["architecture_changed"].sum())
    fig.text(
        0.5, 0.012,
        "Colours retain each country's NDC1 quadrant. Upward and downward triangles identify changes in the sign of the Gap-to-Target; circles indicate no change. "
        f"† Major target-architecture change between cycles (N={architecture_count}). The LOWESS fit is descriptive only.",
        ha="center", fontsize=7.0, color="#555555", wrap=True,
    )
    fig.subplots_adjust(
        left=0.09, right=0.99, top=0.98, bottom=0.205, hspace=0.30
    )
    fig.set_layout_engine("none")
    output = folders.main / "Figure_8_Transitions_in_Ambition_and_Alignment.png"
    temporary = output.with_name(f".{output.name}.tmp")
    fig.savefig(temporary, format="png", bbox_inches="tight")
    temporary.replace(output)
    if show:
        plt.show()
    plt.close(fig)


# -----------------------------------------------------------------------------
# Appendix figures
# -----------------------------------------------------------------------------

def plot_target_architecture(
    targets: pd.DataFrame,
    final_iso3: set[str],
    folders: OutputFolders,
    show: bool,
) -> None:
    targets = targets[targets["iso3"].isin(final_iso3)].copy()
    targets["target_metric_graph"] = targets["target_metric_std"].replace(
        {
            "PER_CAPITA_TARGET": "PER_CAPITA_COMBINED",
            "BAU_PER_CAPITA_TARGET": "PER_CAPITA_COMBINED",
        }
    )
    fig, axes = plt.subplots(1, 3, figsize=(12.0, 4.2))
    specs = [
        ("target_metric_graph", TARGET_TYPE_ORDER, TARGET_TYPE_LABEL, "Target metric"),
        ("target_scope_std", (*COVERAGE_ORDER, "UNKNOWN"), {**COVERAGE_LABEL, "UNKNOWN": "Unknown"}, "Target coverage"),
        ("conditionality_std", CONDITIONALITY_ORDER, CONDITIONALITY_LABEL, "Conditionality"),
    ]
    cycle_map = {"First_NDC": "NDC1", "Second_NDC": "NDC2"}
    for ax, (column, order, labels, title) in zip(axes, specs):
        counts = (
            targets.assign(cycle=targets["NDCs_number"].map(cycle_map))
            .groupby([column, "cycle"], dropna=False)
            .size()
            .rename("N")
            .reset_index()
        )
        visible = [c for c in order if c in set(counts[column])]
        pivot = counts.pivot(index=column, columns="cycle", values="N").reindex(visible).fillna(0)
        xpos = np.arange(len(visible))
        width = 0.36
        for offset, cyc, color in ((-width / 2, "NDC1", CYCLE_COLOR["n1"]), (width / 2, "NDC2", CYCLE_COLOR["n2"])):
            values = pivot.get(cyc, pd.Series(0, index=visible)).to_numpy()
            bars = ax.bar(xpos + offset, values, width, label=cyc, color=color, alpha=0.88)
            ax.bar_label(bars, fmt="%.0f", fontsize=7, padding=1)
        set_category_ticks(ax, [labels.get(c, c) for c in visible], rotation=38)
        ax.set_ylabel("Number of countries")
        ax.set_title(title)
        ax.legend(frameon=False, ncol=2)
    save_figure(fig, "Appendix_Target_Architecture", folders.appendix, show)


def plot_scenario_specific_medians(
    countries: pd.DataFrame, folders: OutputFolders, show: bool
) -> None:
    rows = []
    for cycle in CYCLES:
        sample = common_scenario_sample(countries, cycle)
        for scenario in ("main", *SCENARIOS):
            column = f"gap_{cycle}" if scenario == "main" else f"gap_{cycle}_{scenario}"
            rows.append(
                {
                    "cycle": CYCLE_SHORT[cycle],
                    "scenario": SCENARIO_LABEL[scenario],
                    "median": sample[column].median(),
                    "N": len(sample),
                }
            )
    data = pd.DataFrame(rows)
    order = [SCENARIO_LABEL[s] for s in ("main", *SCENARIOS)]
    fig, ax = plt.subplots(figsize=(7.8, 4.8))
    ypos = np.arange(len(order), dtype=float)
    offsets = {"NDC1": -0.12, "NDC2": 0.12}
    for cycle, color, marker in (
        ("NDC1", CYCLE_COLOR["n1"], "o"),
        ("NDC2", CYCLE_COLOR["n2"], "s"),
    ):
        part = data.loc[data["cycle"].eq(cycle)].set_index("scenario").reindex(order)
        ax.scatter(
            part["median"].to_numpy(),
            ypos + offsets[cycle],
            color=color,
            marker=marker,
            s=48,
            label=cycle,
            zorder=3,
        )
        for value, yy in zip(part["median"].to_numpy(), ypos + offsets[cycle]):
            ax.text(
                value + 0.10,
                yy,
                f"{value:.1f}",
                ha="left",
                va="center",
                fontsize=7.6,
                color=color,
            )
    ax.axvline(0, color="#555555", ls="--", lw=0.9)
    ax.set_yticks(ypos)
    ax.set_yticklabels(order)
    ax.invert_yaxis()
    ax.set_xlabel("Country median Gap-to-Target (%)")
    ax.set_ylabel("")
    n1 = int(data.loc[data["cycle"].eq("NDC1"), "N"].iloc[0])
    n2 = int(data.loc[data["cycle"].eq("NDC2"), "N"].iloc[0])
    ax.legend(title="", frameon=False, ncol=2, loc="best")
    save_figure(fig, "Appendix_Scenario_Specific_Medians", folders.appendix, show)


def extract_natural_earth_archive(archive: Path, destination: Path) -> Path:
    destination.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(archive) as zipped:
        zipped.extractall(destination)
    candidates = sorted(destination.glob("*admin_0_countries.shp"))
    if not candidates:
        candidates = sorted(destination.glob("*.shp"))
    if not candidates:
        raise FileNotFoundError(
            f"No country shapefile was found after extracting {archive}."
        )
    return candidates[0]


def download_natural_earth(cache_dir: Path) -> Path:
    """Download and cache the official Natural Earth Admin 0 country layer."""
    cache_dir.mkdir(parents=True, exist_ok=True)
    extracted = cache_dir / "ne_110m_admin_0_countries"
    existing = sorted(extracted.glob("*admin_0_countries.shp"))
    if existing:
        return existing[0]

    archive = cache_dir / "ne_110m_admin_0_countries.zip"
    if not archive.exists():
        print("Downloading Natural Earth country boundaries for the maps...")
        try:
            urllib.request.urlretrieve(NATURAL_EARTH_URL, archive)
        except Exception as exc:
            raise RuntimeError(
                "The Natural Earth boundary archive could not be downloaded. "
                "Download ne_110m_admin_0_countries.zip manually and pass its "
                "path with --world-shapefile."
            ) from exc
    return extract_natural_earth_archive(archive, extracted)


def load_world_boundaries(
    world_shapefile: Path | None, cache_dir: Path
):
    import geopandas as gpd

    if world_shapefile is not None:
        if not world_shapefile.exists():
            raise FileNotFoundError(f"World shapefile not found: {world_shapefile}")
        source = world_shapefile
        if world_shapefile.suffix.lower() == ".zip":
            source = extract_natural_earth_archive(
                world_shapefile, cache_dir / world_shapefile.stem
            )
    else:
        source = download_natural_earth(cache_dir)
    return gpd.read_file(source)


def map_iso3_column(world) -> pd.Series:
    """Construct the most complete ISO3 merge key available in Natural Earth."""
    candidates = [
        column for column in ("ISO_A3", "iso_a3", "ADM0_A3", "WB_A3")
        if column in world.columns
    ]
    if not candidates:
        raise ValueError("The world boundary file contains no recognised ISO3 column.")
    result = pd.Series(pd.NA, index=world.index, dtype="string")
    for column in candidates:
        values = clean_iso(world[column]).replace({"-99": pd.NA, "": pd.NA})
        result = result.fillna(values)
    return result


def plot_world_maps(
    countries: pd.DataFrame,
    folders: OutputFolders,
    show: bool,
    world_shapefile: Path | None = None,
) -> None:
    try:
        import geopandas as gpd
    except ImportError as exc:
        raise RuntimeError(
            "The maps require GeoPandas. Install it once with: "
            "python -m pip install geopandas pyogrio shapely"
        ) from exc

    try:
        world = load_world_boundaries(
            world_shapefile,
            PROJECT_ROOT / "data" / "processed" / "geospatial" / "Natural_Earth_cache",
        )
        world = world.copy()
        world["map_iso3"] = map_iso3_column(world)
    except Exception as exc:
        raise RuntimeError(f"The world maps could not be generated: {exc}") from exc

    available = pd.concat(
        [numeric(countries[f"gap_{cycle}"]).abs() for cycle in CYCLES],
        ignore_index=True,
    ).dropna()
    bound = float(available.quantile(0.95)) if len(available) else 1.0
    bound = max(bound, 1.0)
    norm = TwoSlopeNorm(vmin=-bound, vcenter=0.0, vmax=bound)

    fig, axes = plt.subplots(1, 2, figsize=(12.0, 4.7))
    for ax, cycle in zip(axes, CYCLES):
        gap = f"gap_{cycle}"
        data = countries[["iso3", gap]].dropna().rename(columns={gap: "gap"})
        merged = world.merge(data, left_on="map_iso3", right_on="iso3", how="left")

        # Plot available estimates and missing cycle-specific observations in
        # separate layers. Hatching prevents missing values from being mistaken
        # for observations located close to the zero-alignment threshold.
        available_map = merged.loc[merged["gap"].notna()]
        missing_map = merged.loc[merged["gap"].isna()]
        available_map.plot(
            column="gap",
            cmap="RdBu_r",
            norm=norm,
            linewidth=0.20,
            edgecolor="#777777",
            legend=False,
            ax=ax,
        )
        if not missing_map.empty:
            missing_map.plot(
                ax=ax,
                facecolor="#D9D9D9",
                edgecolor="white",
                linewidth=0.25,
                hatch="///",
                zorder=3,
            )
        panel = "(a)" if cycle == "n1" else "(b)"
        ax.set_title(f"{panel} {CYCLE_SHORT[cycle]}")
        ax.set_axis_off()

    scalar_map = mpl.cm.ScalarMappable(norm=norm, cmap="RdBu_r")
    scalar_map.set_array([])
    colorbar = fig.colorbar(
        scalar_map,
        ax=axes,
        orientation="horizontal",
        fraction=0.055,
        pad=0.045,
        aspect=38,
        extend="both",
    )
    colorbar.set_label("Country median Gap-to-Target (%)")
    missing_legend = Patch(
        facecolor="#D9D9D9",
        edgecolor="#777777",
        hatch="///",
        label="No Gap estimate for the NDC cycle",
    )
    fig.legend(
        handles=[missing_legend],
        frameon=True,
        loc="lower left",
        bbox_to_anchor=(0.02, 0.02),
        fontsize=8.2,
    )
    fig.subplots_adjust(left=0.01, right=0.99, top=0.95, bottom=0.18, wspace=0.02)
    fig.set_layout_engine("none")
    for extension in ("png",):
        output = folders.appendix / f"Appendix_World_Maps_Country_Medians.{extension}"
        temporary = output.with_name(f".{output.name}.tmp")
        fig.savefig(temporary, format=extension, bbox_inches="tight")
        temporary.replace(output)
    if show:
        plt.show()
    plt.close(fig)


# -----------------------------------------------------------------------------
# Country-level scenario sensitivity for BAU-referenced targets only
# -----------------------------------------------------------------------------

def reference_group(reference: object, metric: object) -> tuple[str, str]:
    """Use declared reference first; never infer BAU from scenario sensitivity."""
    ref = "" if pd.isna(reference) else str(reference).strip().upper()
    met = "" if pd.isna(metric) else str(metric).strip().upper()
    tokens = set(re.split(r"[^A-Z0-9]+", ref))
    if "BAU" in tokens and ("BASE_YEAR" in ref or "BASEYEAR" in tokens):
        return "OTHER_OR_UNRESOLVED", "ambiguous_reference"
    if "BAU" in tokens:
        return "BAU", "declared_reference"
    if ref in {"BASE_YEAR", "BASE_YEAR_REDUCTION", "INTENSITY_BASE", "PER_CAPITA_BASE"}:
        return "BASE_YEAR", "declared_reference"
    # Only unambiguous metrics can supply a fallback when the reference is absent.
    if ref in {"", "UNKNOWN", "OTHER_OR_UNKNOWN"}:
        if met in {"BAU_REDUCTION", "BAU_PER_CAPITA_TARGET"}:
            return "BAU", "explicit_metric_fallback"
        if met == "BASE_YEAR_REDUCTION":
            return "BASE_YEAR", "explicit_metric_fallback"
    return "OTHER_OR_UNRESOLVED", "not_assigned"


def country_scenario_summary(panel: pd.DataFrame, cycle: str) -> pd.DataFrame:
    """Median of each scenario on identical years for each country and cycle.

    The range includes the main estimate. The selected dumbbell endpoint is the
    alternative with the greatest absolute deviation (ties follow SCENARIOS).
    Entirely unavailable alternatives never exclude available alternatives.
    Use common years among available series where possible; otherwise retain
    available-period medians with an explicit differing-years flag.
    """
    columns = common_scenario_columns(cycle)
    main = f"gap_{cycle}"
    rows = []
    for iso3, raw in panel.groupby("iso3", sort=True):
        group = raw.copy()
        for column in columns:
            group[column] = numeric(group[column])
        active = group.loc[group[main].notna()]
        if active.empty:
            continue
        metric_col = f"target_metric_std_{cycle}"
        ref_col = f"target_reference_std_{cycle}"
        def metadata(column: str) -> object:
            values = group[column].dropna().unique() if column in group else []
            if len(values) > 1:
                raise ValueError(f"Conflicting {column} for {iso3}/{cycle}")
            return values[0] if len(values) else ""
        reference, metric = metadata(ref_col), metadata(metric_col)
        family, source = reference_group(reference, metric)
        name = str(metadata("Country")).strip() or str(iso3)
        available_scenarios = [s for s in SCENARIOS if active[f"gap_{cycle}_{s}"].notna().any()]
        available_columns = [main, *[f"gap_{cycle}_{s}" for s in available_scenarios]]
        common = active.dropna(subset=available_columns)
        comparison = common if len(common) else active
        available_years = sorted(active["year"].astype(int).tolist())
        common_years = sorted(common["year"].astype(int).tolist())
        row = {
            "iso3": iso3, "Country": name, "cycle": CYCLE_SHORT[cycle],
            "reference_group": family, "classification_source": source,
            "target_reference": reference, "target_metric": metric,
            "n_years_main": len(active), "n_years_common": len(common),
            "years_main": ";".join(map(str, available_years)),
            "years_common": ";".join(map(str, common_years)),
            "main_full_period": float(active[main].median()),
            "status": ("no_alternatives" if not available_scenarios else
                       "common_years" if len(common) else "different_years"),
            "n_scenarios_available": len(available_scenarios),
            "scenarios_available": ";".join(available_scenarios),
            "scenarios_missing": ";".join(s for s in SCENARIOS if s not in available_scenarios),
        }
        row["main"] = float(common[main].median()) if len(common) else row["main_full_period"]
        for scenario in SCENARIOS:
            observed = comparison.loc[comparison[f"gap_{cycle}_{scenario}"].notna()]
            row[scenario] = float(observed[f"gap_{cycle}_{scenario}"].median()) if len(observed) else np.nan
            row[f"years_{scenario}"] = ";".join(map(str, sorted(observed["year"].astype(int))))
        if available_scenarios:
            values = [row["main"], *[row[s] for s in available_scenarios]]
            row["low"], row["high"] = min(values), max(values)
            extreme = max(available_scenarios, key=lambda s: abs(row[s] - row["main"]))
            row["farthest_scenario"] = extreme
            row["farthest_value"] = row[extreme]
        else:
            row["low"] = row["high"] = row["main"]
            row["farthest_scenario"], row["farthest_value"] = "", np.nan
        rows.append(row)
    return pd.DataFrame(rows)


def plot_country_scenario_sensitivity(
    panel: pd.DataFrame, folders: OutputFolders, show: bool,
    style: str = "range", countries_per_page: int = 32,
) -> None:
    if countries_per_page < 1:
        raise ValueError("--countries-per-page must be positive")
    modes = ("range", "dumbbell") if style == "both" else (style,)
    for cycle in CYCLES:
        summary = country_scenario_summary(panel, cycle)
        summary.to_csv(folders.diagnostics / f"Country_Scenarios_{CYCLE_SHORT[cycle]}.csv", index=False)
        if summary.empty:
            continue
        non_bau = summary["reference_group"].ne("BAU").sum()
        if non_bau:
            print(f"{CYCLE_SHORT[cycle]}: {non_bau} non-BAU countries retained in diagnostics only for country scenario plots.")
        for family, family_label in REFERENCE_LABEL.items():
            sample = summary.loc[summary["reference_group"].eq(family)].sort_values(
                ["main", "iso3"], ascending=[True, True]
            ).reset_index(drop=True)
            if sample.empty:
                continue
            # Identical x limits across modes/pages of a group, with no clipping.
            low, high = min(0., sample["low"].min()), max(0., sample["high"].max())
            margin = max(1., (high - low) * .06)
            page_count = math.ceil(len(sample) / countries_per_page)
            pages = np.array_split(np.arange(len(sample)), page_count)
            for page in range(page_count):
                part = sample.iloc[pages[page]]
                y = np.arange(len(part))
                color = CYCLE_COLOR[cycle]
                for mode in modes:
                    row_height = .45 if mode == "range" else .27
                    fig, ax = plt.subplots(figsize=(10.0, max(4.5, len(part)*row_height + 2.2)))
                    comparable = part["n_scenarios_available"].gt(0).to_numpy()
                    main_y = y.astype(float)
                    if mode == "range":
                        # Six fixed vertical lanes prevent identical estimates from
                        # hiding one another. The horizontal values are unchanged.
                        lanes = np.linspace(-.375, .375, 6)
                        main_y = y + lanes[0]
                        for yy in y[::2]:
                            ax.axhspan(yy-.49, yy+.49, color="#F4F5F7", zorder=0)
                        ax.hlines(y[comparable], part.loc[comparable, "low"],
                                  part.loc[comparable, "high"], color="#BBC0C7", lw=1.3, zorder=1)
                        for lane, scenario in zip(lanes[1:], SCENARIOS):
                            ax.scatter(part[scenario], y + lane, marker=SCENARIO_MARKERS[scenario],
                                       s=18, facecolors="none", edgecolors=SCENARIO_COLORS[scenario],
                                       linewidths=1., zorder=2, label=SCENARIO_LABEL[scenario])
                    else:
                        ax.hlines(y[comparable], part.loc[comparable, "main"],
                                  part.loc[comparable, "farthest_value"], color="#A4ABB3", lw=1.8, zorder=1)
                        ax.scatter(part["farthest_value"], y, s=64, facecolors="none",
                                   edgecolors="#555555", linewidths=1.3, zorder=4,
                                   label="Farthest alternative")
                        for yy, row in zip(y, part.to_dict("records")):
                            if row["farthest_scenario"]:
                                # Fixed right column avoids covering data points.
                                ax.text(1.015, yy, SCENARIO_LABEL[row["farthest_scenario"]],
                                        transform=ax.get_yaxis_transform(), fontsize=7.5, va="center")
                    ax.scatter(part["main"], main_y, s=24, color=color, zorder=5, label="Main estimate")
                    labels = [f"{row['Country']} ({row['iso3']})" +
                              (" †" if row["status"] == "different_years" else "")
                              for row in part.to_dict("records")]
                    ax.set_yticks(y, labels)
                    ax.invert_yaxis()
                    ax.set_ylim(len(part)-.5, -.5)
                    ax.set_xlim(low-margin, high+margin)
                    ax.axvline(0, color="#333333", ls="--", lw=.9)
                    ax.set_xlabel("Country-level median Gap-to-Target (%)")
                    ax.grid(axis="y", visible=False)
                    ax.tick_params(axis="y", length=0, labelsize=8)
                    ax.spines["left"].set_visible(False)
                    handles, names = ax.get_legend_handles_labels()
                    # Put main first, without changing the underlying plotting order.
                    height = fig.get_figheight()
                    fig.legend(handles[-1:]+handles[:-1], names[-1:]+names[:-1],
                              frameon=False, ncol=3 if mode == "range" else 2,
                              loc="center", bbox_to_anchor=(.58, .65/height), fontsize=8)
                    fig.text(.01, .01,
                             "Medians use common years across available specifications; missing scenarios are omitted.\n" +
                             ("Range: main and available alternatives; vertical offsets separate symbols, not values." if mode == "range" else
                              "Open circle: alternative farthest from the main estimate; scenario identified at right.") +
                             ("\n† No shared years among available series: each median uses its available years." if part["status"].eq("different_years").any() else ""),
                             fontsize=7, color="#555555")
                    fig.tight_layout(rect=(0, min(.38, 1.4/height), .91 if mode == "dumbbell" else 1, 1))
                    # save_figure calls tight_layout again; retain the explicit footer margin.
                    fig.set_layout_engine("none")
                    stem = f"Country_{CYCLE_SHORT[cycle]}_{family}_{mode}_p{page+1:02d}"
                    for extension in ("png",):
                        output = folders.appendix / f"{stem}.{extension}"
                        temporary = output.with_name(f".{output.name}.tmp")
                        fig.savefig(temporary, format=extension, bbox_inches="tight")
                        temporary.replace(output)
                    if show:
                        plt.show()
                    plt.close(fig)


# -----------------------------------------------------------------------------
# Main orchestration
# -----------------------------------------------------------------------------
def generate_all(args: argparse.Namespace) -> None:
    configure_style()
    folders = create_output_folders(args.output_dir)
    final, targets, icr = load_inputs(args.final_data, args.targets_data, args.icr_data)
    panel = enrich_panel(final, targets, icr)
    countries = add_graph_categories(country_level(panel))
    write_diagnostics(panel, countries, folders)

    plot_annual_median_uncertainty(
        panel, folders, args.show
    )
    plot_cycle_medians_and_distributions(
        countries, folders, args.show
    )
    plot_alignment_positions_and_transitions(
        countries, folders, args.show
    )
    plot_categorical_medians(
        countries,
        "target_metric_graph",
        TARGET_TYPE_ORDER,
        TARGET_TYPE_LABEL,
        "Median Gap-to-Target by target metric",
        "Figure_4_Gap_by_Target_Type",
        folders.main,
        args.show,
        rotate=28,
    )
    plot_categorical_medians(
        countries,
        "target_scope_std",
        COVERAGE_ORDER,
        COVERAGE_LABEL,
        "Median Gap-to-Target by target coverage",
        "Figure_5_Gap_by_Target_Coverage",
        folders.main,
        args.show,
        rotate=12,
    )
    plot_categorical_medians(
        countries,
        "conditionality_graph",
        CONDITIONALITY_GRAPH_ORDER,
        CONDITIONALITY_GRAPH_LABEL,
        "Median Gap-to-Target by NDC conditionality structure",
        "Figure_6_Gap_by_Conditionality_Structure",
        folders.main,
        args.show,
        rotate=10,
    )
    plot_categorical_medians(
        countries,
        "Incomegroup",
        INCOME_ORDER,
        {x: x for x in INCOME_ORDER},
        "Median Gap-to-Target by income group",
        "Figure_7_Gap_by_Income_Group",
        folders.main,
        args.show,
        rotate=18,
    )
    plot_ambition_alignment(countries, folders, args.show)
    plot_ambition_alignment_transitions(countries, folders, args.show)
    plot_paired_conditionality(countries, folders, args.show)

    plot_categorical_medians(
        countries,
        "Region",
        REGION_ORDER,
        {x: x for x in REGION_ORDER},
        "Median Gap-to-Target by world region",
        "Appendix_Gap_by_Region",
        folders.appendix,
        args.show,
        rotate=38,
    )
    plot_categorical_medians(
        countries,
        "Detailed_region",
        DETAILED_REGION_ORDER,
        {x: x for x in DETAILED_REGION_ORDER},
        "Median Gap-to-Target by detailed geographic region",
        "Appendix_Gap_by_Detailed_Region",
        folders.appendix,
        args.show,
        rotate=38,
        flag_small_samples=True,
    )
    plot_target_architecture(targets, set(final["iso3"].dropna()), folders, args.show)
    plot_scenario_specific_medians(countries, folders, args.show)
    plot_country_scenario_sensitivity(
        panel, folders, args.show, args.country_plot_style, args.countries_per_page
    )
    if not args.no_maps:
        plot_world_maps(
            countries, folders, args.show, args.world_shapefile
        )

    print(f"Figures written to: {folders.root}")
    print(f"Main figures: {folders.main}")
    print(f"Appendix figures: {folders.appendix}")
    print(f"Diagnostics: {folders.diagnostics}")


def main() -> int:
    try:
        args = parse_args()
        generate_all(args)
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
