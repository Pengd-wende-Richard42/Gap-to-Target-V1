"""Generate publication-ready figures for the Kyoto Gap-to-Target indicators.

The default paths are resolved from the repository root and can be overridden
with ``--input`` and ``--output-dir`` when needed.
"""

from __future__ import annotations

import argparse
import urllib.request
import zipfile
from pathlib import Path

import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.colors import TwoSlopeNorm
from matplotlib.lines import Line2D
from matplotlib.patches import Patch


VARIANTS = {
    "gap_basic_pct": "Benchmark",
    "gap_lulucf_pct": "LULUCF-adjusted",
    "gap_trade_pct": "Transaction-adjusted",
    "gap_full_pct": "Fully adjusted",
}

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INPUT = PROJECT_ROOT / "data" / "processed" / "kyoto" / "Gap_Kyoto.dta"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs" / "figures" / "kyoto"
DEFAULT_WORLD_SHAPEFILE = (
    PROJECT_ROOT / "data" / "raw" / "geospatial" /
    "ne_110m_admin_0_countries.zip"
)

NATURAL_EARTH_URL = (
    "https://naturalearth.s3.amazonaws.com/110m_cultural/"
    "ne_110m_admin_0_countries.zip"
)

COLORS = {
    "gap_basic_pct": "#D97706",
    "gap_lulucf_pct": "#6B8E23",
    "gap_trade_pct": "#7B61A8",
    "gap_full_pct": "#1F5A7A",
}

LINE_STYLES = {
    "gap_basic_pct": ("-", "o"),
    "gap_lulucf_pct": ("--", "s"),
    "gap_trade_pct": ("-.", "^"),
    "gap_full_pct": (":", "D"),
}

COUNTRY_NAMES = {
    "AUS": "Australia", "AUT": "Austria", "BEL": "Belgium",
    "BGR": "Bulgaria", "CHE": "Switzerland", "CZE": "Czech Republic",
    "DEU": "Germany", "DNK": "Denmark", "ESP": "Spain",
    "EST": "Estonia", "FIN": "Finland", "FRA": "France",
    "GBR": "United Kingdom", "GRC": "Greece", "HRV": "Croatia",
    "HUN": "Hungary", "IRL": "Ireland", "ISL": "Iceland",
    "ITA": "Italy", "JPN": "Japan", "LIE": "Liechtenstein",
    "LTU": "Lithuania", "LUX": "Luxembourg", "LVA": "Latvia",
    "MCO": "Monaco", "NLD": "Netherlands", "NOR": "Norway",
    "NZL": "New Zealand", "POL": "Poland", "PRT": "Portugal",
    "ROU": "Romania", "RUS": "Russian Federation", "SVK": "Slovakia",
    "SVN": "Slovenia", "SWE": "Sweden", "UKR": "Ukraine",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input", type=Path, default=DEFAULT_INPUT,
        help=f"Input Stata file (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR,
        help=f"Directory in which figures are saved (default: {DEFAULT_OUTPUT_DIR})",
    )
    parser.add_argument(
        "--final-year", type=int, default=2012,
        help="Year used for the country-level dumbbell plot and Kyoto map",
    )
    parser.add_argument(
        "--no-map", action="store_true",
        help="Skip the optional Kyoto world map.",
    )
    parser.add_argument(
        "--world-shapefile", type=Path, default=DEFAULT_WORLD_SHAPEFILE,
        help=(
            "Optional Natural Earth country-boundary .shp or .zip file. When "
            "omitted, the official 110m Admin 0 archive is downloaded and cached."
        ),
    )
    return parser.parse_args()


def configure_style() -> None:
    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "font.size": 10,
        "axes.titlesize": 14,
        "axes.titleweight": "bold",
        "axes.labelsize": 10,
        "figure.facecolor": "white",
        "axes.facecolor": "white",
        "savefig.facecolor": "white",
    })


def validate_data(df: pd.DataFrame) -> None:
    required = {"iso3", "year", *VARIANTS.keys()}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Missing required variables: {sorted(missing)}")
    if df.duplicated(["iso3", "year"]).any():
        raise ValueError("The dataset contains duplicate country-year observations.")


def finish_axis(ax: plt.Axes) -> None:
    ax.grid(axis="y", color="#E5E7EB", linewidth=0.8)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.tick_params(axis="y", length=0)


def save_figure(fig: plt.Figure, output_base: Path) -> None:
    fig.savefig(output_base.with_suffix(".png"), dpi=300, bbox_inches="tight")
    plt.close(fig)


def plot_annual_medians(df: pd.DataFrame, output_dir: Path) -> pd.DataFrame:
    medians = df.groupby("year", as_index=False)[list(VARIANTS)].median()
    fig, ax = plt.subplots(figsize=(8.2, 5.2))

    for variable, label in VARIANTS.items():
        line_style, marker = LINE_STYLES[variable]
        ax.plot(
            medians["year"], medians[variable], marker=marker, linestyle=line_style, linewidth=2,
            markersize=5, color=COLORS[variable], label=label,
        )

    ax.axhline(0, color="#333333", linestyle="--", linewidth=1)
    ax.set_xlabel("Year")
    ax.set_ylabel("Median Gap-to-Target (%)")
    ax.set_xticks(sorted(medians["year"].unique()))
    finish_axis(ax)
    ax.legend(frameon=False, ncol=2, loc="best")
    fig.text(
        0.01, 0.01,
        "Note: Medians are computed across countries. Negative values indicate "
        "emissions below the annualised Kyoto benchmark.",
        fontsize=8.5, color="#555555",
    )
    fig.subplots_adjust(bottom=0.16)
    save_figure(fig, output_dir / "kyoto_annual_medians")
    medians.to_csv(output_dir / "kyoto_annual_medians.csv", index=False)
    return medians


def plot_distributions(df: pd.DataFrame, output_dir: Path) -> None:
    values = [df[column].dropna().to_numpy() for column in VARIANTS]
    labels = list(VARIANTS.values())
    fig, ax = plt.subplots(figsize=(8.2, 5.4))

    box = ax.boxplot(
        values, tick_labels=labels, patch_artist=True, widths=0.58,
        showfliers=True,
        medianprops={"color": "#111827", "linewidth": 1.6},
        whiskerprops={"color": "#6B7280"},
        capprops={"color": "#6B7280"},
        flierprops={
            "marker": "o", "markersize": 3, "markerfacecolor": "#9CA3AF",
            "markeredgecolor": "none", "alpha": 0.6,
        },
    )
    for patch, color in zip(box["boxes"], COLORS.values()):
        patch.set_facecolor(color)
        patch.set_alpha(0.78)
        patch.set_edgecolor("white")

    ax.axhline(0, color="#333333", linestyle="--", linewidth=1)
    ax.set_ylabel("Gap-to-Target (%)")
    finish_axis(ax)
    ax.tick_params(axis="x", rotation=12)
    fig.text(
        0.01, 0.01,
        "Note: Boxes show the interquartile range; horizontal lines inside boxes "
        "show the median.",
        fontsize=8.5, color="#555555",
    )
    fig.subplots_adjust(bottom=0.20)
    save_figure(fig, output_dir / "kyoto_variant_distributions")


def plot_dumbbell(df: pd.DataFrame, output_dir: Path, year: int) -> None:
    plot = (
        df.loc[df["year"].eq(year), ["iso3", "gap_basic_pct", "gap_full_pct"]]
        .dropna()
        .sort_values("gap_full_pct", ascending=True)
        .reset_index(drop=True)
    )
    if plot.empty:
        raise ValueError(f"No complete benchmark/fully adjusted observations for {year}.")

    y = np.arange(len(plot))
    height = max(7.0, 0.30 * len(plot) + 1.8)
    fig, ax = plt.subplots(figsize=(9.2, height))

    for row, ypos in zip(plot.itertuples(index=False), y):
        ax.plot(
            [row.gap_basic_pct, row.gap_full_pct], [ypos, ypos],
            color="#B9C1C8", linewidth=1.5, zorder=1,
        )

    ax.scatter(
        plot["gap_basic_pct"], y, s=68, facecolors="none",
        edgecolor=COLORS["gap_basic_pct"], linewidth=1.5, zorder=4,
    )
    ax.scatter(
        plot["gap_full_pct"], y, s=38, color=COLORS["gap_full_pct"],
        edgecolor="white", linewidth=0.5, zorder=3,
    )
    ax.axvline(0, color="#333333", linewidth=1, linestyle="--")
    ax.set_yticks(y, plot["iso3"].map(COUNTRY_NAMES).fillna(plot["iso3"]))
    ax.set_xlabel("Gap-to-Target (%)")
    ax.grid(axis="x", color="#E5E7EB", linewidth=0.8)
    ax.grid(axis="y", visible=False)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.tick_params(axis="y", length=0, labelsize=8.5)
    ax.legend(
        handles=[
            Line2D([0], [0], marker="o", color="none",
                   markerfacecolor="none", markeredgecolor=COLORS["gap_basic_pct"],
                   markeredgewidth=1.5, markersize=8,
                   label="Benchmark"),
            Line2D([0], [0], marker="o", color="none",
                   markerfacecolor=COLORS["gap_full_pct"],
                   markeredgecolor="white", markersize=7,
                   label="Fully adjusted"),
        ],
        frameon=False, ncol=2, loc="lower center", bbox_to_anchor=(0.5, -0.075),
    )
    fig.text(
        0.01, 0.01,
        "Note: Countries are ranked by the fully adjusted Gap. Negative values "
        "indicate emissions below the annualised Kyoto benchmark.",
        fontsize=8.5, color="#555555",
    )
    fig.subplots_adjust(left=0.23, right=0.98, bottom=0.11)
    save_figure(fig, output_dir / f"kyoto_dumbbell_{year}")


def clean_iso(series: pd.Series) -> pd.Series:
    return series.astype("string").str.strip().str.upper()


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
    cache_dir.mkdir(parents=True, exist_ok=True)
    extracted = cache_dir / "ne_110m_admin_0_countries"
    existing = sorted(extracted.glob("*admin_0_countries.shp"))
    if existing:
        return existing[0]

    archive = cache_dir / "ne_110m_admin_0_countries.zip"
    if not archive.exists():
        print("Downloading Natural Earth country boundaries for the Kyoto map...")
        try:
            urllib.request.urlretrieve(NATURAL_EARTH_URL, archive)
        except Exception as exc:
            raise RuntimeError(
                "The Natural Earth boundary archive could not be downloaded. "
                "Download ne_110m_admin_0_countries.zip manually and pass its "
                "path with --world-shapefile."
            ) from exc
    return extract_natural_earth_archive(archive, extracted)


def load_world_boundaries(world_shapefile: Path | None, cache_dir: Path):
    try:
        import geopandas as gpd
    except ImportError as exc:
        raise RuntimeError(
            "The Kyoto map requires GeoPandas. Install it once with: "
            "python -m pip install geopandas pyogrio shapely"
        ) from exc

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


def plot_kyoto_map(
    df: pd.DataFrame,
    output_dir: Path,
    year: int,
    world_shapefile: Path | None,
) -> None:
    """Map the fully adjusted Kyoto Gap in the final CP1 year."""
    world = load_world_boundaries(
        world_shapefile,
        PROJECT_ROOT / "data" / "processed" / "geospatial" / "Natural_Earth_cache",
    ).copy()
    world["map_iso3"] = map_iso3_column(world)

    values = (
        df.loc[df["year"].eq(year), ["iso3", "gap_full_pct"]]
        .dropna(subset=["gap_full_pct"])
        .copy()
    )
    values["iso3"] = clean_iso(values["iso3"])
    values = values.rename(columns={"gap_full_pct": "gap"})
    if values.empty:
        raise ValueError(f"No fully adjusted Kyoto Gap is available for {year}.")

    merged = world.merge(values, left_on="map_iso3", right_on="iso3", how="left")
    available_map = merged.loc[merged["gap"].notna()]
    missing_map = merged.loc[merged["gap"].isna()]

    bound = max(float(values["gap"].abs().quantile(0.95)), 1.0)
    norm = TwoSlopeNorm(vmin=-bound, vcenter=0.0, vmax=bound)
    fig, ax = plt.subplots(figsize=(11.0, 5.4))
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
    ax.set_axis_off()

    scalar_map = mpl.cm.ScalarMappable(norm=norm, cmap="RdBu_r")
    scalar_map.set_array([])
    colorbar = fig.colorbar(
        scalar_map,
        ax=ax,
        orientation="horizontal",
        fraction=0.055,
        pad=0.035,
        aspect=42,
        extend="both",
    )
    colorbar.set_label("Fully adjusted Gap-to-Target (%)")
    missing_legend = Patch(
        facecolor="#D9D9D9",
        edgecolor="#777777",
        hatch="///",
        label="No Kyoto Gap estimate",
    )
    fig.legend(
        handles=[missing_legend],
        frameon=True,
        loc="lower left",
        bbox_to_anchor=(0.02, 0.02),
        fontsize=8.5,
    )
    fig.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.17)
    fig.set_layout_engine("none")
    save_figure(fig, output_dir / f"kyoto_map_fully_adjusted_{year}")


def main() -> None:
    args = parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    configure_style()

    df = pd.read_stata(args.input)
    validate_data(df)
    plot_annual_medians(df, args.output_dir)
    plot_distributions(df, args.output_dir)
    plot_dumbbell(df, args.output_dir, args.final_year)
    if not args.no_map:
        plot_kyoto_map(
            df, args.output_dir, args.final_year, args.world_shapefile
        )

    print(f"Figures saved in: {args.output_dir.resolve()}")


if __name__ == "__main__":
    main()
