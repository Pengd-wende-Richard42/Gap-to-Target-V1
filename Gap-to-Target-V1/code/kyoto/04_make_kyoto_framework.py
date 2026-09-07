"""Draw the common Gap-to-Target research framework; export PNG only."""

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch


PROJECT_ROOT = Path(__file__).resolve().parents[2]
OUTPUT_DIR = PROJECT_ROOT / "outputs" / "figures" / "kyoto"
OUTPUT_NAME = "gap_to_target_common_research_framework.png"
COLORS = {
    "ink": "#1F2937",
    "muted": "#64748B",
    "border": "#CBD5E1",
    "kyoto": "#DCEAF7",
    "paris": "#FCE8D5",
    "emissions": "#E0F2EC",
    "process": "#EEE8FA",
    "path": "#F1F5F9",
    "gap": "#FDE2E4",
    "complement": "#FFF3CD",
}


def add_text(ax, x, y, value, size=9, bold=False, align="center", color="ink"):
    return ax.text(
        x, y, value, transform=ax.transAxes, ha=align, va="center",
        fontsize=size, fontweight="bold" if bold else "normal",
        color=COLORS[color],
    )


def add_box(ax, x, y, width, height, color, title):
    ax.add_patch(FancyBboxPatch(
        (x, y), width, height, transform=ax.transAxes,
        boxstyle="round,pad=0.009,rounding_size=0.018",
        linewidth=1, edgecolor=COLORS["border"], facecolor=COLORS[color],
    ))
    add_text(ax, x + width / 2, y + height - 0.032, title, 11, True)


def add_arrow(ax, start, end, connectionstyle="arc3"):
    ax.add_patch(FancyArrowPatch(
        start, end, transform=ax.transAxes, arrowstyle="-|>",
        mutation_scale=13, linewidth=1.2, color=COLORS["muted"],
        connectionstyle=connectionstyle, shrinkA=3, shrinkB=3,
    ))


def section_label(ax, y, value):
    add_text(ax, 0.035, y, value, 8, True, "left", "muted")


def build_figure():
    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "axes.unicode_minus": True,
        "pdf.fonttype": 42,
        "ps.fonttype": 42,
    })
    fig, ax = plt.subplots(figsize=(12.5, 8.0), facecolor="white")
    ax.set(xlim=(0, 1), ylim=(0, 1))
    ax.axis("off")

    section_label(ax, 0.963, "AGREEMENT-SPECIFIC INPUTS")
    add_box(ax, 0.03, 0.765, 0.29, 0.17, "kyoto", "Kyoto Protocol")
    add_text(ax, 0.175, 0.862, "Assigned amount for CP1")
    add_text(ax, 0.175, 0.826, "Annex A emissions and LULUCF")
    add_text(ax, 0.175, 0.790, "Transactions in Kyoto units")

    add_box(ax, 0.355, 0.765, 0.29, 0.17, "emissions", "Observed emissions")
    add_text(ax, 0.500, 0.862, "Annual country-level emissions")
    add_text(ax, 0.500, 0.826, "Matched to the declared target scope")

    add_box(ax, 0.68, 0.765, 0.29, 0.17, "paris", "Paris Agreement")
    add_text(ax, 0.825, 0.862, "First and second NDC cycles")
    add_text(ax, 0.825, 0.826, "Heterogeneous target architectures")
    add_text(ax, 0.825, 0.790, "Prospective inputs where required (BAU)")

    add_arrow(ax, (0.175, 0.758), (0.300, 0.685), "angle3,angleA=-90,angleB=180")
    add_arrow(ax, (0.455, 0.758), (0.405, 0.675), "arc3,rad=0.12")
    add_arrow(ax, (0.545, 0.758), (0.595, 0.675), "arc3,rad=-0.12")
    add_arrow(ax, (0.825, 0.758), (0.700, 0.685), "angle3,angleA=-90,angleB=0")

    section_label(ax, 0.700, "AGREEMENT-SPECIFIC HARMONISATION")
    add_box(ax, 0.07, 0.505, 0.40, 0.16, "process", "Kyoto accounting")
    add_text(ax, 0.270, 0.600, "Annualise the five-year assigned amount")
    add_text(ax, 0.270, 0.561, "Apply LULUCF and unit-transaction variants")

    add_box(ax, 0.53, 0.505, 0.40, 0.16, "process", "Paris target conversion")
    add_text(ax, 0.730, 0.600, "Match sector, gas and forest coverage")
    add_text(ax, 0.730, 0.561, "Convert each NDC into a target-year outcome")

    add_arrow(ax, (0.270, 0.497), (0.425, 0.420), "angle3,angleA=-90,angleB=180")
    add_arrow(ax, (0.730, 0.497), (0.575, 0.420), "angle3,angleA=-90,angleB=0")

    section_label(ax, 0.435, "ANNUAL BENCHMARK")
    add_box(ax, 0.16, 0.295, 0.68, 0.11, "path", "Target-consistent country-year benchmark")
    add_text(ax, 0.500, 0.348, "Translate each commitment into the emissions benchmark applicable in year $t$")
    add_text(ax, 0.500, 0.313, "Agreement-specific rules are preserved within annual structure", 8, color="muted")

    add_arrow(ax, (0.500, 0.287), (0.500, 0.238))
    section_label(ax, 0.249, "OUTPUTS")
    add_box(ax, 0.07, 0.065, 0.58, 0.16, "gap", "Annual Gap-to-Target")
    add_text(ax, 0.360, 0.132, r"$Gap^k_{i,t}=100\times\dfrac{E^{obs,k}_{i,t}-B^k_{i,t}}{B^k_{i,t}}$", 14)
    add_text(ax, 0.360, 0.082, "Negative: below benchmark  ·  Zero: aligned  ·  Positive: above benchmark", 8, color="muted")

    add_box(ax, 0.69, 0.065, 0.24, 0.16, "complement", "Complementary analyses")
    add_text(ax, 0.810, 0.151, "Commitment ambition")
    add_text(ax, 0.810, 0.118, "Accounting and scenario sensitivity")
    add_text(ax, 0.810, 0.085, "Country-level audit", 8, color="muted")

    fig.subplots_adjust(left=0.015, right=0.985, bottom=0.02, top=0.99)
    return fig


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--dpi", type=int, default=600)
    args = parser.parse_args()
    if args.dpi < 1:
        parser.error("--dpi must be positive")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    fig = build_figure()
    fig.savefig(output, dpi=args.dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"Created: {output}")


if __name__ == "__main__":
    main()
