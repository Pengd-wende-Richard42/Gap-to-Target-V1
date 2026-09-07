"""Create the Paris research-framework figure; export PNG only.

The figure separates commitment ambition (descriptive) from the annual
Gap-to-Target (core indicator). It deliberately excludes the unstable
implementation-progress scores.

The output is a 600 dpi PNG.
"""

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch


PROJECT_ROOT = Path(__file__).resolve().parents[2]
OUTPUT_DIR = PROJECT_ROOT / "outputs" / "figures" / "paris"
OUTPUT_STEM = "gap_to_target_research_framework"


COLORS = {
    "ink": "#1F2937",
    "muted": "#64748B",
    "border": "#CBD5E1",
    "ndc": "#DCEAF7",
    "emissions": "#E0F2EC",
    "scenarios": "#FCE8D5",
    "process": "#EEE8FA",
    "path": "#F1F5F9",
    "ambition": "#FFF3CD",
    "gap": "#FDE2E4",
    "white": "#FFFFFF",
}


def add_box(ax, x, y, width, height, facecolor, title, lines=(), title_size=10.0):
    """Draw a rounded methodological block in axis-relative coordinates."""
    patch = FancyBboxPatch(
        (x, y), width, height,
        boxstyle="round,pad=0.012,rounding_size=0.015",
        linewidth=1.0,
        edgecolor=COLORS["border"],
        facecolor=facecolor,
        transform=ax.transAxes,
        clip_on=False,
    )
    ax.add_patch(patch)
    ax.text(
        x + width / 2, y + height - 0.028, title,
        transform=ax.transAxes, ha="center", va="top",
        fontsize=title_size, fontweight="bold", color=COLORS["ink"],
    )
    line_y = y + height - 0.066
    for line in lines:
        ax.text(
            x + 0.018, line_y, line,
            transform=ax.transAxes, ha="left", va="top",
            fontsize=8.25, color=COLORS["ink"],
        )
        line_y -= 0.027
    return patch


def add_arrow(ax, start, end, connectionstyle="arc3"):
    """Add a clean directional arrow between two axis-relative points."""
    arrow = FancyArrowPatch(
        start, end,
        transform=ax.transAxes,
        arrowstyle="-|>", mutation_scale=12,
        linewidth=1.1, color=COLORS["muted"],
        connectionstyle=connectionstyle,
        shrinkA=2, shrinkB=2,
    )
    ax.add_patch(arrow)


def section_label(ax, x, y, text):
    ax.text(
        x, y, text,
        transform=ax.transAxes, ha="left", va="bottom",
        fontsize=7.6, fontweight="bold", color=COLORS["muted"],
    )


def build_figure():
    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "axes.unicode_minus": True,
        "pdf.fonttype": 42,
        "ps.fonttype": 42,
    })

    fig, ax = plt.subplots(figsize=(12.5, 9.2), facecolor=COLORS["white"])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    # ------------------------------------------------------------
    # 1. Inputs
    # ------------------------------------------------------------
    section_label(ax, 0.03, 0.955, "INPUTS")

    add_box(
        ax, 0.03, 0.78, 0.28, 0.16, COLORS["ndc"],
        "National commitments",
        [
            "• First and second NDC cycles",
            "• BAU, base-year, intensity, per-capita",
            "  and absolute/fixed-level targets",
            "• Main, unconditional and conditional",
        ],
    )
    add_box(
        ax, 0.36, 0.78, 0.28, 0.16, COLORS["emissions"],
        "Observed emissions",
        [
            "• EDGAR IPCC 2006 sector emissions",
            "• FAOSTAT Forest Land where net forest",
            "  emissions are covered by the target",
            "• Annual observations, 2015–2024",
        ],
    )
    add_box(
        ax, 0.69, 0.78, 0.28, 0.16, COLORS["scenarios"],
        "Prospective inputs",
        [
            "• Official NDC BAU when available",
            "• Projected emissions, GDP and population",
            "  projections (Gütschow et al., 2021)",
            "• RCP-SSP and model alternatives",
        ],
    )

    # Inputs converge on the harmonisation stage.
    add_arrow(ax, (0.17, 0.78), (0.44, 0.735), "angle3,angleA=-90,angleB=180")
    add_arrow(ax, (0.50, 0.78), (0.50, 0.735))
    add_arrow(ax, (0.83, 0.78), (0.56, 0.735), "angle3,angleA=-90,angleB=0")

    # ------------------------------------------------------------
    # 2. Harmonisation
    # ------------------------------------------------------------
    section_label(ax, 0.03, 0.725, "HARMONISATION")
    add_box(
        ax, 0.15, 0.555, 0.70, 0.155, COLORS["process"],
        "Scope-consistent conversion to a common emissions target",
        [
            "1  Match sector, gas and forest coverage",
            r"2  Convert each target architecture into $E^{target}_{ic}$",
            "3  Keep First NDC and Second NDC cycles separate",
            "4  Document and audit country-level coding decisions",
        ],
        title_size=10.4,
    )
    add_arrow(ax, (0.50, 0.555), (0.50, 0.505))

    # ------------------------------------------------------------
    # 3. Annual target-consistent pathway
    # ------------------------------------------------------------
    section_label(ax, 0.03, 0.515, "TARGET-CONSISTENT ANNUAL PATH")
    add_box(
        ax, 0.08, 0.345, 0.84, 0.155, COLORS["path"],
        "From the pre-implementation anchor to the NDC target year",
        [], title_size=10.4,
    )

    # Core annual-path equation, explicitly labelled above the timeline.
    ax.text(
        0.50, 0.448, "Annualisation formula",
        transform=ax.transAxes, ha="center", va="center",
        fontsize=7.7, fontweight="bold", color=COLORS["muted"],
    )
    ax.text(
        0.50, 0.423,
        r"$E^{path}_{i,t,c}=E_{i,s_{i,c}-1}+"
        r"\dfrac{t-(s_{i,c}-1)}{T_{i,c}-(s_{i,c}-1)}"
        r"\left(E^{target}_{i,c}-E_{i,s_{i,c}-1}\right)$",
        transform=ax.transAxes, ha="center", va="center",
        fontsize=8.9, color=COLORS["ink"],
    )

    y_timeline = 0.375
    x_points = [0.20, 0.50, 0.80]
    ax.plot(
        [x_points[0], x_points[-1]], [y_timeline, y_timeline],
        transform=ax.transAxes, color=COLORS["ink"], linewidth=1.6,
    )
    for x in x_points:
        ax.plot(
            x, y_timeline, marker="o", markersize=6,
            markerfacecolor=COLORS["white"], markeredgecolor=COLORS["ink"],
            markeredgewidth=1.4, transform=ax.transAxes,
        )

    timeline_labels = [
        (0.20, r"$s_{i,c}-1$", r"$E_{i,s_{i,c}-1}$"),
        (0.50, r"year $t$", r"$E^{path}_{i,t,c}$"),
        (0.80, r"target year $T_{i,c}$", r"$E^{target}_{i,c}$"),
    ]
    for x, upper, lower in timeline_labels:
        ax.text(
            x, y_timeline + 0.018, upper,
            transform=ax.transAxes, ha="center", va="bottom",
            fontsize=8.4, color=COLORS["muted"],
        )
        ax.text(
            x, y_timeline - 0.021, lower,
            transform=ax.transAxes, ha="center", va="top",
            fontsize=10.0, fontweight="bold", color=COLORS["ink"],
        )

    add_arrow(ax, (0.39, 0.345), (0.29, 0.315))
    add_arrow(ax, (0.61, 0.345), (0.71, 0.315))

    # ------------------------------------------------------------
    # 4. Core outputs — no implementation-progress score
    # ------------------------------------------------------------
    section_label(ax, 0.03, 0.305, "CORE OUTPUTS")
    add_box(
        ax, 0.08, 0.145, 0.40, 0.145, COLORS["ambition"],
        "Commitment ambition", [], title_size=10.1,
    )
    ax.text(
        0.28, 0.215,
        r"$A_{ic}=100\times\dfrac{E_{ib}-E^{target}_{ic}}{E_{ib}}$",
        transform=ax.transAxes, ha="center", va="center",
        fontsize=11.4, color=COLORS["ink"],
    )
    ax.text(
        0.28, 0.169, "Stringency of the pledged emissions reduction",
        transform=ax.transAxes, ha="center", va="center",
        fontsize=8.1, color=COLORS["muted"],
    )

    add_box(
        ax, 0.52, 0.145, 0.45, 0.145, COLORS["gap"],
        "Annual Gap-to-Target Indicator", [], title_size=10.1,
    )
    ax.text(
        0.72, 0.215,
        r"$Gap_{itc}=100\times\dfrac{E^{obs}_{itc}-E^{path}_{itc}}{E^{path}_{itc}}$",
        transform=ax.transAxes, ha="center", va="center",
        fontsize=11.0, color=COLORS["ink"],
    )
    ax.text(
        0.745, 0.169, "Negative: below the pathway  ·  Zero: aligned with the pathway  ·  Positive: above the pathway",
        transform=ax.transAxes, ha="center", va="center",
        fontsize=8.1, color=COLORS["muted"],
    )

    # ------------------------------------------------------------
    # 5. Uncertainty layer
    # ------------------------------------------------------------
    ax.text(
        0.50, 0.075,
        "UNCERTAINTY ANALYSIS",
        transform=ax.transAxes, ha="center", va="center",
        fontsize=7.6, fontweight="bold", color=COLORS["muted"],
    )
    ax.text(
        0.50, 0.040,
        "Targets requiring prospective inputs: repeat conversion using available RCP4.5-SSP2, "
        "RCP8.5-SSP2, OECD, PIK and IIASA alternatives; unavailable scenarios are omitted.",
        transform=ax.transAxes, ha="center", va="center",
        fontsize=8.0, color=COLORS["ink"],
    )

    fig.subplots_adjust(left=0.02, right=0.98, top=0.985, bottom=0.02)
    return fig


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    fig = build_figure()

    png_path = OUTPUT_DIR / f"{OUTPUT_STEM}.png"

    fig.savefig(png_path, dpi=600, bbox_inches="tight", facecolor=COLORS["white"])
    plt.close(fig)

    
    print(f"Created: {png_path}")


if __name__ == "__main__":
    main()
