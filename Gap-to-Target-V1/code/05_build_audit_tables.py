#!/usr/bin/env python3
"""Build country-level audit tables for the Paris and Kyoto applications.

The script reads the outputs created by the replication pipeline and creates:

1. a complete country-cycle Paris audit (CSV and XLSX);
2. a compact landscape longtable for the paper's Appendix;
3. a CP1 Kyoto country audit (CSV and XLSX);
4. a compact Kyoto longtable;
5. a short metadata and consistency report.

Example
-------
python code/05_build_audit_tables.py \
    --appendix-scope final \
    --edgar-release "[insert exact EDGAR release]" \
    --edgar-access-date "[insert access date]"
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import numpy as np
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PARIS_DIR = PROJECT_ROOT / "data" / "processed" / "paris"
KYOTO_DIR = PROJECT_ROOT / "data" / "processed" / "kyoto"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs" / "audit"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate reproducible country-level audit tables."
    )
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR, type=Path)
    parser.add_argument(
        "--appendix-scope",
        choices=("final", "all"),
        default="final",
        help=(
            "Rows included in the LaTeX Paris table. 'final' retains countries "
            "with at least one cycle-specific Gap; 'all' reports the full universe. "
            "CSV/XLSX outputs always contain the full universe."
        ),
    )
    parser.add_argument(
        "--edgar-release",
        default="EDGAR GHG 2025 release (1970--2024)",
        help="Exact EDGAR release/vintage reported in the manuscript.",
    )
    parser.add_argument(
        "--edgar-access-date",
        default="2026",
        help="EDGAR access date reported in the manuscript.",
    )
    return parser.parse_args()


def clean_text(value: object) -> object:
    if pd.isna(value):
        return np.nan
    text = re.sub(r"\s+", " ", str(value).strip())
    return text if text else np.nan


def first_nonmissing(*values: object) -> object:
    for value in values:
        if not pd.isna(value) and str(value).strip():
            return clean_text(value)
    return np.nan


def year_text(value: object) -> str:
    if pd.isna(value):
        return "--"
    try:
        return str(int(float(value)))
    except (TypeError, ValueError):
        return str(value)


def code_labels(series: pd.Series, mapping: dict[str, str]) -> pd.Series:
    cleaned = series.map(clean_text)
    return cleaned.map(lambda x: mapping.get(x, x) if not pd.isna(x) else np.nan)


def latex_escape(value: object) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return "--"
    text = str(value)
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    return "".join(replacements.get(char, char) for char in text)


def human_label(value: object) -> str:
    """Convert internal machine codes into compact, readable audit labels."""
    if pd.isna(value):
        return "--"
    text = re.sub(r"_+", " ", str(value).strip())
    text = re.sub(r"\s+", " ", text)
    return text.capitalize()


def gap_availability(final: pd.DataFrame, gap_column: str, cycle: str) -> pd.DataFrame:
    available = final.loc[final[gap_column].notna(), ["iso3", "year"]].copy()
    if available.empty:
        return pd.DataFrame(
            columns=["iso3", "NDCs_number", "gap_available", "first_gap_year", "last_gap_year", "n_gap_years"]
        )
    summary = (
        available.groupby("iso3", as_index=False)
        .agg(first_gap_year=("year", "min"), last_gap_year=("year", "max"), n_gap_years=("year", "nunique"))
    )
    summary["NDCs_number"] = cycle
    summary["gap_available"] = 1
    return summary


def build_paris_audit(targets: pd.DataFrame, final: pd.DataFrame) -> pd.DataFrame:
    targets = targets.copy()
    targets.columns = [str(column).strip() for column in targets.columns]
    targets["iso3"] = targets["iso3"].astype("string").str.strip().str.upper()
    targets["Country"] = targets["Country"].map(clean_text)
    # The source contains an EU-27 aggregate in addition to sovereign countries.
    # It has no ISO3 code and is not part of the 197-country universe.
    targets = targets.loc[targets["iso3"].notna()].copy()

    if targets.duplicated(["iso3", "NDCs_number"]).any():
        duplicates = targets.loc[
            targets.duplicated(["iso3", "NDCs_number"], keep=False),
            ["iso3", "NDCs_number"],
        ]
        raise ValueError(f"Duplicate country-cycle target rows found:\n{duplicates}")

    final = final.copy()
    final["iso3"] = final["iso3"].astype("string").str.strip().str.upper()
    availability = pd.concat(
        [
            gap_availability(final, "gap_n1", "First_NDC"),
            gap_availability(final, "gap_n2", "Second_NDC"),
        ],
        ignore_index=True,
    )

    audit = targets.merge(
        availability,
        how="left",
        on=["iso3", "NDCs_number"],
        validate="one_to_one",
    )
    audit["gap_available"] = audit["gap_available"].fillna(0).astype(int)
    audit["cycle"] = audit["NDCs_number"].map(
        {"First_NDC": "NDC1", "Second_NDC": "NDC2"}
    )

    metric_codes = {
        "BASE_YEAR_REDUCTION": "BY",
        "BAU_REDUCTION": "BAU",
        "INTENSITY_TARGET": "INT",
        "PER_CAPITA_TARGET": "PC",
        "BAU_PER_CAPITA_TARGET": "BAU-PC",
        "FIXED_LEVEL_TARGET": "FL",
        "TRAJECTORY_TARGET": "TRJ",
        "NO_QUANTIFIED_GHG_TARGET": "NQ",
        "OTHER_OR_UNKNOWN": "UNK",
    }
    conditionality_codes = {
        "UNCONDITIONAL_ONLY": "U",
        "CONDITIONAL_ONLY": "C",
        "MIXED_CONDITIONALITY": "U+C",
        "PARTIALLY_CONDITIONAL": "PC",
        "NO_QUANTIFIED_GHG_TARGET": "NQ",
        "NO_SUBMISSION": "NS",
        "UNKNOWN": "UNK",
    }
    # Use the final harmonised scope rather than the preliminary coverage field.
    # This prevents manually resolved observations from remaining labelled "Unclear".
    coverage_codes = {
        "ECONOMY_WIDE_OR_BROAD": "Economy-wide/broad",
        "MULTISECTORAL": "Multisectoral",
        "SECTORAL": "Sectoral",
        "UNKNOWN": "Not specified",
        "NO_QUANTIFIED_GHG_TARGET": "Not applicable",
        "NO_SUBMISSION": "Not applicable",
    }
    audit["metric_code"] = code_labels(audit["target_metric_std"], metric_codes)
    audit["conditionality_code"] = code_labels(
        audit["conditionality_std"], conditionality_codes
    )
    audit["coverage_code"] = code_labels(audit["target_scope_std"], coverage_codes)
    audit["coverage_code"] = audit["coverage_code"].fillna("Not specified")

    def implementation_period(row: pd.Series) -> str:
        start = year_text(row.get("start_year"))
        target = year_text(row.get("target_year"))
        if start == "--" and target == "--":
            return "--"
        if start == "--":
            return f"?--{target}"
        if target == "--":
            return f"{start}--?"
        short_target = target[-2:] if len(target) == 4 and target[:2] == start[:2] else target
        return f"{start}--{short_target}"

    audit["implementation_period"] = audit.apply(implementation_period, axis=1)
    audit["forest_rule"] = np.where(
        pd.to_numeric(audit["fao_forest_proxy_flag"], errors="coerce").fillna(0).eq(1),
        "FAO added",
        "Not added",
    )
    audit["prospective_input"] = audit.apply(
        lambda row: first_nonmissing(row.get("bau_source"), row.get("target_method"), "None"),
        axis=1,
    )
    audit.loc[
        ~audit["metric_code"].isin(["BAU", "BAU-PC", "INT", "PC"]),
        "prospective_input",
    ] = "None"

    audit["final_decision"] = np.where(
        audit["gap_available"].eq(1), "Constructed", "Not constructed"
    )
    audit["decision_reason"] = audit.apply(
        lambda row: first_nonmissing(
            row.get("audit_reason"),
            row.get("target_note"),
            row.get("target_status"),
            row.get("target_usability_status"),
        ),
        axis=1,
    )

    keep = [
        "Country",
        "iso3",
        "cycle",
        "NDCs_number",
        "submission_year",
        "target_metric_std",
        "metric_code",
        "base_year",
        "target_year",
        "start_year",
        "implementation_period",
        "conditionality_std",
        "conditionality_code",
        "main_target_rule_std",
        "main_target_pct_std",
        "unconditional_pct_std",
        "conditional_total_pct_std",
        "target_scope_std",
        "quantified_target_coverage_std",
        "coverage_code",
        "scope_sector_text_std",
        "emissions_scope_method_std",
        "edgar_sector_groups_std",
        "sector_mapping_text_std",
        "forest_rule",
        "fao_forest_proxy_flag",
        "prospective_input",
        "bau_source",
        "target_method",
        "harmonization_method",
        "target_usability_status",
        "audit_decision",
        "decision_reason",
        "final_decision",
        "gap_available",
        "first_gap_year",
        "last_gap_year",
        "n_gap_years",
        "coding_quality_flag",
        "coding_note",
        "audit_schema_version",
    ]
    keep = [column for column in keep if column in audit.columns]
    return audit[keep].sort_values(["iso3", "cycle"]).reset_index(drop=True)


def build_kyoto_audit(workbook: pd.DataFrame, final: pd.DataFrame) -> pd.DataFrame:
    audit = workbook.copy()
    audit["iso3"] = audit["iso3"].astype("string").str.strip().str.upper()
    audit = audit.loc[pd.to_numeric(audit["commitment_period"], errors="coerce").eq(1)].copy()
    # The replication panel contains the same audited quantities under concise
    # analytical names. Reconstruct the three availability fields previously
    # stored only in the stand-alone audit workbook.
    audit["official_AA_available"] = audit["assigned_amount_t"].notna().astype(int)
    audit["lulucf_net_raw_t"] = audit["lulucf_net_t"]
    audit["net_acquisition_raw_units"] = audit["net_acquisition_units"]
    final = final.loc[final["year"].between(2008, 2012)].copy()
    final["iso3"] = final["iso3"].astype("string").str.strip().str.upper()

    def all_available(series: pd.Series) -> int:
        values = pd.to_numeric(series, errors="coerce")
        return int(values.notna().all() and len(values) == 5)

    grouped = (
        audit.groupby(["iso3", "party"], as_index=False)
        .agg(
            first_year=("year", "min"),
            last_year=("year", "max"),
            n_years=("year", "nunique"),
            assigned_amount_t=("assigned_amount_t", "first"),
            annualized_assigned_amount_t=("annualized_assigned_amount_t", "first"),
            official_AA_available=("official_AA_available", "max"),
            annexA_complete=("annexA_emissions_t", all_available),
            lulucf_years=("lulucf_net_raw_t", lambda x: int(pd.to_numeric(x, errors="coerce").notna().sum())),
            trade_years=("net_acquisition_raw_units", lambda x: int(pd.to_numeric(x, errors="coerce").notna().sum())),
            gap_basic_years=("gap_basic_pct", lambda x: int(pd.to_numeric(x, errors="coerce").notna().sum())),
            gap_lulucf_years=("gap_lulucf_pct", lambda x: int(pd.to_numeric(x, errors="coerce").notna().sum())),
            gap_trade_years=("gap_trade_pct", lambda x: int(pd.to_numeric(x, errors="coerce").notna().sum())),
            gap_full_years=("gap_full_pct", lambda x: int(pd.to_numeric(x, errors="coerce").notna().sum())),
            annexA_source=("annexA_emissions_t_source_category", "first"),
            trade_source=("source_flow_type", lambda x: "; ".join(sorted(set(x.dropna().astype(str))))),
        )
    )
    final_status = (
        final.groupby("iso3", as_index=False)
        .agg(
            final_basic_years=("gap_basic_pct", lambda x: int(x.notna().sum())),
            final_lulucf_years=("gap_lulucf_pct", lambda x: int(x.notna().sum())),
            final_trade_years=("gap_trade_pct", lambda x: int(x.notna().sum())),
            final_full_years=("gap_full_pct", lambda x: int(x.notna().sum())),
        )
    )
    grouped = grouped.merge(final_status, on="iso3", how="left", validate="one_to_one")
    for column in ["final_basic_years", "final_lulucf_years", "final_trade_years", "final_full_years"]:
        grouped[column] = grouped[column].fillna(0).astype(int)
    grouped["final_decision"] = np.where(
        grouped["final_basic_years"].gt(0), "Retained", "Excluded"
    )
    grouped["accounting_variants"] = grouped.apply(
        lambda row: "; ".join(
            label
            for label, column in [
                ("B", "final_basic_years"),
                ("L", "final_lulucf_years"),
                ("T", "final_trade_years"),
                ("F", "final_full_years"),
            ]
            if row[column] > 0
        )
        or "None",
        axis=1,
    )
    return grouped.sort_values("iso3").reset_index(drop=True)


def render_longtable(
    data: pd.DataFrame,
    columns: list[tuple[str, str, str]],
    caption: str,
    label: str,
    notes: str,
) -> str:
    colspec = " ".join(spec for _, _, spec in columns)
    header = " & ".join(r"\textbf{" + latex_escape(title) + "}" for _, title, _ in columns)
    lines = [
        r"\begin{landscape}",
        r"\scriptsize",
        r"\setlength{\tabcolsep}{2.2pt}",
        r"\renewcommand{\arraystretch}{1.08}",
        rf"\begin{{longtable}}{{{colspec}}}",
        rf"\caption{{{latex_escape(caption)}}}\label{{{label}}}\\",
        r"\toprule",
        header + r" \\",
        r"\midrule",
        r"\endfirsthead",
        rf"\multicolumn{{{len(columns)}}}{{c}}{{\textit{{Table \thetable{{}} continued}}}}\\",
        r"\toprule",
        header + r" \\",
        r"\midrule",
        r"\endhead",
        r"\midrule",
        rf"\multicolumn{{{len(columns)}}}{{r}}{{\textit{{Continued on next page}}}}\\",
        r"\endfoot",
        r"\bottomrule",
        r"\endlastfoot",
    ]
    for _, row in data.iterrows():
        lines.append(" & ".join(latex_escape(row[column]) for column, _, _ in columns) + r" \\")
    lines.extend(
        [
            r"\end{longtable}",
            r"\begin{minipage}{0.98\linewidth}",
            r"\scriptsize",
            r"\textit{Notes:} " + notes,
            r"\end{minipage}",
            r"\end{landscape}",
            "",
        ]
    )
    return "\n".join(lines)


def paris_latex(audit: pd.DataFrame, scope: str) -> str:
    table = audit.copy()
    if scope == "final":
        retained = table.groupby("iso3")["gap_available"].transform("max").eq(1)
        table = table.loc[retained].copy()
    table["target_year_print"] = table["target_year"].map(year_text)
    table["gap_status_print"] = table.apply(
        lambda row: (
            f"Yes ({year_text(row['first_gap_year'])}--{year_text(row['last_gap_year'])})"
            if row["gap_available"] == 1
            else "No: " + human_label(row["decision_reason"])
        ),
        axis=1,
    )
    table["emissions_rule_print"] = table["emissions_scope_method_std"].replace(
        {
            "EDGAR_SECTORAL": "Sector-matched EDGAR",
            "EDGAR_TOTAL": "National EDGAR total",
            "EDGAR_TOTAL_FALLBACK": "EDGAR total fallback",
            "EDGAR_SECTORAL_PLUS_FAO_FOREST": "Sector EDGAR + FAO forest",
            "EDGAR_TOTAL_PLUS_FAO_FOREST": "EDGAR total + FAO forest",
            "FAO_FOREST_ONLY": "FAO forest only",
        }
    ).map(human_label)
    table["prospective_print"] = table["prospective_input"].replace(
        {
            "PMRCPBIE": "SSP2--RCP6.0",
            "PMRCPBIE_GDPPPP_SSP2": "SSP2 GDP",
            "PMRCPBIE_POP_LEVEL_SSP2": "SSP2 population",
            "NDC_OFFICIAL": "Official NDC",
        }
    )
    columns = [
        ("iso3", "ISO3", "p{0.8cm}"),
        ("cycle", "Cycle", "p{0.75cm}"),
        ("metric_code", "Metric", "p{0.8cm}"),
        ("implementation_period", "Period", "p{1.35cm}"),
        ("conditionality_code", "Cond.", "p{0.75cm}"),
        ("coverage_code", "Coverage", "p{1.35cm}"),
        ("emissions_rule_print", "Observed-emissions rule", "p{2.1cm}"),
        ("forest_rule", "Forest", "p{1.1cm}"),
        ("prospective_print", "Prospective input", "p{1.75cm}"),
        ("gap_status_print", "Gap status / audit note", "p{3.9cm}"),
    ]
    notes = (
        r"The unit of observation is the country--NDC-cycle. Metric codes are BY (base-year), "
        r"BAU (business-as-usual), INT (intensity), PC (per capita), FL (fixed level), "
        r"TRJ (trajectory), NQ (no quantified GHG target), and UNK (unknown). Conditionality "
        r"codes are U (unconditional only), C (conditional only), U+C (both), PC (partially "
        r"conditional), NS (no submission), and NQ (no quantified target). ``FAO added'' means "
        r"that FAOSTAT \textit{Forest Land} is added to the non-LULUCF EDGAR aggregate. "
        r"Coverage is marked ``Not specified'' only when the target scope remains indeterminate "
        r"after harmonisation, and ``Not applicable'' when the cycle contains no quantified GHG "
        r"target or no submission. "
        r"A missing Gap is not imputed from the other NDC cycle."
    )
    return render_longtable(
        table,
        columns,
        "Country-cycle audit of Paris NDC harmonisation decisions",
        "tab:paris_country_cycle_audit",
        notes,
    )


def kyoto_latex(audit: pd.DataFrame) -> str:
    table = audit.copy()
    table["aa_print"] = np.where(table["official_AA_available"].eq(1), "Yes", "No")
    table["annex_print"] = np.where(table["annexA_complete"].eq(1), "5/5", "Incomplete")
    table["lulucf_print"] = table["lulucf_years"].astype(str) + "/5"
    table["trade_print"] = table["trade_years"].astype(str) + "/5"
    table["trade_source_print"] = table["trade_source"].map(human_label)
    columns = [
        ("iso3", "ISO3", "p{0.9cm}"),
        ("party", "Party", "p{3.1cm}"),
        ("aa_print", "Official AA", "p{1.1cm}"),
        ("annex_print", "Annex A", "p{1.2cm}"),
        ("lulucf_print", "LULUCF", "p{1.1cm}"),
        ("trade_print", "ITL flow", "p{1.1cm}"),
        ("accounting_variants", "Available variants", "p{2.1cm}"),
        ("final_decision", "CP1 status", "p{1.5cm}"),
        ("trade_source_print", "Transaction construction", "p{3.2cm}"),
    ]
    notes = (
        r"AA denotes the official assigned amount for the first commitment period (CP1, "
        r"2008--2012). Availability is reported over the five CP1 years. Variants are coded "
        r"B (benchmark), L (LULUCF-adjusted), T (transaction-adjusted), and F (fully adjusted). "
        r"ITL denotes the International Transaction Log."
    )
    return render_longtable(
        table,
        columns,
        "Country audit of Kyoto CP1 accounting inputs",
        "tab:kyoto_country_audit",
        notes,
    )


def write_outputs(
    output_dir: Path,
    paris: pd.DataFrame,
    kyoto: pd.DataFrame,
    appendix_scope: str,
    edgar_release: str,
    edgar_access_date: str,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    paris.to_csv(output_dir / "paris_country_cycle_audit_full.csv", index=False)
    kyoto.to_csv(output_dir / "kyoto_cp1_country_audit_full.csv", index=False)

    metadata = pd.DataFrame(
        [
            ["EDGAR source file", "EDGAR_AR5_GHG_1970_2024.xlsx"],
            ["EDGAR temporal coverage", "1970--2024"],
            ["EDGAR release/vintage", edgar_release],
            ["EDGAR access date", edgar_access_date],
            ["FAOSTAT source file", "Emissions from forests (Global, National - Annual) - FAOSTAT.xlsx"],
            ["Paris target rows", len(paris)],
            ["Paris unique countries", paris["iso3"].nunique()],
            ["Paris NDC1 Gap available", int(paris.loc[paris["cycle"].eq("NDC1"), "gap_available"].sum())],
            ["Paris NDC2 Gap available", int(paris.loc[paris["cycle"].eq("NDC2"), "gap_available"].sum())],
            ["Kyoto CP1 audit countries", kyoto["iso3"].nunique()],
            ["Kyoto CP1 retained countries", int(kyoto["final_decision"].eq("Retained").sum())],
        ],
        columns=["item", "value"],
    )

    with pd.ExcelWriter(output_dir / "country_commitment_audit_workbook.xlsx") as writer:
        paris.to_excel(writer, sheet_name="Paris_country_cycle", index=False)
        kyoto.to_excel(writer, sheet_name="Kyoto_CP1_country", index=False)
        metadata.to_excel(writer, sheet_name="Metadata", index=False)

    (output_dir / "paris_country_cycle_audit_appendix.tex").write_text(
        paris_latex(paris, appendix_scope), encoding="utf-8"
    )
    (output_dir / "kyoto_cp1_country_audit_appendix.tex").write_text(
        kyoto_latex(kyoto), encoding="utf-8"
    )
    metadata.to_csv(output_dir / "audit_metadata_and_counts.csv", index=False)


def main() -> None:
    args = parse_args()
    paris_targets = pd.read_excel(PARIS_DIR / "Paris_NDC_Targets.xlsx")
    paris_final = pd.read_stata(
        PARIS_DIR / "Paris_GapToTarget_Final.dta", convert_categoricals=False
    )
    kyoto_intermediate = pd.read_excel(
        KYOTO_DIR / "kyoto_annual_indicator_input.xlsx", sheet_name="kyoto_panel"
    )
    kyoto_final = pd.read_stata(
        KYOTO_DIR / "Gap_Kyoto.dta", convert_categoricals=False
    )

    paris = build_paris_audit(paris_targets, paris_final)
    kyoto = build_kyoto_audit(kyoto_intermediate, kyoto_final)
    write_outputs(
        args.output_dir,
        paris,
        kyoto,
        args.appendix_scope,
        args.edgar_release,
        args.edgar_access_date,
    )

    print(f"Outputs written to: {args.output_dir.resolve()}")
    print(f"Paris audit: {len(paris)} country-cycle rows; {paris['iso3'].nunique()} countries")
    print(
        "Paris available Gaps: "
        f"NDC1={int(paris.loc[paris['cycle'].eq('NDC1'), 'gap_available'].sum())}, "
        f"NDC2={int(paris.loc[paris['cycle'].eq('NDC2'), 'gap_available'].sum())}"
    )
    print(
        f"Kyoto CP1: {len(kyoto)} audited countries; "
        f"{int(kyoto['final_decision'].eq('Retained').sum())} retained"
    )


if __name__ == "__main__":
    main()
