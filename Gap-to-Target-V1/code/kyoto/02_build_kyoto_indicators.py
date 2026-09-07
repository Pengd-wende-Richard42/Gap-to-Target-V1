
# -*- coding: utf-8 -*-
r"""
Build a Kyoto panel for constructing the revised Gap-to-Target indicators over
the first Kyoto commitment period (CP1: 2008--2012).

Inputs are resolved relative to the repository root.

Required files:
1) Annual Accounting quantity, No type of value, Net emissions_removals, Emissions,
   AAUs, ERUs, RMUs, CERs, tCERs, lCERs and Total_Firstperiod_Kyoto.xlsx

2) itl_country_panel_reporting_period.csv
   (created from the ITL annual reports)

Outputs written to ``data/processed/kyoto``:
- kyoto_annual_indicator_input.csv
- kyoto_annual_indicator_input.xlsx
- kyoto_annual_indicator_input.dta
- kyoto_merge_audit.csv

Main constructed variables:
- assigned_amount_t
- annualized_assigned_amount_t
- annexA_emissions_t
- lulucf_net_t
- net_acquisition_units
- forwarding_units
- cumulative_target_basic_t
- cumulative_target_trade_t
- emissions_plus_lulucf_t
- gap_basic_pct
- gap_trade_pct
- gap_full_pct

The basic indicator remains an annual comparison between Annex A emissions and
one fifth of the CP1 Assigned Amount. The trade-adjusted indicators use
cumulative emissions and cumulative budget positions. This gives each ITL flow
its correct one-for-one accounting weight and makes the 2012 value correspond
to the end-of-period Kyoto balance.

Important:
The ITL reports use reporting windows (typically Nov-Oct or Oct-Sep), not exact
calendar years. The script preserves period_start / period_end and does NOT hide
that distinction. For the final paper, calendar-year SEF data would be preferable
if available.
"""

from __future__ import annotations

import math
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import numpy as np
import pandas as pd


# =============================================================================
# PATHS
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[2]
BASE = PROJECT_ROOT / "data" / "raw" / "kyoto"
OUTPUT_DIR = PROJECT_ROOT / "data" / "processed" / "kyoto"

CP1_FILE = BASE / "Annual_Accounting_quantity_Total_Firstperiod_Kyoto.xlsx"
ITL_FILE = OUTPUT_DIR / "itl" / "itl_country_panel_reporting_period.csv"

# The empirical window is deliberately restricted to the legally homogeneous
# first commitment period. Years from CP2 must not enter this dataset.
ANALYSIS_START_YEAR = 2008
ANALYSIS_END_YEAR = 2012
ANALYSIS_COMMITMENT_PERIOD = 1


# =============================================================================
# COUNTRY CROSSWALK
# =============================================================================

ISO3 = {
    "Australia": "AUS",
    "Austria": "AUT",
    "Belarus": "BLR",
    "Belgium": "BEL",
    "Bulgaria": "BGR",
    "Canada": "CAN",
    "Croatia": "HRV",
    "Cyprus": "CYP",
    "Czech Republic": "CZE",
    "Czechia": "CZE",
    "Denmark": "DNK",
    "Estonia": "EST",
    "European Community": "EUU",
    "European Union": "EUU",
    "Finland": "FIN",
    "France": "FRA",
    "Germany": "DEU",
    "Greece": "GRC",
    "Hungary": "HUN",
    "Iceland": "ISL",
    "Ireland": "IRL",
    "Italy": "ITA",
    "Japan": "JPN",
    "Kazakhstan": "KAZ",
    "Latvia": "LVA",
    "Liechtenstein": "LIE",
    "Lithuania": "LTU",
    "Luxembourg": "LUX",
    "Malta": "MLT",
    "Monaco": "MCO",
    "Netherlands": "NLD",
    "Netherlands (Kingdom of the)": "NLD",
    "New Zealand": "NZL",
    "Norway": "NOR",
    "Poland": "POL",
    "Portugal": "PRT",
    "Romania": "ROU",
    "Russian Federation": "RUS",
    "Slovakia": "SVK",
    "Slovenia": "SVN",
    "Spain": "ESP",
    "Sweden": "SWE",
    "Switzerland": "CHE",
    "Ukraine": "UKR",
    "United Kingdom": "GBR",
    "United Kingdom of Great Britain and Northern Ireland": "GBR",
}


# =============================================================================
# LOW-LEVEL XLSX READER
# =============================================================================
# The UNFCCC exports used here contain an invalid workbook sheet state ("show"),
# which can make openpyxl fail. Therefore this script reads the XLSX XML directly.
# =============================================================================

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def excel_col_to_idx(cell_ref: str) -> int:
    letters = re.match(r"([A-Z]+)", cell_ref).group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_unfccc_xlsx_raw(path: Path) -> pd.DataFrame:
    """Read sheet1.xml from a UNFCCC XLSX export without openpyxl."""
    with zipfile.ZipFile(path) as z:
        # Shared strings
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(f"{{{NS_MAIN}}}si"):
                txt = "".join(
                    (t.text or "")
                    for t in si.iter(f"{{{NS_MAIN}}}t")
                )
                shared.append(txt)

        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        sheet_data = sheet.find(f"{{{NS_MAIN}}}sheetData")

        sparse_rows = {}
        max_col = 0
        max_row = 0

        for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
            excel_row = int(row.attrib.get("r", "0"))
            max_row = max(max_row, excel_row)
            row_dict = {}

            for c in row.findall(f"{{{NS_MAIN}}}c"):
                ref = c.attrib["r"]
                j = excel_col_to_idx(ref)
                max_col = max(max_col, j)

                cell_type = c.attrib.get("t")
                v = c.find(f"{{{NS_MAIN}}}v")

                if cell_type == "inlineStr":
                    is_node = c.find(f"{{{NS_MAIN}}}is")
                    if is_node is None:
                        value = ""
                    else:
                        value = "".join(
                            (t.text or "")
                            for t in is_node.iter(f"{{{NS_MAIN}}}t")
                        )
                elif v is None:
                    value = ""
                else:
                    value = v.text or ""
                    if cell_type == "s" and value != "":
                        value = shared[int(value)]

                row_dict[j] = value

            # Preserve the original Excel row number (1-based).
            sparse_rows[excel_row] = row_dict

        matrix = []
        for excel_row in range(1, max_row + 1):
            r = sparse_rows.get(excel_row, {})
            matrix.append([r.get(j, "") for j in range(max_col + 1)])

    return pd.DataFrame(matrix)


def clean_text(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    x = str(x).replace("\n", " ")
    return re.sub(r"\s+", " ", x).strip()


def normalize_party_name(x) -> str:
    """Normalize UNFCCC Party labels used in the pivot exports."""
    s = clean_text(x)
    s = re.sub(r"\s*\(KP\)\s*$", "", s, flags=re.IGNORECASE)
    return s.strip()


def parse_num(x):
    """Convert UNFCCC numeric cells; treat notation codes such as NE as missing."""
    s = clean_text(x)
    if s == "":
        return np.nan

    # UNFCCC notation keys / non-numeric content
    if s.upper() in {"NE", "NO", "NA", "IE", "C", "-", "—", "–"}:
        return np.nan

    s = s.replace(",", "")
    s = s.replace("−", "-").replace("–", "-").replace("—", "-")
    try:
        return float(s)
    except ValueError:
        return np.nan


# =============================================================================
# BUILD COLUMN METADATA FROM THE UNFCCC PIVOT EXPORT
# =============================================================================

def forward_fill_across(values):
    out = []
    last = ""
    for v in values:
        v = clean_text(v)
        if v:
            last = v
        out.append(last)
    return out


def build_column_metadata(raw: pd.DataFrame) -> pd.DataFrame:
    """
    Rows 4-7 in Excel (0-based indices 3-6) contain:
      Year / Type of value / Gas / Unit
    Forward-fill merged header cells horizontally.
    """
    year = forward_fill_across(raw.iloc[3].tolist())
    value_type = forward_fill_across(raw.iloc[4].tolist())
    gas = forward_fill_across(raw.iloc[5].tolist())
    unit = [clean_text(x) for x in raw.iloc[6].tolist()]

    meta = pd.DataFrame({
        "col_idx": range(raw.shape[1]),
        "year_label": year,
        "value_type": value_type,
        "gas": gas,
        "unit": unit,
    })

    return meta


def match_category(category: str, patterns: list[str]) -> bool:
    c = clean_text(category).lower()
    return all(p.lower() in c for p in patterns)


def first_numeric_in_row(raw: pd.DataFrame, row_idx: int):
    vals = [parse_num(x) for x in raw.iloc[row_idx, 2:].tolist()]
    vals = [x for x in vals if pd.notna(x)]
    return vals[0] if vals else np.nan


# =============================================================================
# EXTRACT KYOTO ACCOUNTING DATA
# =============================================================================

def extract_assigned_amount(raw: pd.DataFrame, period: int) -> pd.DataFrame:
    """
    One Assigned Amount per Party.
    CP1: Article 3 paragraphs 7 and 8
    CP2: Article 3 paragraphs 7bis, 8 and 8bis / Doha Amendment
    """
    rows = []

    for i in range(7, len(raw)):
        party = normalize_party_name(raw.iat[i, 0])
        category = clean_text(raw.iat[i, 1])

        if not party or not category:
            continue

        c = category.lower()

        if period == 1:
            is_aa = (
                "assigned amount" in c
                and "article 3" in c
                and "doha" not in c
            )
        else:
            is_aa = (
                "assigned amount" in c
                and ("7bis" in c or "doha amendment" in c)
            )

        if not is_aa:
            continue

        val = first_numeric_in_row(raw, i)
        rows.append({
            "party": party,
            "assigned_amount_t": val,
            "commitment_period": period,
        })

    out = pd.DataFrame(rows).drop_duplicates("party")
    out["iso3"] = out["party"].map(ISO3)
    # Unmapped aggregate/non-country labels must not create duplicate NaN merge keys.
    out = out[out["iso3"].notna()].drop_duplicates("iso3")

    n_years = 5 if period == 1 else 8
    out["commitment_years"] = n_years
    out["annualized_assigned_amount_t"] = out["assigned_amount_t"] / n_years

    return out


def extract_qelrc(raw: pd.DataFrame, period: int) -> pd.DataFrame:
    rows = []

    for i in range(7, len(raw)):
        party = normalize_party_name(raw.iat[i, 0])
        category = clean_text(raw.iat[i, 1]).lower()

        if not party or not category:
            continue

        if period == 1:
            ok = (
                "quantified emission limitation or reduction commitment" in category
                and "first commitment period" in category
            )
        else:
            ok = (
                "quantified emission limitation or reduction commitment" in category
                and "second commitment period" in category
            )

        if ok:
            val = first_numeric_in_row(raw, i)
            rows.append({"party": party, "qelrc_pct": val})

    out = pd.DataFrame(rows).drop_duplicates("party")
    if len(out):
        out["iso3"] = out["party"].map(ISO3)
        out = out[out["iso3"].notna()].drop_duplicates("iso3")
    return out


def find_column(meta, year, value_type, gas="Aggregate GHGs",
                unit_contains="CO₂ equivalent"):
    y = str(year)

    mask = (
        meta["year_label"].eq(y)
        & meta["value_type"].str.lower().eq(value_type.lower())
        & meta["gas"].str.lower().eq(gas.lower())
        & meta["unit"].str.contains(unit_contains, case=False, na=False)
    )

    hits = meta.loc[mask, "col_idx"].tolist()
    return hits


def extract_annual_category(
    raw: pd.DataFrame,
    meta: pd.DataFrame,
    years: list[int],
    category_patterns: list[str],
    value_type: str,
    output_name: str,
) -> pd.DataFrame:

    # There may be one row per party/category and one desired column per year.
    records = []

    # Desired columns by year
    year_cols = {}
    for year in years:
        hits = find_column(meta, year, value_type=value_type)
        year_cols[year] = hits

    for i in range(7, len(raw)):
        party = normalize_party_name(raw.iat[i, 0])
        category = clean_text(raw.iat[i, 1])

        if not party or not category:
            continue

        if not match_category(category, category_patterns):
            continue

        for year in years:
            hits = year_cols[year]
            if not hits:
                continue

            # Usually a unique aggregate-GHG tCO2e column.
            # If more than one appears, keep the first non-missing numeric value.
            val = np.nan
            chosen_col = np.nan
            for j in hits:
                x = parse_num(raw.iat[i, j])
                if pd.notna(x):
                    val = x
                    chosen_col = j
                    break

            records.append({
                "party": party,
                "year": year,
                output_name: val,
                f"{output_name}_source_category": category,
                f"{output_name}_source_col": chosen_col,
            })

    out = pd.DataFrame(records)

    if len(out):
        out["iso3"] = out["party"].map(ISO3)
        out = out.drop_duplicates(["party", "year"])

    return out


def extract_period(raw: pd.DataFrame, period: int) -> pd.DataFrame:
    meta = build_column_metadata(raw)

    years = list(range(2008, 2013)) if period == 1 else list(range(2013, 2021))

    aa = extract_assigned_amount(raw, period)
    qelrc = extract_qelrc(raw, period)

    # Official Annex A emissions
    annex = extract_annual_category(
        raw=raw,
        meta=meta,
        years=years,
        category_patterns=["total ghg emissions", "annex a"],
        value_type="Emissions",
        output_name="annexA_emissions_t",
    )

    # Annual reported flows are retained for audit, not used as accounting credits.
    lulucf = extract_annual_category(
        raw=raw,
        meta=meta,
        years=years,
        category_patterns=["lulucf activities", "articles 3.3 and 3.4"],
        value_type="Net emissions/removals",
        output_name="lulucf_net_t",
    )

    panel = annex.merge(
        lulucf[["party", "year", "lulucf_net_t"]],
        on=["party", "year"],
        how="outer",
    )

    panel["iso3"] = panel["party"].map(ISO3)

    if period != 1:
        raise ValueError("LULUCF accounting allocation is defined for CP1 only.")
    accounting_cols = meta.loc[
        meta["value_type"].str.strip().eq("Accounting quantity")
        & meta["year_label"].str.strip().eq("2008-2012"), "col_idx"
    ].tolist()
    if len(accounting_cols) != 1:
        raise ValueError("Expected one CP1 accounting-quantity column.")
    activities = {"Afforestation and reforestation", "Deforestation", "Forest management",
                  "Cropland management", "Grazing land management", "Revegetation"}
    records = []
    for _, row in raw.iterrows():
        activity = clean_text(row.iloc[1])
        if activity not in activities:
            continue
        party = normalize_party_name(row.iloc[0])
        value = clean_text(row.iloc[accounting_cols[0]])
        number = parse_num(value)
        # Blank/NA/NO cells do not add a credit; unresolved notation blocks the total.
        unresolved = pd.isna(number) and bool(value) and not set(value.upper().split(",")) <= {"NA", "NO"}
        records.append({"party": party, "activity": activity, "source_value": value,
                        "quantity_t": number, "unresolved": unresolved})
    detail = pd.DataFrame(records)
    if detail.duplicated(["party", "activity"]).any():
        raise ValueError("Duplicate LULUCF activity accounting rows.")
    totals = detail.groupby("party")["quantity_t"].sum(min_count=1)
    blocked = detail.groupby("party")["unresolved"].any()
    totals.loc[blocked] = np.nan
    panel = panel.rename(columns={"lulucf_net_t": "lulucf_reported_t"})
    panel["lulucf_cp1_accounting_t"] = panel["party"].map(totals)
    panel["lulucf_net_t"] = panel["lulucf_cp1_accounting_t"] / 5
    panel["lulucf_method"] = "CP1_ACCOUNTING_EQUAL_ALLOCATION"
    audit_dir = PROJECT_ROOT / "outputs" / "audit" / "kyoto"
    audit_dir.mkdir(parents=True, exist_ok=True)
    detail.to_csv(audit_dir / "lulucf_accounting_by_activity.csv", index=False)

    panel = panel.merge(
        aa[[
            "iso3", "assigned_amount_t", "annualized_assigned_amount_t",
            "commitment_period", "commitment_years"
        ]],
        on="iso3",
        how="left",
        validate="m:1",
    )

    if len(qelrc):
        panel = panel.merge(
            qelrc[["iso3", "qelrc_pct"]],
            on="iso3",
            how="left",
            validate="m:1",
        )

    return panel


# =============================================================================
# ITL DATA
# =============================================================================

def load_itl(path: Path) -> pd.DataFrame:
    itl = pd.read_csv(path)

    # We use only Party registries and preserve the ITL reporting window.
    if "entity_type" in itl.columns:
        itl = itl[itl["entity_type"].eq("party_registry")].copy()

    keep = [
        "iso3",
        "registry",
        "report_year",
        "period_start",
        "period_end",
        "source_flow_type",
        "acquisition_units",
        "transfer_units",
        "net_acquisition_units",
        "net_transfer_units",
        "forwarding_units",
        "internal_transfer_units",
        "issuance_units",
        "retirement_units",
        "cancellation_units",
    ]

    keep = [c for c in keep if c in itl.columns]
    itl = itl[keep].copy()

    itl = itl.rename(columns={"report_year": "year"})

    # Force transaction variables numeric.
    for c in [
        "acquisition_units", "transfer_units", "net_acquisition_units",
        "net_transfer_units", "forwarding_units",
        "internal_transfer_units", "issuance_units",
        "retirement_units", "cancellation_units"
    ]:
        if c in itl:
            itl[c] = pd.to_numeric(itl[c], errors="coerce")

    return itl


# =============================================================================
# CONSTRUCT INDICATOR INPUTS
# =============================================================================

def safe_gap(emissions, target):
    return np.where(
        target.notna() & target.gt(0) & emissions.notna(),
        100.0 * (emissions - target) / target,
        np.nan,
    )


def build_indicator_panel(kyoto: pd.DataFrame, itl: pd.DataFrame) -> pd.DataFrame:
    if itl[["acquisition_units", "transfer_units"]].lt(0).any().any():
        raise ValueError("Negative gross ITL flows: use the corrected annual extraction.")
    if itl["source_flow_type"].eq("derived_flow_from_cumulative").any():
        raise ValueError("Obsolete ITL 2011 conversion; apply FCCC/KP/CMP/2011/7/Corr.1.")
    df = kyoto.merge(
        itl,
        on=["iso3", "year"],
        how="left",
        validate="1:1",
    )

    # Raw-data availability flags are retained separately from the values used
    # in the scenario variants.
    df["itl_transaction_observed"] = df["net_acquisition_units"].notna()
    df["lulucf_observed"] = df["lulucf_net_t"].notna()

    # Units are compatible:
    # 1 Kyoto unit = 1 t CO2e.
    # All official UNFCCC values extracted here are kept in t CO2e.
    # Therefore no conversion is needed before combining them.

    # -------------------------------------------------------------------------
    # TARGETS
    # -------------------------------------------------------------------------

    # 1. Pure annualized Assigned Amount benchmark
    df["annual_target_basic_t"] = df["annualized_assigned_amount_t"]

    # -------------------------------------------------------------------------
    # EMISSIONS
    # -------------------------------------------------------------------------

    # Kyoto-specific accounting robustness. The basic Gap is the common
    # starting point: when LULUCF information is unavailable, the adjustment is
    # neutral and the LULUCF variant therefore equals the corresponding basic
    # variant. The flag below keeps this fallback fully traceable.
    # positive LULUCF net value raises accounted emissions;
    # negative value (net removals) lowers them.
    df["lulucf_basic_fallback"] = df["lulucf_net_t"].isna()
    df["lulucf_net_used_t"] = df["lulucf_net_t"].fillna(0)
    df["emissions_plus_lulucf_t"] = (
        df["annexA_emissions_t"]
        + df["lulucf_net_used_t"]
    )

    df = df.sort_values(["iso3", "year"]).copy()

    # -------------------------------------------------------------------------
    # CUMULATIVE KYOTO ACCOUNTING
    # -------------------------------------------------------------------------

    # At year t, k/5 of the Assigned Amount has elapsed. In 2012 this budget is
    # exactly equal to the complete CP1 Assigned Amount.
    df["period_year_index"] = df["year"] - ANALYSIS_START_YEAR + 1
    df["cumulative_target_basic_t"] = (
        df["assigned_amount_t"]
        * df["period_year_index"]
        / df["commitment_years"]
    )

    by_country = df.groupby("iso3", sort=False)
    df["cumulative_annexA_emissions_t"] = by_country["annexA_emissions_t"].cumsum()
    df["cumulative_lulucf_net_t"] = by_country["lulucf_net_used_t"].cumsum()
    df["cumulative_emissions_plus_lulucf_t"] = (
        df["cumulative_annexA_emissions_t"]
        + df["cumulative_lulucf_net_t"]
    )

    # ITL hierarchy used in the trade variant:
    #   1) observed annual flow;
    #   2) latest previously observed flow for an internal gap (LOCF);
    #   3) neutral adjustment when neither is available, so the variant falls
    #      back to the cumulative basic Gap.
    carried_trade = by_country["net_acquisition_units"].ffill()
    df["itl_locf_imputed"] = (
        df["net_acquisition_units"].isna()
        & carried_trade.notna()
        & df["year"].gt(ANALYSIS_START_YEAR)
    )
    df["itl_basic_fallback"] = (
        df["net_acquisition_units"].isna()
        & carried_trade.isna()
    )
    df["net_acquisition_units_used"] = carried_trade.fillna(0)

    df["cumulative_net_acquisition_units"] = (
        df["net_acquisition_units_used"]
        .groupby(df["iso3"])
        .cumsum()
    )

    df["cumulative_target_trade_t"] = (
        df["cumulative_target_basic_t"]
        + df["cumulative_net_acquisition_units"]
    )

    # Annualised trade-adjusted benchmark. The cumulative transaction position
    # observed by year t is spread over the five CP1 years, preserving an annual
    # Gap definition comparable with the Paris indicators.
    df["annual_target_trade_t"] = (
        df["assigned_amount_t"]
        + df["cumulative_net_acquisition_units"]
    ) / df["commitment_years"]
    if df["annual_target_trade_t"].le(0).any():
        raise ValueError("Non-positive adjusted budget: audit transactions before interpreting relative Gaps.")
    # -------------------------------------------------------------------------
    # GAPS
    # -------------------------------------------------------------------------

    # Baseline: official Annex A emissions vs annualized Assigned Amount
    df["gap_basic_t"] = df["annexA_emissions_t"] - df["annual_target_basic_t"]
    df["gap_basic_pct"] = safe_gap(
        df["annexA_emissions_t"],
        df["annual_target_basic_t"],
    )

    # Annual LULUCF robustness without transaction-unit adjustments.
    df["gap_lulucf_t"] = (
        df["emissions_plus_lulucf_t"] - df["annual_target_basic_t"]
    )
    df["gap_lulucf_pct"] = safe_gap(
        df["emissions_plus_lulucf_t"],
        df["annual_target_basic_t"],
    )

    # Annual trade-adjusted and full variants.
    df["gap_trade_t"] = df["annexA_emissions_t"] - df["annual_target_trade_t"]
    df["gap_trade_pct"] = safe_gap(
        df["annexA_emissions_t"], df["annual_target_trade_t"]
    )
    df["gap_full_t"] = (
        df["emissions_plus_lulucf_t"] - df["annual_target_trade_t"]
    )
    df["gap_full_pct"] = safe_gap(
        df["emissions_plus_lulucf_t"], df["annual_target_trade_t"]
    )

    # Implementation/progress scores are intentionally not constructed. They
    # are excluded from the current empirical framework because year-to-year
    # ratios become unstable when the target gap is close to zero.

    # Useful units in ktCO2e for merging with datasets already stored in kt.
    t_cols = [
        "assigned_amount_t",
        "annualized_assigned_amount_t",
        "annexA_emissions_t",
        "lulucf_net_t",
        "lulucf_net_used_t",
        "annual_target_basic_t", "annual_target_trade_t",
        "annual_target_trade_t",
        "emissions_plus_lulucf_t",
        "cumulative_target_basic_t",
        "cumulative_target_trade_t",
        "cumulative_annexA_emissions_t",
        "cumulative_lulucf_net_t",
        "cumulative_emissions_plus_lulucf_t",
        "gap_basic_t", "gap_lulucf_t", "gap_trade_t", "gap_full_t",
    ]

    for c in t_cols:
        if c in df:
            df[c.replace("_t", "_kt")] = df[c] / 1000.0

    # ITL unit counts are 1 unit = 1 tCO2e.
    for c in [
        "acquisition_units", "transfer_units", "net_acquisition_units",
        "forwarding_units", "retirement_units", "cancellation_units",
    ]:
        if c in df:
            df[c.replace("_units", "_kt")] = df[c] / 1000.0

    return df


# =============================================================================
# AUDIT
# =============================================================================

def make_audit(df: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for cp, g in df.groupby("commitment_period", dropna=False):
        rows.append({
            "commitment_period": cp,
            "min_year": g["year"].min(),
            "max_year": g["year"].max(),
            "rows": len(g),
            "countries": g["iso3"].nunique(),
            "assigned_amount_nonmissing": g["assigned_amount_t"].notna().sum(),
            "annexA_nonmissing": g["annexA_emissions_t"].notna().sum(),
            "lulucf_nonmissing": g["lulucf_net_t"].notna().sum(),
            "itl_net_acq_nonmissing": g["net_acquisition_units"].notna().sum(),
            "gap_basic_nonmissing": g["gap_basic_pct"].notna().sum(),
            "gap_trade_nonmissing": g["gap_trade_pct"].notna().sum(),
            "gap_full_nonmissing": g["gap_full_pct"].notna().sum(),
            "itl_locf_imputed": g["itl_locf_imputed"].sum(),
            "itl_basic_fallback": g["itl_basic_fallback"].sum(),
            "lulucf_basic_fallback": g["lulucf_basic_fallback"].sum(),
        })

    return pd.DataFrame(rows)


def make_method_comparison(df: pd.DataFrame) -> pd.DataFrame:
    """Summarise the retained Gap variants for transparent validation."""
    variables = [
        "gap_basic_pct",
        "gap_lulucf_pct",
        "gap_trade_pct",
        "gap_full_pct",
    ]
    variables = [c for c in variables if c in df.columns]

    rows = []
    for year, group in df.groupby("year", sort=True):
        for variable in variables:
            values = pd.to_numeric(group[variable], errors="coerce").dropna()
            rows.append({
                "year": year,
                "variable": variable,
                "n": len(values),
                "mean": values.mean() if len(values) else np.nan,
                "median": values.median() if len(values) else np.nan,
                "std": values.std(ddof=1) if len(values) > 1 else np.nan,
                "min": values.min() if len(values) else np.nan,
                "max": values.max() if len(values) else np.nan,
            })
    return pd.DataFrame(rows)


def build_gap_stata_export(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str]]:
    """Create the compact, labelled Stata dataset used in the analysis."""
    gap = df[[
        "iso3", "year",
        # Annual Gap-to-Target variants retained for the analysis
        "gap_basic_t", "gap_basic_pct",
        "gap_lulucf_t", "gap_lulucf_pct",
        "gap_trade_t", "gap_trade_pct",
        "gap_full_t", "gap_full_pct",
    ]].copy()

    labels = {
        "iso3": "ISO3 country code",
        "year": "Calendar year (Kyoto CP1: 2008-2012)",
        "gap_basic_t": "Annual Annex A emissions minus annualised Assigned Amount (tCO2e)",
        "gap_basic_pct": "Annual basic Gap-to-Target relative to annualised Assigned Amount (%)",
        "gap_lulucf_t": "Annual Annex A plus Art. 3.3/3.4 LULUCF gap (tCO2e)",
        "gap_lulucf_pct": "Annual Annex A plus Art. 3.3/3.4 LULUCF Gap-to-Target (%)",
        "gap_trade_t": "Annual Annex A gap with cumulative net trade annualised over CP1 (tCO2e)",
        "gap_trade_pct": "Annual Annex A Gap-to-Target with annualised cumulative net trade (%)",
        "gap_full_t": "Annual Annex A plus LULUCF gap with annualised net trade (tCO2e)",
        "gap_full_pct": "Annual full Gap-to-Target including LULUCF and annualised net trade (%)",
    }

    if gap.duplicated(["iso3", "year"]).any():
        raise AssertionError("Duplicate country-year observations in Gap_Kyoto.dta.")
    if not gap["year"].between(ANALYSIS_START_YEAR, ANALYSIS_END_YEAR).all():
        raise AssertionError("Gap_Kyoto.dta contains observations outside 2008--2012.")

    return gap, labels


# =============================================================================
# RUN
# =============================================================================

def main():
    for f in [CP1_FILE, ITL_FILE]:
        if not f.exists():
            raise FileNotFoundError(f"Missing input file: {f}")

    print("Reading UNFCCC CP1 export...")
    raw1 = read_unfccc_xlsx_raw(CP1_FILE)

    print("Extracting CP1...")
    kyoto = extract_period(raw1, period=ANALYSIS_COMMITMENT_PERIOD)
    kyoto = kyoto[
        kyoto["year"].between(ANALYSIS_START_YEAR, ANALYSIS_END_YEAR)
    ].copy()

    print("Reading ITL transactions...")
    itl = load_itl(ITL_FILE)
    itl = itl[itl["year"].between(ANALYSIS_START_YEAR, ANALYSIS_END_YEAR)].copy()

    print("Building annual Kyoto indicator panel...")
    out = build_indicator_panel(kyoto, itl)

    # Exclude aggregate EU registry if present; retain individual countries.
    out = out[out["iso3"].ne("EUU")].copy()

    # Legal-period safeguards: fail loudly if CP2 observations enter the panel.
    if not out["year"].between(ANALYSIS_START_YEAR, ANALYSIS_END_YEAR).all():
        raise AssertionError("The Kyoto panel contains observations outside 2008--2012.")
    if not out["commitment_period"].eq(ANALYSIS_COMMITMENT_PERIOD).all():
        raise AssertionError("The Kyoto panel contains observations outside CP1.")
    if out.duplicated(["iso3", "year"]).any():
        raise AssertionError("Duplicate country-year observations detected.")

    # In 2012, the cumulative basic budget must equal the complete Assigned
    # Amount, up to floating-point precision.
    end = out[out["year"].eq(ANALYSIS_END_YEAR)].copy()
    valid_end = end[["cumulative_target_basic_t", "assigned_amount_t"]].dropna()
    if not np.allclose(
        valid_end["cumulative_target_basic_t"],
        valid_end["assigned_amount_t"],
        rtol=1e-10,
        atol=1e-6,
    ):
        raise AssertionError("The 2012 cumulative budget does not equal CP1 Assigned Amount.")

    # Order the most relevant variables first.
    first = [
        "iso3", "party", "year", "commitment_period", "commitment_years",
        "qelrc_pct",
        "assigned_amount_t", "annualized_assigned_amount_t",
        "annexA_emissions_t", "lulucf_net_t", "lulucf_net_used_t",
        "emissions_plus_lulucf_t", "lulucf_observed",
        "lulucf_basic_fallback",
        "acquisition_units", "transfer_units", "net_acquisition_units",
        "forwarding_units", "retirement_units", "cancellation_units",
        "annual_target_basic_t",
        "period_year_index", "cumulative_target_basic_t",
        "net_acquisition_units_used", "itl_transaction_observed",
        "itl_locf_imputed", "itl_basic_fallback",
        "cumulative_net_acquisition_units", "cumulative_target_trade_t",
        "cumulative_annexA_emissions_t", "cumulative_lulucf_net_t",
        "cumulative_emissions_plus_lulucf_t",
        "gap_basic_t", "gap_basic_pct", "gap_lulucf_t", "gap_lulucf_pct",
        "gap_trade_t", "gap_trade_pct", "gap_full_t", "gap_full_pct",
        "period_start", "period_end", "source_flow_type",
    ]
    first = [c for c in first if c in out.columns]
    rest = [c for c in out.columns if c not in first]
    out = out[first + rest]

    audit = make_audit(out)
    method_comparison = make_method_comparison(out)

    # -------------------------------------------------------------------------
    # EXPORTS
    # -------------------------------------------------------------------------

    # Input files remain in Row data, while all constructed Gap datasets and
    # validation outputs are stored in the dedicated GAP kyoto directory.
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    csv_out = OUTPUT_DIR / "kyoto_annual_indicator_input.csv"
    xlsx_out = OUTPUT_DIR / "kyoto_annual_indicator_input.xlsx"
    dta_out = OUTPUT_DIR / "kyoto_annual_indicator_input.dta"
    audit_out = OUTPUT_DIR / "kyoto_merge_audit.csv"
    comparison_out = OUTPUT_DIR / "kyoto_method_comparison.csv"
    gap_dta_out = OUTPUT_DIR / "Gap_Kyoto.dta"

    out.to_csv(csv_out, index=False)
    audit.to_csv(audit_out, index=False)
    method_comparison.to_csv(comparison_out, index=False)

    with pd.ExcelWriter(xlsx_out, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="kyoto_panel", index=False)
        audit.to_excel(writer, sheet_name="audit", index=False)
        method_comparison.to_excel(writer, sheet_name="method_comparison", index=False)

    # Stata-compatible export
    stata = out.copy()

    # Convert booleans to 0/1
    for c in stata.select_dtypes(include="bool").columns:
        stata[c] = stata[c].astype("int8")

    # Shorten names >32 chars
    rename_stata = {
        "annualized_assigned_amount_t": "annual_assigned_t",
        "annualized_assigned_amount_kt": "annual_assigned_kt",
        "emissions_plus_lulucf_t": "emis_plus_lulucf_t",
        "emissions_plus_lulucf_kt": "emis_plus_lulucf_kt",
        "cumulative_emissions_plus_lulucf_t": "cum_emis_lulucf_t",
        "cumulative_emissions_plus_lulucf_kt": "cum_emis_lulucf_kt",
        "itl_transaction_observed": "itl_observed",
    }
    stata = stata.rename(columns=rename_stata)

    # Guarantee Stata variable-name length
    stata.columns = [
        c[:32] if len(c) > 32 else c
        for c in stata.columns
    ]

    stata.to_stata(dta_out, write_index=False, version=118)

    # Compact analytical file: identifiers and Gap variants only, with Stata
    # variable labels documenting each accounting perimeter and specification.
    gap_stata, gap_variable_labels = build_gap_stata_export(out)
    gap_stata.to_stata(
        gap_dta_out,
        write_index=False,
        version=118,
        data_label="Kyoto CP1 Gap-to-Target variants, 2008-2012",
        variable_labels=gap_variable_labels,
    )

    print("\n=== DONE ===")
    print(audit.to_string(index=False))
    print("\nSaved:")
    print(csv_out)
    print(xlsx_out)
    print(dta_out)
    print(audit_out)
    print(comparison_out)
    print(gap_dta_out)


if __name__ == "__main__":
    main()
