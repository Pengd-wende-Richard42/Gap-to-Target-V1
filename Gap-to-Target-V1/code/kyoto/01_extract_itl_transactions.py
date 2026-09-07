
from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import camelot
import numpy as np
import pandas as pd

# =============================================================================
# CONFIGURATION
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[2]
BASE = PROJECT_ROOT / "data" / "raw" / "kyoto" / "itl_reports"
OUTPUT_DIR = PROJECT_ROOT / "data" / "processed" / "kyoto" / "itl"

FILES = {year: BASE / f"ITL_{year}.pdf" for year in range(2008, 2016)}

# Pages vérifiées dans les 7 rapports contenant effectivement les annexes
# de transactions (le rapport 2008 ne publie pas encore ces annexes).
TABLE_PAGES = {
    2009: {"counts": [28, 29], "units": [30, 31]},
    2010: {"counts": [31, 32], "units": [33, 34]},
    2011: {"counts": [21, 22], "units": [23, 24]},
    2012: {"counts": [21, 22], "units": [23, 24]},
    2013: {"counts": [28, 29], "units": [30, 31]},
    2014: {"counts": [22, 23], "units": [24, 25]},
    2015: {"counts": [22, 23], "units": [24, 25]},
}

# Périodes couvertes par les annexes ITL.
# ATTENTION : ce ne sont pas des années civiles.
PERIODS = {
    2009: ("2008-11-01", "2009-10-31", "flow"),
    2010: ("2009-11-01", "2010-10-31", "flow"),
    # FCCC/KP/CMP/2011/7/Corr.1: the annex start year is 2010, not 2008.
    2011: ("2010-11-01", "2011-10-31", "flow"),
    2012: ("2011-11-01", "2012-10-31", "flow"),
    2013: ("2012-11-01", "2013-09-30", "flow"),
    2014: ("2013-10-01", "2014-09-30", "flow"),
    2015: ("2014-10-01", "2015-09-30", "flow"),
}

UNIT_COLS = [
    "acquisition_units",
    "transfer_units",
    "net_transfer_reported",
    "forwarding_units",
    "internal_transfer_units",
    "issuance_units",
    "retirement_units",
    "cancellation_units",
]

COUNT_COLS = [
    "acquisition_tx_count",
    "transfer_tx_count",
    "forwarding_tx_count",
    "internal_transfer_tx_count",
    "issuance_tx_count",
    "retirement_tx_count",
    "cancellation_tx_count",
    "total_tx_count",
]

ISO3 = {
    "Australia": "AUS", "Austria": "AUT", "Belarus": "BLR", "Belgium": "BEL",
    "Bulgaria": "BGR", "Canada": "CAN", "Croatia": "HRV", "Cyprus": "CYP",
    "Czechia": "CZE", "Denmark": "DNK", "Estonia": "EST", "Finland": "FIN",
    "France": "FRA", "Germany": "DEU", "Greece": "GRC", "Hungary": "HUN",
    "Iceland": "ISL", "Ireland": "IRL", "Italy": "ITA", "Japan": "JPN",
    "Kazakhstan": "KAZ", "Latvia": "LVA", "Liechtenstein": "LIE",
    "Lithuania": "LTU", "Luxembourg": "LUX", "Malta": "MLT", "Monaco": "MCO",
    "Netherlands": "NLD", "New Zealand": "NZL", "Norway": "NOR", "Poland": "POL",
    "Portugal": "PRT", "Romania": "ROU", "Russian Federation": "RUS",
    "Slovakia": "SVK", "Slovenia": "SVN", "Spain": "ESP", "Sweden": "SWE",
    "Switzerland": "CHE", "Ukraine": "UKR", "United Kingdom": "GBR",
    "European Union": "EUU", "Clean development mechanism": "CDM",
}

# =============================================================================
# HELPERS
# =============================================================================

def clean_text(x) -> str:
    if x is None:
        return ""
    x = str(x).replace("\n", " ")
    return re.sub(r"\s+", " ", x).strip()


def parse_int(x) -> Optional[int]:
    s = clean_text(x)
    if not s:
        return None
    s = s.replace("–", "-").replace("−", "-").replace("—", "-")
    if not re.search(r"\d", s):
        return None
    sign = -1 if s.startswith("-") else 1
    digits = re.sub(r"\D", "", s)
    return sign * int(digits) if digits else None


def normalize_registry(label: str) -> str:
    label = clean_text(label)

    # Remove footnote letter stuck to a registry name, e.g. "Australiai".
    for suffix in list("abcdefghi"):
        if label.endswith(suffix):
            candidate = label[:-1].strip()
            if candidate in ISO3:
                label = candidate
                break

    # Normalize common historical labels / multiline fragments.
    low = label.lower()
    if low.startswith("clean development mechanism") or low == "clean":
        return "Clean development mechanism"
    if low.startswith("european community") or low.startswith("european union") or low == "european":
        return "European Union"
    if low.startswith("czech"):
        return "Czechia"
    if low.startswith("russian"):
        return "Russian Federation"
    if low.startswith("united kingdom"):
        return "United Kingdom"

    return label


def is_incomplete_prefix(label: str) -> bool:
    l = clean_text(label).lower()
    prefixes = {
        "clean", "clean development",
        "european", "european community",
        "czech",
        "russian",
        "united", "united kingdom", "united kingdom of",
        "united kingdom of great britain",
        "united kingdom of great britain and northern",
    }
    return l in prefixes


def header_location(df: pd.DataFrame, kind: str):
    for i, row in df.iterrows():
        vals = [clean_text(v).lower() for v in row.tolist()]
        ridx = next((j for j, v in enumerate(vals) if v == "registry"), None)
        if ridx is None:
            continue
        tail = " | ".join(vals[ridx:])
        if "acquisition" not in tail or "transfer" not in tail:
            continue
        if kind == "units" and "net transfer" not in tail:
            continue
        return i, ridx
    return None, None


def choose_camelot_table(pdf: Path, page: int, kind: str):
    tables = camelot.read_pdf(str(pdf), pages=str(page), flavor="stream")
    scored = []

    for table in tables:
        df = table.df.copy()
        h, ridx = header_location(df, kind)
        score = df.shape[0]
        if df.shape[1] >= 8:
            score += 20
        if h is not None:
            score += 100
        scored.append((score, df, h, ridx))

    if not scored:
        return None, None, None

    _, df, h, ridx = max(scored, key=lambda z: z[0])
    return df, h, ridx


def extract_page_rows(pdf: Path, page: int, kind: str):
    """Extract registry + 8 numeric columns from one annex page."""
    expected_numeric = 8
    df, h, ridx = choose_camelot_table(pdf, page, kind)

    if df is None:
        return []

    if h is not None:
        data = df.iloc[h + 1 :, ridx : ridx + 1 + expected_numeric].copy()
    else:
        # Continuation page without repeated header.
        data = df.copy()
        nonempty_cols = []
        for c in data.columns:
            if not all(clean_text(v) == "" for v in data[c]):
                nonempty_cols.append(c)
        data = data[nonempty_cols]

        if data.shape[1] > 1 + expected_numeric:
            data = data.iloc[:, -(1 + expected_numeric):]
        if data.shape[1] < 1 + expected_numeric:
            return []
        data = data.iloc[:, : 1 + expected_numeric]

    rows = []
    pending_before = []

    for _, row in data.iterrows():
        vals = [clean_text(v) for v in row.tolist()]
        label = vals[0]
        numbers = [parse_int(v) for v in vals[1:1 + expected_numeric]]
        has_data = sum(v is not None for v in numbers) >= 6

        if not has_data:
            if not label or label.lower() in {"registry", "total"} or label.isdigit() or label.startswith("FCCC/"):
                continue

            # Some PDFs put the continuation of a country name AFTER the numeric row:
            # Clean -> development -> mechanism, Russian -> Federation, etc.
            if rows and is_incomplete_prefix(rows[-1][0]):
                rows[-1] = (clean_text(rows[-1][0] + " " + label), rows[-1][1])
            else:
                pending_before.append(label)
            continue

        if pending_before:
            label = clean_text(" ".join(pending_before + ([label] if label else [])))
            pending_before = []

        if not label or label.lower() == "total":
            continue

        rows.append((label, numbers))

    return rows


def extract_kind(year: int, kind: str) -> pd.DataFrame:
    names = UNIT_COLS if kind == "units" else COUNT_COLS
    records = []

    for page in TABLE_PAGES[year][kind]:
        for raw_label, numbers in extract_page_rows(FILES[year], page, kind):
            registry = normalize_registry(raw_label)

            # Skip obvious table furniture that occasionally survives parsing.
            if registry.lower() == "total":
                continue

            rec = {
                "report_year": year,
                "source_page": page,
                "registry_raw": raw_label,
                "registry": registry,
            }
            rec.update(dict(zip(names, numbers)))
            records.append(rec)

    out = pd.DataFrame(records)
    if out.empty:
        return out

    # Keep one row per registry/report year.
    out = out.drop_duplicates(["report_year", "registry"], keep="last")
    return out


def add_metadata(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["period_start"] = pd.to_datetime(df["report_year"].map(lambda y: PERIODS[y][0]))
    df["period_end"] = pd.to_datetime(df["report_year"].map(lambda y: PERIODS[y][1]))
    df["source_flow_type"] = df["report_year"].map(lambda y: PERIODS[y][2])
    df["period_days"] = (df["period_end"] - df["period_start"]).dt.days + 1
    df["iso3"] = df["registry"].map(ISO3)
    df["entity_type"] = np.where(
        df["iso3"].isin(["EUU", "CDM"]),
        "non_country_registry",
        "party_registry",
    )
    return df


def validate_annual_flows(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Keep annual flows unchanged; apply the official 2011 net-column correction."""
    out = df.copy()
    idx = out["report_year"].eq(2011)
    primitive = [c for c in columns if "net_transfer" not in c and "total_tx_count" not in c]
    if out[primitive].lt(0).any().any():
        raise ValueError("Negative gross transactions: check extraction before building indicators.")
    if "net_transfer_reported" in columns:
        # The original published value remains in net_transfer_reported_source.
        out.loc[idx, "net_transfer_reported"] = (
            out.loc[idx, "transfer_units"] - out.loc[idx, "acquisition_units"])
    return out


# =============================================================================
# MAIN
# =============================================================================

def main():
    missing = [str(p) for p in FILES.values() if not p.exists()]
    if missing:
        raise FileNotFoundError("Missing PDFs:\n" + "\n".join(missing))

    units_all = []
    counts_all = []
    audit = []

    for year in range(2009, 2016):
        u = extract_kind(year, "units")
        c = extract_kind(year, "counts")

        units_all.append(u)
        counts_all.append(c)

        audit.append({
            "report_year": year,
            "unit_rows": len(u),
            "count_rows": len(c),
            "unit_missing_iso3": int(u["registry"].map(ISO3).isna().sum()) if len(u) else 0,
            "count_missing_iso3": int(c["registry"].map(ISO3).isna().sum()) if len(c) else 0,
        })

    units_source = add_metadata(pd.concat(units_all, ignore_index=True))
    counts_source = add_metadata(pd.concat(counts_all, ignore_index=True))

    # Keep source values for auditability.
    for c in UNIT_COLS:
        units_source[c + "_source"] = units_source[c]

    units = validate_annual_flows(units_source, UNIT_COLS)

    # Never trust/reuse the reported net column mechanically.
    # Official definition in the footnotes: Net transfer = Transfer - Acquisition.
    units["net_transfer_units"] = units["transfer_units"] - units["acquisition_units"]
    units["net_acquisition_units"] = units["acquisition_units"] - units["transfer_units"]

    # Diagnostic comparison with the PDF's own reported net column.
    units["reported_net_minus_recomputed"] = (
        units["net_transfer_reported"] - units["net_transfer_units"]
    )
    units["flag_reported_net_inconsistent"] = (
        units["reported_net_minus_recomputed"].fillna(0) != 0
    )

    # Forwarding is NOT included in acquisition: it is a separate flow from the CDM registry.
    # Keep both variables separately and offer this broader inflow measure for later robustness.
    units["net_external_inflow_plus_forwarding"] = (
        units["net_acquisition_units"] + units["forwarding_units"].fillna(0)
    )

    counts = validate_annual_flows(counts_source, COUNT_COLS)

    # Independently check transaction-count totals.
    count_components = [
        "acquisition_tx_count", "transfer_tx_count", "forwarding_tx_count",
        "internal_transfer_tx_count", "issuance_tx_count",
        "retirement_tx_count", "cancellation_tx_count"
    ]
    counts["total_tx_count_recomputed"] = counts[count_components].sum(axis=1, min_count=1)

    combined = units.merge(
        counts[["report_year", "registry"] + COUNT_COLS + ["total_tx_count_recomputed"]],
        on=["report_year", "registry"],
        how="left",
        validate="1:1",
    )

    # Country-only panel, merge-ready with Assigned Amounts / inventories.
    country = combined[
        (combined["entity_type"] == "party_registry") &
        combined["iso3"].notna()
    ].copy()

    # Explicitly note that 2008 has no transaction-unit annex in ITL_2008.pdf.
    audit.insert(0, {
        "report_year": 2008,
        "unit_rows": 0,
        "count_rows": 0,
        "unit_missing_iso3": 0,
        "count_missing_iso3": 0,
        "note": "ITL 2008 report contains no country transaction-unit annex; do not code as zero."
    })

    audit_df = pd.DataFrame(audit)

    OUT = OUTPUT_DIR
    OUT.mkdir(parents=True, exist_ok=True)

    units_source.to_csv(OUT / "itl_units_source_tables.csv", index=False)
    units.to_csv(OUT / "itl_units_reporting_period_flows.csv", index=False)
    counts.to_csv(OUT / "itl_transaction_counts_reporting_period.csv", index=False)
    country.to_csv(OUT / "itl_country_panel_reporting_period.csv", index=False)
    audit_df.to_csv(OUT / "itl_extraction_audit.csv", index=False)

    # Excel inspection workbook.
    with pd.ExcelWriter(OUT / "itl_country_panel_reporting_period.xlsx", engine="openpyxl") as writer:
        country.to_excel(writer, sheet_name="country_panel", index=False)
        units_source.to_excel(writer, sheet_name="source_units", index=False)
        counts.to_excel(writer, sheet_name="transaction_counts", index=False)
        audit_df.to_excel(writer, sheet_name="audit", index=False)

    # Stata: store dates as YYYY-MM-DD strings for maximum compatibility.
    stata = country.copy()
    stata["period_start"] = stata["period_start"].dt.strftime("%Y-%m-%d")
    stata["period_end"] = stata["period_end"].dt.strftime("%Y-%m-%d")
    # Stata variable names <= 32 chars.
    stata = stata.rename(columns={
        "net_external_inflow_plus_forwarding": "net_inflow_plus_forwarding",
        "flag_reported_net_inconsistent": "flag_net_inconsistent",
        "reported_net_minus_recomputed": "net_reported_minus_recalc",
    })

    # Force all quantitative columns to numeric before Stata export.
    non_numeric = {
        "registry_raw", "registry", "period_start", "period_end",
        "source_flow_type", "iso3", "entity_type"
    }
    for col in stata.columns:
        if col not in non_numeric:
            if stata[col].dtype == object:
                stata[col] = pd.to_numeric(stata[col], errors="coerce")

    stata.to_stata(
        OUT / "itl_country_panel_reporting_period.dta",
        write_index=False,
        version=118,
    )

    print("\n=== Extraction audit ===")
    print(audit_df.to_string(index=False))

    print("\n=== Country panel ===")
    print("Rows:", len(country))
    print("Countries:", country["iso3"].nunique())
    print("Report years:", sorted(country["report_year"].unique()))

    print("\nReported-net inconsistencies by report year:")
    print(
        country.groupby("report_year")["flag_reported_net_inconsistent"]
        .sum()
        .to_string()
    )

    print("\nFiles written to:", OUT)


if __name__ == "__main__":
    main()
