# -*- coding: utf-8 -*-
"""
Paris Gap-to-Target indicator construction
============================

Construction du nouvel indicateur Gap-to-Target pour l'Accord de Paris.

Cadre méthodologique appliqué
-----------------------------
1. Utiliser la base NDC harmonisée (First NDC / Second NDC).
2. Construire séparément les cibles unconditional et conditional.
3. Harmoniser les types de cibles vers un niveau d'émissions cible E_target :
   - fixed / explicit target level
   - base-year reduction
   - BAU reduction
   - intensity target
   - per-capita target
4. Définir le début de mise en oeuvre :
       s_ic = official implementation start year, si disponible;
              submission year, sinon.
   Définir l'année d'ancrage de la trajectoire :
       b_ic = s_ic - 1.
5. Construire la trajectoire linéaire compatible avec la cible à partir
   de l'émission observée l'année précédant le début de mise en oeuvre :
       E_path_itc = E_ib + ((t-b)/(T-b)) * (E_target_ic - E_ib)
6. Construire :
       Gap_abs = E_it - E_path_itc
       Gap     = 100 * (E_it - E_path_itc) / E_path_itc
       Gap_asinh = asinh(Gap_abs)
       Ambition = 100 * (E_ib - E_target_ic) / E_ib
7. BAU externe seulement lorsque nécessaire :
   - central : RCP6-SSP2, médiane OECD / PIK / IIASA
   - incertitude scénario : RCP4.5-SSP2 et RCP8.5-SSP2,
     médiane OECD / PIK / IIASA
   - sensibilité variante socio-économique/downscaling :
     OECD, PIK, IIASA séparément sous RCP6-SSP2

IMPORTANT
---------
- PMRCPBIE emissions: IPCM0EL / KYOTOGHGAR4 / GgCO2eq.
- PMRCPBIE socioeconomic scaling: GDPPPP and POP under SSP2 for
  intensity and per-capita target conversion.
- EDGAR national totals remain available as a diagnostic benchmark.
- Analytical observed emissions are read from the scope-harmonised bridge:
  economy-wide/all-sector targets use EDGAR national totals; explicit sectoral
  targets use the corresponding EDGAR IPCC-2006 sector union; FAOSTAT Forest
  land is added only when the audited NDC accounting perimeter requires it.
- Le module "intensity target" est volontairement conservateur : si les
  variables nécessaires ne sont pas identifiées sans ambiguïté, la cible
  reste manquante et l'observation est auditée au lieu d'être imputée.
"""

from __future__ import annotations

from pathlib import Path
import re
import warnings
import unicodedata

import numpy as np
import pandas as pd

# 0. CONFIGURATION

PROJECT_ROOT = Path(__file__).resolve().parents[2]
BASE_DIR = PROJECT_ROOT / "data" / "raw" / "paris"
PROCESSED_DIR = PROJECT_ROOT / "data" / "processed" / "paris"

NDC_FILE = PROCESSED_DIR / "NDC_Harmonized_Database.xlsx"
NDC_SHEET = "NDC_Harmonized"

# Cycle-specific observed-emissions perimeter produced by the Harmonizer.
SCOPE_EMISSIONS_FILE = PROCESSED_DIR / "NDC_Observed_Emissions_Scope_Harmonized.xlsx"
SCOPE_EMISSIONS_SHEET = "NDC_Scope_Observed"

PMR_FILE = BASE_DIR / "PMRCPBIE_04Feb20.zip"

# Le script cherche automatiquement le fichier EDGAR si le nom exact diffère.
EDGAR_CANDIDATES = [
    BASE_DIR / "EDGAR_AR5_GHG_1970_2024.xlsx",
    BASE_DIR / "EDGAR_AR5_GHG_1970_2024(1).xlsx",
    BASE_DIR / "EDGAR_AR5_GHG_1970_2024(2).xlsx",
]

# Facultatif pour le moment : sera mobilisé après audit définitif des
# variables d'intensité si nécessaire.
OWID_CANDIDATES = [
    BASE_DIR / "owid-co2-data.csv",
    BASE_DIR / "OWID.csv",
    BASE_DIR / "owid.csv",
]

OUTPUT_DIR = PROCESSED_DIR
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "Paris_GapToTarget_Database.xlsx"

# Observed Paris analysis window.
# The econometric / graph panels are strictly historical and therefore stop
# at the last available EDGAR year (2024). Every NDC country is represented
# on the common 2015-2024 calendar; Gap variables remain missing before the
# relevant NDC implementation start year.
ANALYSIS_START = 2015
ANALYSIS_END = 2024

# Additional replication / analysis outputs inspired by the previous workflow.
# The comprehensive workbook above is preserved unchanged as the main audit file.
OUT_MASTER_DTA = OUTPUT_DIR / "Paris_GapToTarget_Master.dta"
OUT_ANALYSIS_DTA = OUTPUT_DIR / "Paris_GapToTarget_Analysis.dta"
OUT_GRAPHS_DTA = OUTPUT_DIR / "Paris_GapToTarget_CountryGraphs.dta"

OUT_TARGETS_XLSX = OUTPUT_DIR / "Paris_NDC_Targets.xlsx"
OUT_TARGETS_DTA = OUTPUT_DIR / "Paris_NDC_Targets.dta"
OUT_TARGETS_CSV = OUTPUT_DIR / "Paris_NDC_Targets.csv"

OUT_AUDIT_XLSX = OUTPUT_DIR / "Paris_Target_Audit.xlsx"
OUT_AUDIT_DTA = OUTPUT_DIR / "Paris_Target_Audit.dta"

OUT_CYCLE_PANEL_XLSX = OUTPUT_DIR / "Paris_GapToTarget_Cycle_Panel.xlsx"

# Final econometric dataset: ISO3 + year + Gap + Ambition + Implementation + coverage classes.
OUT_FINAL_DTA = OUTPUT_DIR / "Paris_GapToTarget_Final.dta"
OUT_FINAL_XLSX = OUTPUT_DIR / "Paris_GapToTarget_Final.xlsx"
OUT_FINAL_CSV = OUTPUT_DIR / "Paris_GapToTarget_Final.csv"

# 1. OUTILS GENERAUX

def first_existing(paths):
    for p in paths:
        if p.exists():
            return p
    return None

def clean_text(x):
    """
    Conservative scalar text cleaner.

    If pandas returns a Series because a dataframe contains duplicate column
    names, keep the first non-missing value rather than evaluating pd.isna()
    on the whole Series (which raises "truth value of a Series is ambiguous").
    """
    if isinstance(x, pd.Series):
        vals = x.dropna()
        if vals.empty:
            return ""
        x = vals.iloc[0]

    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).replace("\xa0", " ")).strip()

def clean_col(x):
    return clean_text(x)

def normalize_country_name(x):
    """
    Normalisation uniquement pour l'appariement des noms de pays :
    accents, ponctuation, articles simples et espaces.
    """
    s = clean_text(x)
    s = "".join(
        c for c in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(c)
    )
    s = s.lower()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def to_num(x):
    # If duplicate columns cause pandas to return a Series instead of a scalar,
    # keep the first non-missing value defensively.
    if isinstance(x, pd.Series):
        vals = x.dropna()
        if vals.empty:
            return np.nan
        x = vals.iloc[0]

    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = clean_text(x).replace("%", "").replace(",", ".")
    s = s.replace("−", "-").replace("–", "-")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else np.nan

def pct_to_rate(x):
    v = to_num(x)
    if pd.isna(v):
        return np.nan
    # Les colonnes harmonisées sont en pourcentage (20 = 20%).
    return v / 100.0 if abs(v) > 1 else v

def year_num(x):
    v = to_num(x)
    if pd.isna(v):
        return np.nan
    y = int(round(v))
    return y if 1850 <= y <= 2100 else np.nan

def find_col(df, candidates, required=False):
    """
    Recherche robuste d'une colonne à partir de noms candidats.
    """
    norm = {clean_col(c).lower(): c for c in df.columns}

    for cand in candidates:
        key = clean_col(cand).lower()
        if key in norm:
            return norm[key]

    # recherche partielle prudente
    for cand in candidates:
        key = clean_col(cand).lower()
        hits = [c for c in df.columns if key in clean_col(c).lower()]
        if len(hits) == 1:
            return hits[0]

    if required:
        raise KeyError(
            f"Aucune colonne trouvée parmi : {candidates}\n"
            f"Colonnes disponibles : {list(df.columns)}"
        )
    return None

def safe_div(a, b):
    if pd.isna(a) or pd.isna(b) or b == 0:
        return np.nan
    return a / b

def prepare_for_stata_export(df):
    """
    Nettoyage conservateur pour export Stata :
    - noms uniques;
    - chaînes manquantes -> "";
    - booléens -> int8.
    """
    out = df.copy()
    out = out.loc[:, ~out.columns.duplicated()].copy()

    for c in out.columns:
        if pd.api.types.is_bool_dtype(out[c]):
            out[c] = out[c].astype("int8")
        elif out[c].dtype == "object" or pd.api.types.is_string_dtype(out[c]):
            out[c] = out[c].map(
                lambda x: "" if pd.isna(x) else str(x)
            ).astype(object)

    return out

def stata_safe_varnames(df):
    """
    Rend les noms compatibles Stata (<=32 caractères) en gardant
    une correspondance reproductible.
    """
    rename_map = {}
    used = set()

    for col in df.columns:
        original = str(col)
        base = re.sub(r"[^A-Za-z0-9_]", "_", original.strip())
        base = re.sub(r"_+", "_", base).strip("_")

        if not base:
            base = "var"
        if re.match(r"^[0-9]", base):
            base = "v_" + base

        candidate = base[:32]
        if candidate not in used:
            used.add(candidate)
            rename_map[original] = candidate
            continue

        k = 1
        while True:
            suffix = f"_{k}"
            candidate = base[:32-len(suffix)] + suffix
            if candidate not in used:
                used.add(candidate)
                rename_map[original] = candidate
                break
            k += 1

    return df.rename(columns=rename_map), rename_map

def safe_stata_labels(labels, cols):
    return {
        k: str(v).replace("CO₂", "CO2")[:80]
        for k, v in labels.items()
        if k in cols
    }

# 2. CHARGEMENT NDC

if not NDC_FILE.exists():
    raise FileNotFoundError(
        f"Base NDC harmonisée introuvable : {NDC_FILE}"
    )

ndc = pd.read_excel(NDC_FILE, sheet_name=NDC_SHEET, engine="openpyxl")
ndc.columns = [clean_col(c) for c in ndc.columns]

# Echantillon analytique First / Second NDC
if "analysis_ndc12" in ndc.columns:
    ndc = ndc[ndc["analysis_ndc12"].eq(1)].copy()
else:
    ndc_num_col = find_col(ndc, ["NDCs_number"], required=True)
    ndc = ndc[
        ndc[ndc_num_col].astype(str).isin(["First_NDC", "Second_NDC"])
    ].copy()

COUNTRY_COL = find_col(ndc, ["Country", "country"], required=True)
NDC_CYCLE_COL = find_col(ndc, ["NDCs_number"], required=True)
AUDIT_SCHEMA_VERSION = "20260828_DJI"
if "audit_schema_version" not in ndc or not ndc["audit_schema_version"].astype(str).eq(AUDIT_SCHEMA_VERSION).all():
    raise ValueError("Run the revised Harmonizer first: the NDC workbook has an incompatible audit version.")

ISO3_COL = find_col(
    ndc,
    [
        "ISO3", "iso3", "Country_code_A3", "country_code_a3",
        "Country Code", "country_code", "ISO"
    ],
    required=False
)

# La base NDC harmonisée peut ne pas contenir d'ISO3.
# Dans ce cas, l'appariement sera effectué plus bas à partir des noms
# de pays de la feuille EDGAR "TOTALS BY COUNTRY", avec une table
# d'alias explicite pour les différences de nomenclature.
if ISO3_COL is not None:
    ndc["iso3"] = (
        ndc[ISO3_COL]
        .astype("string")
        .str.strip()
        .str.upper()
    )
    ndc["iso3_match_source"] = "NDC_EXISTING_ISO3"
else:
    ndc["iso3"] = pd.Series(pd.NA, index=ndc.index, dtype="string")
    ndc["iso3_match_source"] = "PENDING_COUNTRY_NAME_MATCH"

# Variables méthodologiques déjà harmonisées
TARGET_METRIC_COL = find_col(ndc, ["target_metric_std"], required=True)
TARGET_SCOPE_COL = find_col(ndc, ["target_scope_std"], required=False)
TARGET_COVERAGE_COL = find_col(
    ndc, ["quantified_target_coverage_std"], required=False
)
TARGET_USABILITY_COL = find_col(
    ndc, ["target_usability_status"], required=False
)

MAIN_PCT_COL = find_col(ndc, ["main_target_pct_std"], required=False)
UNCOND_PCT_COL = find_col(ndc, ["unconditional_pct_std"], required=False)
COND_TOTAL_PCT_COL = find_col(
    ndc, ["conditional_total_pct_std"], required=False
)

ABS_VALUE_COL = find_col(
    ndc, ["main_target_value_ggco2e_std"], required=False
)
ABS_KIND_COL = find_col(
    ndc, ["main_target_value_kind_std"], required=False
)
UNCOND_ABS_VALUE_COL = find_col(
    ndc, ["unconditional_target_value_ggco2e_std"], required=False
)
COND_ABS_VALUE_COL = find_col(
    ndc, ["conditional_target_value_ggco2e_std"], required=False
)

BASE_YEAR_COL = find_col(ndc, ["base_year_std"], required=False)
TARGET_YEAR_COL = find_col(ndc, ["target_year_std"], required=True)
IMPL_START_COL = find_col(
    ndc, ["implementation_start_year"], required=False
)
SUBMISSION_YEAR_COL = find_col(
    ndc, ["submission_year_std"], required=True
)

PER_CAPITA_FLAG_COL = find_col(
    ndc, ["per_capita_target_flag"], required=False
)
CONDITIONALITY_COL = find_col(
    ndc, ["conditionality_std"], required=False
)

# Valeurs explicites éventuelles disponibles dans la base originale.
# Le script les repère seulement si le nom est non ambigu.
OFFICIAL_BAU_VALUE_COL = find_col(
    ndc,
    [
        "BAU_target_value_ggco2e",
        "bau_target_value_ggco2e",
        "BAU_value_ggco2e",
        "bau_value_ggco2e",
        "BAU emissions GgCO2eq",
        "BAU emissions"
    ],
    required=False
)

PER_CAPITA_TARGET_VALUE_COL = find_col(
    ndc,
    [
        "per_capita_target_value",
        "target_per_capita",
        "per_capita_target",
        "Target per capita"
    ],
    required=False
)

INTENSITY_REDUCTION_COL = find_col(
    ndc,
    [
        "intensity_target_pct",
        "intensity_reduction_pct",
        "Intensity target in percent"
    ],
    required=False
)

# Final audit metadata and special target fields.
AUDIT_DECISION_COL = find_col(ndc, ["audit_decision_std"], required=False)
AUDIT_REASON_COL = find_col(ndc, ["audit_reason_std"], required=False)
AUDIT_NOTE_COL = find_col(ndc, ["audit_note_std"], required=False)
HARMONIZATION_METHOD_COL = find_col(
    ndc, ["harmonization_method_std"], required=False
)
OFFICIAL_BAU_STD_COL = find_col(
    ndc, ["official_bau_value_ggco2e_std"], required=False
)
SPECIAL_TARGET_AMOUNT_COL = find_col(
    ndc, ["special_target_amount_ggco2e_std"], required=False
)
SPECIAL_UNCONDITIONAL_AMOUNT_COL = find_col(
    ndc, ["special_unconditional_target_amount_ggco2e_std"], required=False
)
SPECIAL_CONDITIONAL_AMOUNT_COL = find_col(
    ndc, ["special_conditional_target_amount_ggco2e_std"], required=False
)
INTERIM_TARGET_YEAR_COL = find_col(
    ndc, ["interim_target_year_std"], required=False
)
INTERIM_MAIN_PCT_COL = find_col(
    ndc, ["interim_target_pct_std"], required=False
)
INTERIM_UNCOND_PCT_COL = find_col(
    ndc, ["interim_unconditional_pct_std"], required=False
)
INTERIM_COND_PCT_COL = find_col(
    ndc, ["interim_conditional_total_pct_std"], required=False
)

# 3. EDGAR : EMISSIONS HISTORIQUES

EDGAR_FILE = first_existing(EDGAR_CANDIDATES)

if EDGAR_FILE is None:
    raise FileNotFoundError(
        "Fichier EDGAR introuvable dans Row Data."
    )

# L'entête réel de TOTALS BY COUNTRY se trouve à la ligne Excel 10
# => header=9 en indexation Python.
edgar = pd.read_excel(
    EDGAR_FILE,
    sheet_name="TOTALS BY COUNTRY",
    header=9,
    engine="openpyxl"
)

edgar.columns = [clean_col(c) for c in edgar.columns]

edgar_iso_col = find_col(edgar, ["Country_code_A3"], required=True)
edgar_name_col = find_col(edgar, ["Name"], required=True)
edgar_substance_col = find_col(edgar, ["Substance"], required=True)

edgar = edgar[
    edgar[edgar_substance_col].astype(str).eq("GWP_100_AR5_GHG")
].copy()

# Exclusion explicite des bunkers internationaux.
edgar = edgar[
    ~edgar[edgar_iso_col].astype(str).isin(["AIR", "SEA"])
].copy()

# 3.1 Appariement des pays NDC -> ISO3
#
# On privilégie toujours un ISO3 déjà présent dans la base NDC.
# S'il est absent, on construit l'ISO3 à partir du nom du pays.
# L'appariement automatique repose d'abord sur les noms EDGAR eux-mêmes,
# puis sur une table d'alias explicite pour les différences de nomenclature.

edgar_name_to_iso = (
    edgar[[edgar_name_col, edgar_iso_col]]
    .dropna()
    .drop_duplicates()
    .assign(
        _country_key=lambda x: x[edgar_name_col].map(normalize_country_name)
    )
    .set_index("_country_key")[edgar_iso_col]
    .astype(str)
    .str.strip()
    .str.upper()
    .to_dict()
)

# Alias correspondant aux principales différences entre Climate Watch/NDC,
# EDGAR et les codes ISO3 utilisés dans PMRCPBIE.
COUNTRY_ALIASES_ISO3 = {
    "andorra": "AND",
    "libya": "LBY",
    "liechtenstein": "LIE",
    "monaco": "MCO",
    "montenegro": "MNE",
    "san marino": "SMR",
    "serbia": "SRB",
    "south sudan": "SSD",
    "bahamas": "BHS",
    "bolivia": "BOL",
    "bolivia plurinational state of": "BOL",
    "brunei": "BRN",
    "brunei darussalam": "BRN",
    "cabo verde": "CPV",
    "cape verde": "CPV",
    "congo": "COG",
    "republic of the congo": "COG",
    "democratic republic of the congo": "COD",
    "democratic republic of congo": "COD",
    "dr congo": "COD",
    "cote d ivoire": "CIV",
    "ivory coast": "CIV",
    "czech republic": "CZE",
    "czechia": "CZE",
    "east timor": "TLS",
    "timor leste": "TLS",
    "eswatini": "SWZ",
    "swaziland": "SWZ",
    "gambia": "GMB",
    "the gambia": "GMB",
    "iran": "IRN",
    "iran islamic republic of": "IRN",
    "laos": "LAO",
    "lao people s democratic republic": "LAO",
    "micronesia": "FSM",
    "micronesia federated states of": "FSM",
    "moldova": "MDA",
    "republic of moldova": "MDA",
    "north korea": "PRK",
    "democratic people s republic of korea": "PRK",
    "south korea": "KOR",
    "republic of korea": "KOR",
    "north macedonia": "MKD",
    "macedonia": "MKD",
    "palestine": "PSE",
    "state of palestine": "PSE",
    "russia": "RUS",
    "russian federation": "RUS",
    "syria": "SYR",
    "syrian arab republic": "SYR",
    "taiwan": "TWN",
    "taiwan province of china": "TWN",
    "tanzania": "TZA",
    "united republic of tanzania": "TZA",
    "turkey": "TUR",
    "turkiye": "TUR",
    "united kingdom": "GBR",
    "united kingdom of great britain and northern ireland": "GBR",
    "united states": "USA",
    "united states of america": "USA",
    "venezuela": "VEN",
    "venezuela bolivarian republic of": "VEN",
    "vietnam": "VNM",
    "viet nam": "VNM",
    "vatican city": "VAT",
    "holy see": "VAT",
    "united arab emirates": "ARE",
    "uae": "ARE",
    "saint kitts and nevis": "KNA",
    "saint lucia": "LCA",
    "saint vincent and the grenadines": "VCT",
    "sao tome and principe": "STP",
    "sao tome principe": "STP",
    "kosovo": "XKX",
}

def country_name_to_iso3(country_name):
    key = normalize_country_name(country_name)

    # 1. correspondance directe avec la nomenclature EDGAR
    if key in edgar_name_to_iso:
        return edgar_name_to_iso[key], "EDGAR_NAME_MATCH"

    # 2. alias explicite
    if key in COUNTRY_ALIASES_ISO3:
        return COUNTRY_ALIASES_ISO3[key], "MANUAL_ALIAS"

    return pd.NA, "UNRESOLVED"

if ISO3_COL is None:
    country_matches = ndc[COUNTRY_COL].map(country_name_to_iso3)

    ndc["iso3"] = pd.Series(
        [x[0] for x in country_matches],
        index=ndc.index,
        dtype="string"
    )
    ndc["iso3_match_source"] = [
        x[1] for x in country_matches
    ]

else:
    # Remplir seulement les ISO3 manquants/vides, sans écraser ceux
    # déjà présents dans la base harmonisée.
    missing_iso = (
        ndc["iso3"].isna()
        | ndc["iso3"].astype("string").str.strip().isin(["", "<NA>", "NAN"])
    )

    if missing_iso.any():
        country_matches = ndc.loc[
            missing_iso, COUNTRY_COL
        ].map(country_name_to_iso3)

        ndc.loc[missing_iso, "iso3"] = [
            x[0] for x in country_matches
        ]
        ndc.loc[missing_iso, "iso3_match_source"] = [
            x[1] for x in country_matches
        ]

# Normalisation finale
ndc["iso3"] = (
    ndc["iso3"]
    .astype("string")
    .str.strip()
    .str.upper()
)

unresolved_countries = sorted(
    ndc.loc[
        ndc["iso3"].isna()
        | ndc["iso3"].isin(["", "<NA>", "NAN"]),
        COUNTRY_COL
    ]
    .dropna()
    .astype(str)
    .unique()
)

print("\n" + "=" * 80)
print("APPARIEMENT DES PAYS NDC -> ISO3")
print("=" * 80)
print(
    ndc["iso3_match_source"]
    .value_counts(dropna=False)
    .to_string()
)

if unresolved_countries:
    print("\nPays NDC non résolus automatiquement :")
    for country in unresolved_countries:
        print(f"  - {country}")
    print(
        "\nCes lignes seront conservées dans l'audit mais aucune cible "
        "ne sera calculée tant que leur ISO3 n'est pas résolu."
    )
else:
    print("\nTous les pays NDC ont été appariés à un code ISO3.")

year_cols_edgar = [
    c for c in edgar.columns
    if re.fullmatch(r"Y_(19|20)\d{2}", str(c))
]

edgar_long = edgar[
    [edgar_iso_col, edgar_name_col] + year_cols_edgar
].melt(
    id_vars=[edgar_iso_col, edgar_name_col],
    var_name="year",
    value_name="E_obs"
)

edgar_long["iso3"] = (
    edgar_long[edgar_iso_col]
    .astype("string")
    .str.strip()
    .str.upper()
)
edgar_long["year"] = (
    edgar_long["year"]
    .str.replace("Y_", "", regex=False)
    .astype(int)
)
edgar_long["E_obs"] = pd.to_numeric(
    edgar_long["E_obs"], errors="coerce"
)

edgar_lookup = (
    edgar_long
    .dropna(subset=["iso3", "year"])
    .set_index(["iso3", "year"])["E_obs"]
    .to_dict()
)

# 3.2 Emissions observées harmonisées au périmètre de chaque NDC
# Key principle: the same emissions perimeter must be used for (i) the declared
# base year, (ii) the trajectory anchor b=s-1, and (iii) every observed year t.
# National EDGAR totals are retained only as a backward-compatible fallback when
# no country-cycle scope record exists at all.  A scope record whose value is
# missing (e.g. required FAO proxy unavailable) remains missing and is NOT
# silently replaced by the national total.
if not SCOPE_EMISSIONS_FILE.exists():
    raise FileNotFoundError(
        f"Fichier de périmètre harmonisé introuvable : {SCOPE_EMISSIONS_FILE}. "
        "Exécuter d'abord le NDC Harmonizer."
    )

scope_obs = pd.read_excel(
    SCOPE_EMISSIONS_FILE,
    sheet_name=SCOPE_EMISSIONS_SHEET,
    engine="openpyxl",
)
scope_obs.columns = [clean_col(c) for c in scope_obs.columns]
if "audit_schema_version" not in scope_obs or not scope_obs["audit_schema_version"].astype(str).eq(AUDIT_SCHEMA_VERSION).all():
    raise ValueError("The observed-emissions scope file is stale; regenerate both Harmonizer workbooks together.")

_scope_required = [
    "iso3", "NDCs_number", "year", "observed_emissions_scope_ggco2e"
]
for _c in _scope_required:
    if _c not in scope_obs.columns:
        raise KeyError(f"Colonne de périmètre manquante : {_c}")

scope_obs["iso3"] = scope_obs["iso3"].astype("string").str.strip().str.upper()
scope_obs["NDCs_number"] = scope_obs["NDCs_number"].astype("string").str.strip()
scope_obs["year"] = pd.to_numeric(scope_obs["year"], errors="coerce")
scope_obs["observed_emissions_scope_ggco2e"] = pd.to_numeric(
    scope_obs["observed_emissions_scope_ggco2e"], errors="coerce"
)

scope_emissions_lookup = (
    scope_obs
    .dropna(subset=["iso3", "NDCs_number", "year"])
    .set_index(["iso3", "NDCs_number", "year"])["observed_emissions_scope_ggco2e"]
    .to_dict()
)
# Country-cycle availability is used to distinguish a genuinely absent scope
# series from a present-but-missing observation.
scope_cycle_set = set(
    (str(r.iso3), str(r.NDCs_number))
    for r in scope_obs.dropna(subset=["iso3", "NDCs_number"]).itertuples()
)

def observed_emissions_for_cycle(iso3, ndc_cycle, year):
    iso = clean_text(iso3).upper()
    cyc = clean_text(ndc_cycle)
    y = year_num(year)
    if not iso or not cyc or pd.isna(y):
        return np.nan
    key = (iso, cyc, int(y))
    if (iso, cyc) in scope_cycle_set:
        # Important: if the harmonised scope exists but this year is missing,
        # preserve NaN rather than substituting a national total.
        return scope_emissions_lookup.get(key, np.nan)
    return edgar_lookup.get((iso, int(y)), np.nan)

# 4. PMRCPBIE : BAU, PIB ET POPULATION

if not PMR_FILE.exists():
    raise FileNotFoundError(
        f"Fichier PMRCPBIE introuvable : {PMR_FILE}"
    )

pmr = pd.read_csv(PMR_FILE, low_memory=False)
pmr.columns = [clean_col(c) for c in pmr.columns]

required_pmr = ["source", "scenario", "country", "category", "entity", "unit"]
for c in required_pmr:
    if c not in pmr.columns:
        raise KeyError(f"Colonne PMRCPBIE manquante : {c}")

pmr["country"] = (
    pmr["country"].astype("string").str.strip().str.upper()
)

year_cols_pmr = [
    c for c in pmr.columns
    if str(c).isdigit() and 1850 <= int(c) <= 2100
]

# 4.1 Emissions BAU

SCENARIO_MAP = {
    "main_oecd": "RCP6SSP2OECD",
    "main_pik": "RCP6SSP2PIK",
    "main_iiasa": "RCP6SSP2IIASA",

    "rcp45_oecd": "RCP45SSP2OECD",
    "rcp45_pik": "RCP45SSP2PIK",
    "rcp45_iiasa": "RCP45SSP2IIASA",

    "rcp85_oecd": "RCP85SSP2OECD",
    "rcp85_pik": "RCP85SSP2PIK",
    "rcp85_iiasa": "RCP85SSP2IIASA",
}

pmr_ghg = pmr[
    (pmr["category"] == "IPCM0EL")
    & (pmr["entity"] == "KYOTOGHGAR4")
    & (pmr["unit"] == "GgCO2eq")
    & (pmr["scenario"].isin(SCENARIO_MAP.values()))
].copy()

ghg_long = pmr_ghg[
    ["scenario", "country"] + year_cols_pmr
].melt(
    id_vars=["scenario", "country"],
    var_name="year",
    value_name="bau_value"
)

ghg_long["year"] = ghg_long["year"].astype(int)
ghg_long["bau_value"] = pd.to_numeric(
    ghg_long["bau_value"], errors="coerce"
)

ghg_wide = (
    ghg_long
    .pivot_table(
        index=["country", "year"],
        columns="scenario",
        values="bau_value",
        aggfunc="first"
    )
    .reset_index()
)

# Créer les colonnes courtes
reverse_map = {v: k for k, v in SCENARIO_MAP.items()}
ghg_wide = ghg_wide.rename(columns=reverse_map)

for c in reverse_map.values():
    if c not in ghg_wide.columns:
        ghg_wide[c] = np.nan

# Nombre de variantes disponibles par RCP
ghg_wide["bau_n_models_rcp6"] = ghg_wide[
    ["main_oecd", "main_pik", "main_iiasa"]
].notna().sum(axis=1)

ghg_wide["bau_n_models_rcp45"] = ghg_wide[
    ["rcp45_oecd", "rcp45_pik", "rcp45_iiasa"]
].notna().sum(axis=1)

ghg_wide["bau_n_models_rcp85"] = ghg_wide[
    ["rcp85_oecd", "rcp85_pik", "rcp85_iiasa"]
].notna().sum(axis=1)

# Médiane disponible.
# L'audit conserve le nombre de variantes afin de distinguer 3, 2, 1 modèles.
ghg_wide["bau_rcp6_median"] = ghg_wide[
    ["main_oecd", "main_pik", "main_iiasa"]
].median(axis=1, skipna=True)

ghg_wide["bau_rcp45_median"] = ghg_wide[
    ["rcp45_oecd", "rcp45_pik", "rcp45_iiasa"]
].median(axis=1, skipna=True)

ghg_wide["bau_rcp85_median"] = ghg_wide[
    ["rcp85_oecd", "rcp85_pik", "rcp85_iiasa"]
].median(axis=1, skipna=True)

bau_lookup = (
    ghg_wide
    .set_index(["country", "year"])
    .to_dict(orient="index")
)

# 4.2 Population and GDP PPP — SSP2, model/scenario resolved
#
# Intensity and per-capita targets are converted using SCALE RATIOS:
#
#   Intensity:
#       E_T = E_base * (1-r) * GDP_T / GDP_base
#
#   Per capita:
#       E_T = E_base * (1-r) * POP_T / POP_base
#
# Using the same PMRCPBIE series at base and target years makes the
# socioeconomic units cancel. This avoids mixing historical OWID GDP
# units with PMRCPBIE GDP units.
#
# Central conversion:
#   RCP6-SSP2 median across OECD / PIK / IIASA.
#
# Sensitivity:
#   RCP4.5-SSP2 median, RCP8.5-SSP2 median,
#   and OECD / PIK / IIASA separately under RCP6-SSP2.

def build_socio_lookup(entity, category, unit, convert_factor=1.0):
    tmp = pmr[
        (pmr["category"] == category)
        & (pmr["entity"] == entity)
        & (pmr["unit"] == unit)
        & (pmr["scenario"].isin(SCENARIO_MAP.values()))
    ].copy()

    long = tmp[
        ["scenario", "country"] + year_cols_pmr
    ].melt(
        id_vars=["scenario", "country"],
        var_name="year",
        value_name="value"
    )

    long["year"] = long["year"].astype(int)
    long["value"] = (
        pd.to_numeric(long["value"], errors="coerce")
        * float(convert_factor)
    )

    wide = (
        long
        .pivot_table(
            index=["country", "year"],
            columns="scenario",
            values="value",
            aggfunc="first"
        )
        .reset_index()
    )

    rev = {v: k for k, v in SCENARIO_MAP.items()}
    wide = wide.rename(columns=rev)

    for c in rev.values():
        if c not in wide.columns:
            wide[c] = np.nan

    wide["main_median"] = wide[
        ["main_oecd", "main_pik", "main_iiasa"]
    ].median(axis=1, skipna=True)

    wide["rcp45_median"] = wide[
        ["rcp45_oecd", "rcp45_pik", "rcp45_iiasa"]
    ].median(axis=1, skipna=True)

    wide["rcp85_median"] = wide[
        ["rcp85_oecd", "rcp85_pik", "rcp85_iiasa"]
    ].median(axis=1, skipna=True)

    wide["n_models_main"] = wide[
        ["main_oecd", "main_pik", "main_iiasa"]
    ].notna().sum(axis=1)

    return wide.set_index(["country", "year"]).to_dict(orient="index")

# POP: ThousandPers -> persons. The ratio POP_T / POP_base is invariant
# to this conversion, but storing persons makes audit outputs intuitive.
pop_lookup_full = build_socio_lookup(
    entity="POP",
    category="DEMOGR",
    unit="ThousandPers",
    convert_factor=1000.0
)

# GDP PPP: Million2011GKD. We use only ratios, so the unit cancels.
gdp_lookup_full = build_socio_lookup(
    entity="GDPPPP",
    category="ECO",
    unit="Million2011GKD",
    convert_factor=1.0
)

def socio_values(lookup, iso3, year):
    if pd.isna(year):
        return {}
    return lookup.get((iso3, int(year)), {})

def ratio_or_nan(num, den):
    if pd.isna(num) or pd.isna(den) or den == 0:
        return np.nan
    return num / den

def socio_scale_bundle(lookup, iso3, base_year, target_year):
    """
    Return scale factors X_T / X_base for the central specification,
    scenario uncertainty, and model sensitivity.
    """
    base = socio_values(lookup, iso3, base_year)
    targ = socio_values(lookup, iso3, target_year)

    return {
        "main": ratio_or_nan(
            targ.get("main_median", np.nan),
            base.get("main_median", np.nan)
        ),
        "rcp45": ratio_or_nan(
            targ.get("rcp45_median", np.nan),
            base.get("rcp45_median", np.nan)
        ),
        "rcp85": ratio_or_nan(
            targ.get("rcp85_median", np.nan),
            base.get("rcp85_median", np.nan)
        ),
        "oecd": ratio_or_nan(
            targ.get("main_oecd", np.nan),
            base.get("main_oecd", np.nan)
        ),
        "pik": ratio_or_nan(
            targ.get("main_pik", np.nan),
            base.get("main_pik", np.nan)
        ),
        "iiasa": ratio_or_nan(
            targ.get("main_iiasa", np.nan),
            base.get("main_iiasa", np.nan)
        ),
        "n_models": targ.get("n_models_main", np.nan),
    }

# 5. PREPARATION DES NDC

def get_row_year(row, colname):
    if colname is None:
        return np.nan
    return year_num(row.get(colname))

def get_row_num(row, colname):
    if colname is None:
        return np.nan
    return to_num(row.get(colname))

ndc["target_year"] = ndc[TARGET_YEAR_COL].map(year_num)
ndc["submission_year"] = ndc[SUBMISSION_YEAR_COL].map(year_num)

if IMPL_START_COL is not None:
    ndc["implementation_start_year_final"] = (
        ndc[IMPL_START_COL].map(year_num)
    )
else:
    ndc["implementation_start_year_final"] = np.nan

# Règle méthodologique validée :
# année officielle de début si disponible, sinon année de soumission.
ndc["start_year"] = (
    ndc["implementation_start_year_final"]
    .combine_first(ndc["submission_year"])
)

# Année d'ancrage de la trajectoire :
# l'année immédiatement antérieure au début officiel de mise en oeuvre.
ndc["trajectory_baseline_year"] = np.where(
    ndc["start_year"].notna(),
    ndc["start_year"] - 1,
    np.nan
)

if BASE_YEAR_COL is not None:
    ndc["base_year"] = ndc[BASE_YEAR_COL].map(year_num)
else:
    ndc["base_year"] = np.nan

# 6. CONSTRUCTION DES CIBLES

def get_target_rates(row):
    """
    Retourne les taux de réduction unconditional et conditional.

    On ne fusionne jamais les deux composantes.
    """
    main = (
        pct_to_rate(row.get(MAIN_PCT_COL))
        if MAIN_PCT_COL is not None else np.nan
    )
    uncond = (
        pct_to_rate(row.get(UNCOND_PCT_COL))
        if UNCOND_PCT_COL is not None else np.nan
    )
    cond = (
        pct_to_rate(row.get(COND_TOTAL_PCT_COL))
        if COND_TOTAL_PCT_COL is not None else np.nan
    )

    cond_structure = (
        clean_text(row.get(CONDITIONALITY_COL)).upper()
        if CONDITIONALITY_COL is not None else ""
    )

    # Unconditional only : le main target représente l'inconditionnel
    if pd.isna(uncond) and cond_structure == "UNCONDITIONAL_ONLY":
        uncond = main

    # Conditional only : le main target représente le conditionnel
    if pd.isna(cond) and cond_structure == "CONDITIONAL_ONLY":
        cond = main

    # Cas sans structure disponible : main est conservé comme target "main"
    return main, uncond, cond

def direct_absolute_target(row):
    if ABS_VALUE_COL is None:
        return np.nan

    value = to_num(row.get(ABS_VALUE_COL))
    if pd.isna(value):
        return np.nan

    kind = (
        clean_text(row.get(ABS_KIND_COL)).upper()
        if ABS_KIND_COL is not None else ""
    )

    if kind == "EMISSIONS_LEVEL_OR_CAP":
        return value

    return np.nan

def bau_reference(row, iso3, target_year):
    """
    Retourne les benchmarks BAU du pays-année cible.

    Priorité :
    1. BAU officiel quantifié de la NDC si identifiable ;
    2. sinon PMRCPBIE.
    """
    official = np.nan

    # The audited official or implied BAU value has highest priority.
    if OFFICIAL_BAU_STD_COL is not None:
        official = to_num(row.get(OFFICIAL_BAU_STD_COL))

    # Fall back to a pre-existing official BAU field from the raw/harmonized source.
    if pd.isna(official) and OFFICIAL_BAU_VALUE_COL is not None:
        official = to_num(row.get(OFFICIAL_BAU_VALUE_COL))

    if pd.notna(official):
        return {
            "bau_source": "NDC_OFFICIAL",
            "bau_main": official,
            "bau_rcp45": np.nan,
            "bau_rcp85": np.nan,
            "bau_oecd": np.nan,
            "bau_pik": np.nan,
            "bau_iiasa": np.nan,
            "bau_n_models": np.nan,
        }

    vals = bau_lookup.get((iso3, int(target_year)), {})

    return {
        "bau_source": "PMRCPBIE",
        "bau_main": vals.get("bau_rcp6_median", np.nan),
        "bau_rcp45": vals.get("bau_rcp45_median", np.nan),
        "bau_rcp85": vals.get("bau_rcp85_median", np.nan),
        "bau_oecd": vals.get("main_oecd", np.nan),
        "bau_pik": vals.get("main_pik", np.nan),
        "bau_iiasa": vals.get("main_iiasa", np.nan),
        "bau_n_models": vals.get("bau_n_models_rcp6", np.nan),
    }

def target_from_rate(reference, rate):
    if pd.isna(reference) or pd.isna(rate):
        return np.nan
    return reference * (1.0 - rate)

def target_from_absolute_reduction(reference, amount):
    if pd.isna(reference) or pd.isna(amount):
        return np.nan
    return reference - amount

def target_from_scaled_relative(E_base, rate, scale):
    """
    Convert a relative intensity/per-capita target into an absolute
    emissions target:
        E_T = E_base * (1-rate) * scale
    where scale is GDP_T/GDP_base or POP_T/POP_base.
    """
    if pd.isna(E_base) or pd.isna(rate) or pd.isna(scale):
        return np.nan
    return E_base * (1.0 - rate) * scale

def build_targets(row):
    iso3 = clean_text(row["iso3"]).upper()
    metric = clean_text(row[TARGET_METRIC_COL]).upper()
    scope = (
        clean_text(row.get(TARGET_SCOPE_COL)).upper()
        if TARGET_SCOPE_COL is not None else ""
    )
    coverage = (
        clean_text(row.get(TARGET_COVERAGE_COL)).upper()
        if TARGET_COVERAGE_COL is not None else ""
    )
    usability = (
        clean_text(row.get(TARGET_USABILITY_COL)).upper()
        if TARGET_USABILITY_COL is not None else ""
    )

    T = row["target_year"]
    b = row["base_year"]

    result = {
        "target_method": np.nan,
        "target_status": "NOT_BUILT",
        "target_note": np.nan,

        "E_target_main": np.nan,
        "E_target_U": np.nan,
        "E_target_C": np.nan,

        "E_target_rcp45": np.nan,
        "E_target_rcp85": np.nan,
        "E_target_oecd": np.nan,
        "E_target_pik": np.nan,
        "E_target_iiasa": np.nan,

        "bau_source": np.nan,
        "bau_main": np.nan,
        "bau_n_models": np.nan,

# Optional intermediate target for piecewise trajectories.
        "interim_target_year": np.nan,
        "E_interim_main": np.nan,
        "E_interim_U": np.nan,
        "E_interim_C": np.nan,
        "E_interim_rcp45": np.nan,
        "E_interim_rcp85": np.nan,
        "E_interim_oecd": np.nan,
        "E_interim_pik": np.nan,
        "E_interim_iiasa": np.nan,

        # Audit trail propagated to target outputs.
        "audit_decision": (
            row.get(AUDIT_DECISION_COL)
            if AUDIT_DECISION_COL is not None else np.nan
        ),
        "audit_reason": (
            row.get(AUDIT_REASON_COL)
            if AUDIT_REASON_COL is not None else np.nan
        ),
        "harmonization_method": (
            row.get(HARMONIZATION_METHOD_COL)
            if HARMONIZATION_METHOD_COL is not None else np.nan
        ),
    }

    availability = clean_text(row.get("target_availability_std", "REPORTED"))
    if availability in {"NO_SUBMISSION", "NO_QUANTIFIED_GHG_TARGET", "OUTSIDE_ANALYSIS_WINDOW"}:
        result["target_status"] = availability
        result["target_note"] = "No target constructed for this cycle: " + availability
        return result

    if not iso3 or pd.isna(T):
        result["target_note"] = "Missing ISO3 or target year"
        return result

    audit_decision = (
        clean_text(row.get(AUDIT_DECISION_COL)).upper()
        if AUDIT_DECISION_COL is not None else ""
    )
    audit_reason = (
        clean_text(row.get(AUDIT_REASON_COL)).upper()
        if AUDIT_REASON_COL is not None else ""
    )
    special_method = (
        clean_text(row.get(HARMONIZATION_METHOD_COL)).upper()
        if HARMONIZATION_METHOD_COL is not None else ""
    )

    # Manually audited decisions have priority over generic sector rules.
    if audit_decision == "EXCLUDE":
        result["target_status"] = "AUDIT_EXCLUDED"
        result["target_method"] = "NOT_BUILT_AUDIT_EXCLUSION"
        result["target_note"] = audit_reason or "Manual audit exclusion"
        return result

    # Generic fallback for non-audited sector-specific targets.
    if (
        audit_decision != "INCLUDE"
        and coverage in {"SINGLE_SECTOR_GHG", "PARTIAL_SECTORAL_GHG"}
    ):
        result["target_status"] = "EXCLUDED_SECTOR_SCOPE"
        result["target_note"] = (
            "Quantified target is sector-specific and has not been "
            "manually validated for harmonised inclusion."
        )
        return result

    if usability == "NOT_YET_OBSERVABLE":
        result["target_status"] = "NOT_YET_OBSERVABLE"
        result["target_note"] = (
            "NDC implementation begins after the last observed EDGAR year"
        )
        # Still build target if possible below, because it is useful for audit.

    main_rate, u_rate, c_rate = get_target_rates(row)

    # Read the harmonized absolute-value interpretation once, before any
    # branch uses it. In V11 this was first assigned inside the BAU branch,
    # which caused an UnboundLocalError for fixed-level targets.
    abs_kind = (
        clean_text(row.get(ABS_KIND_COL)).upper()
        if ABS_KIND_COL is not None else ""
    )

    # A0. Audited special architectures

    # Absolute annual abatement at target year relative to BAU:
    #   E_target(T) = BAU(T) - A(T)
    # Used for audited cases such as Saudi Arabia and El Salvador.
    if special_method == "ABSOLUTE_ANNUAL_ABATEMENT_BAU":
        amount = (
            to_num(row.get(SPECIAL_TARGET_AMOUNT_COL))
            if SPECIAL_TARGET_AMOUNT_COL is not None else np.nan
        )
        amount_u = (
            to_num(row.get(SPECIAL_UNCONDITIONAL_AMOUNT_COL))
            if SPECIAL_UNCONDITIONAL_AMOUNT_COL is not None else np.nan
        )
        amount_c = (
            to_num(row.get(SPECIAL_CONDITIONAL_AMOUNT_COL))
            if SPECIAL_CONDITIONAL_AMOUNT_COL is not None else np.nan
        )
        bau = bau_reference(row, iso3, T)

        result["bau_source"] = bau["bau_source"]
        result["bau_main"] = bau["bau_main"]
        result["bau_n_models"] = bau["bau_n_models"]
        result["target_method"] = "ABSOLUTE_ANNUAL_ABATEMENT_BAU"

        result["E_target_main"] = target_from_absolute_reduction(
            bau["bau_main"], amount
        )
        result["E_target_U"] = target_from_absolute_reduction(
            bau["bau_main"], amount_u
        )
        result["E_target_C"] = target_from_absolute_reduction(
            bau["bau_main"], amount_c
        )
        if bau["bau_source"] == "PMRCPBIE":
            result["E_target_rcp45"] = target_from_absolute_reduction(
                bau["bau_rcp45"], amount
            )
            result["E_target_rcp85"] = target_from_absolute_reduction(
                bau["bau_rcp85"], amount
            )
            result["E_target_oecd"] = target_from_absolute_reduction(
                bau["bau_oecd"], amount
            )
            result["E_target_pik"] = target_from_absolute_reduction(
                bau["bau_pik"], amount
            )
            result["E_target_iiasa"] = target_from_absolute_reduction(
                bau["bau_iiasa"], amount
            )

        built = pd.notna(result["E_target_main"])
        result["target_status"] = (
            "NOT_YET_OBSERVABLE"
            if usability == "NOT_YET_OBSERVABLE" and built
            else ("OK" if built else "NEEDS_REVIEW")
        )
        if not built:
            result["target_note"] = (
                "Audited annual abatement target lacks a usable BAU or amount."
            )
        return result

    # Cumulative reduction trajectory. The total abatement A_cum over the
    # implementation years is distributed as a linear ramp in annual
    # reductions from 1 unit in the first year to N units in the target year.
    # The resulting terminal annual abatement is:
    #   A_T = A_cum * N / sum(1,...,N)
    # The path itself is reconstructed below using the BAU annual trajectory.
    if special_method == "CUMULATIVE_REDUCTION_LINEAR_RAMP":
        total_amount = (
            to_num(row.get(SPECIAL_TARGET_AMOUNT_COL))
            if SPECIAL_TARGET_AMOUNT_COL is not None else np.nan
        )
        s_local = year_num(row.get("start_year"))
        if pd.isna(s_local) or pd.isna(T) or pd.isna(total_amount):
            result["target_status"] = "NEEDS_REVIEW"
            result["target_method"] = "CUMULATIVE_REDUCTION_LINEAR_RAMP"
            result["target_note"] = "Missing start year, target year, or cumulative amount."
            return result

        n_years = int(T - s_local + 1)
        denom = n_years * (n_years + 1) / 2.0
        terminal_amount = total_amount * n_years / denom

        bau = bau_reference(row, iso3, T)
        result["bau_source"] = bau["bau_source"]
        result["bau_main"] = bau["bau_main"]
        result["bau_n_models"] = bau["bau_n_models"]
        result["target_method"] = "CUMULATIVE_REDUCTION_LINEAR_RAMP"
        result["E_target_main"] = target_from_absolute_reduction(
            bau["bau_main"], terminal_amount
        )

        if bau["bau_source"] == "PMRCPBIE":
            result["E_target_rcp45"] = target_from_absolute_reduction(
                bau["bau_rcp45"], terminal_amount
            )
            result["E_target_rcp85"] = target_from_absolute_reduction(
                bau["bau_rcp85"], terminal_amount
            )
            result["E_target_oecd"] = target_from_absolute_reduction(
                bau["bau_oecd"], terminal_amount
            )
            result["E_target_pik"] = target_from_absolute_reduction(
                bau["bau_pik"], terminal_amount
            )
            result["E_target_iiasa"] = target_from_absolute_reduction(
                bau["bau_iiasa"], terminal_amount
            )

        built = pd.notna(result["E_target_main"])
        result["target_status"] = "OK" if built else "NEEDS_REVIEW"
        if not built:
            result["target_note"] = "Missing BAU for cumulative reduction trajectory."
        return result

    # A. Niveau absolu directement déclaré
    direct = direct_absolute_target(row)

    if abs_kind == "CUMULATIVE_EMISSIONS_BUDGET":
        result["target_status"] = "NEEDS_REVIEW"
        result["target_method"] = "CUMULATIVE_BUDGET_NOT_TERMINAL_LEVEL"
        result["target_note"] = (
            "Cumulative emissions budget cannot be used as the terminal "
            "annual emissions level for the Gap trajectory."
        )
        return result

    if pd.notna(direct):
        result["target_method"] = "DIRECT_ABSOLUTE_LEVEL"
        result["E_target_main"] = direct

        u_abs = (
            to_num(row.get(UNCOND_ABS_VALUE_COL))
            if UNCOND_ABS_VALUE_COL is not None else np.nan
        )
        c_abs = (
            to_num(row.get(COND_ABS_VALUE_COL))
            if COND_ABS_VALUE_COL is not None else np.nan
        )

        if pd.notna(u_abs):
            result["E_target_U"] = u_abs
        if pd.notna(c_abs):
            result["E_target_C"] = c_abs

        result["target_status"] = (
            "NOT_YET_OBSERVABLE"
            if usability == "NOT_YET_OBSERVABLE"
            else "OK"
        )
        return result

    # B. Base-year target
    if metric == "BASE_YEAR_REDUCTION":
        if pd.isna(b):
            result["target_note"] = "Missing base year"
            return result

        E_base = observed_emissions_for_cycle(iso3, row.get(NDC_CYCLE_COL), b)

        if pd.isna(E_base):
            result["target_note"] = "Missing EDGAR emissions in base year"
            return result

        result["target_method"] = "BASE_YEAR_EDGAR_AR5"

        result["E_target_main"] = target_from_rate(E_base, main_rate)
        result["E_target_U"] = target_from_rate(E_base, u_rate)
        result["E_target_C"] = target_from_rate(E_base, c_rate)

        built_any = any(pd.notna(result[k]) for k in [
            "E_target_main", "E_target_U", "E_target_C"
        ])

        if built_any:
            result["target_status"] = (
                "NOT_YET_OBSERVABLE"
                if usability == "NOT_YET_OBSERVABLE"
                else "OK"
            )
        else:
            result["target_status"] = "NEEDS_REVIEW"
            result["target_note"] = (
                "Base-year architecture identified but no usable "
                "percentage target was available"
            )
        return result

    # C. BAU target
    if metric in {"BAU_REDUCTION", "BAU_PER_CAPITA_TARGET"}:
        bau = bau_reference(row, iso3, T)

        result["bau_source"] = bau["bau_source"]
        result["bau_main"] = bau["bau_main"]
        result["bau_n_models"] = bau["bau_n_models"]

        if pd.isna(bau["bau_main"]):
            result["target_note"] = "Missing usable BAU at target year"
            return result

        main_amount = (
            to_num(row.get(ABS_VALUE_COL))
            if ABS_VALUE_COL is not None else np.nan
        )
        u_amount = (
            to_num(row.get(UNCOND_ABS_VALUE_COL))
            if UNCOND_ABS_VALUE_COL is not None else np.nan
        )
        c_amount_raw = (
            to_num(row.get(COND_ABS_VALUE_COL))
            if COND_ABS_VALUE_COL is not None else np.nan
        )

        # BAU target stated as an absolute reduction amount
        if abs_kind == "ABSOLUTE_REDUCTION_AMOUNT" and pd.notna(main_amount):
            result["target_method"] = (
                "BAU_ABSOLUTE_REDUCTION_NDC_OFFICIAL"
                if bau["bau_source"] == "NDC_OFFICIAL"
                else "BAU_ABSOLUTE_REDUCTION_PMR"
            )

            result["E_target_main"] = target_from_absolute_reduction(
                bau["bau_main"], main_amount
            )

            if pd.notna(u_amount):
                result["E_target_U"] = target_from_absolute_reduction(
                    bau["bau_main"], u_amount
                )

            # When main amount is the maximum/total target, it is the
            # appropriate conditional target level. The raw conditional
            # amount may be only the additional component.
            if pd.notna(main_amount):
                result["E_target_C"] = target_from_absolute_reduction(
                    bau["bau_main"], main_amount
                )

            if bau["bau_source"] == "PMRCPBIE":
                result["E_target_rcp45"] = target_from_absolute_reduction(
                    bau["bau_rcp45"], main_amount
                )
                result["E_target_rcp85"] = target_from_absolute_reduction(
                    bau["bau_rcp85"], main_amount
                )
                result["E_target_oecd"] = target_from_absolute_reduction(
                    bau["bau_oecd"], main_amount
                )
                result["E_target_pik"] = target_from_absolute_reduction(
                    bau["bau_pik"], main_amount
                )
                result["E_target_iiasa"] = target_from_absolute_reduction(
                    bau["bau_iiasa"], main_amount
                )

            result["target_status"] = (
                "NOT_YET_OBSERVABLE"
                if usability == "NOT_YET_OBSERVABLE"
                else "OK"
            )
            return result

        # Standard percentage BAU target
        if metric == "BAU_PER_CAPITA_TARGET":
            result["target_method"] = (
                "BAU_PER_CAPITA_NDC_OFFICIAL"
                if bau["bau_source"] == "NDC_OFFICIAL"
                else "BAU_PER_CAPITA_PMR_RCP6SSP2_MEDIAN"
            )
        else:
            result["target_method"] = (
                "BAU_NDC_OFFICIAL"
                if bau["bau_source"] == "NDC_OFFICIAL"
                else "BAU_PMR_RCP6SSP2_MEDIAN"
            )

        result["E_target_main"] = target_from_rate(
            bau["bau_main"], main_rate
        )
        result["E_target_U"] = target_from_rate(
            bau["bau_main"], u_rate
        )
        result["E_target_C"] = target_from_rate(
            bau["bau_main"], c_rate
        )

        # Sensitivities only when BAU is reconstructed via PMR.
        if bau["bau_source"] == "PMRCPBIE":
            result["E_target_rcp45"] = target_from_rate(
                bau["bau_rcp45"], main_rate
            )
            result["E_target_rcp85"] = target_from_rate(
                bau["bau_rcp85"], main_rate
            )
            result["E_target_oecd"] = target_from_rate(
                bau["bau_oecd"], main_rate
            )
            result["E_target_pik"] = target_from_rate(
                bau["bau_pik"], main_rate
            )
            result["E_target_iiasa"] = target_from_rate(
                bau["bau_iiasa"], main_rate
            )

        # Optional verified intermediate BAU target (piecewise path).
        interim_year = (
            year_num(row.get(INTERIM_TARGET_YEAR_COL))
            if INTERIM_TARGET_YEAR_COL is not None else np.nan
        )
        if pd.notna(interim_year):
            i_main = (
                pct_to_rate(row.get(INTERIM_MAIN_PCT_COL))
                if INTERIM_MAIN_PCT_COL is not None else np.nan
            )
            i_u = (
                pct_to_rate(row.get(INTERIM_UNCOND_PCT_COL))
                if INTERIM_UNCOND_PCT_COL is not None else np.nan
            )
            i_c = (
                pct_to_rate(row.get(INTERIM_COND_PCT_COL))
                if INTERIM_COND_PCT_COL is not None else np.nan
            )

            bau_i = bau_reference(row, iso3, interim_year)
            result["interim_target_year"] = interim_year
            result["E_interim_main"] = target_from_rate(
                bau_i["bau_main"], i_main
            )
            result["E_interim_U"] = target_from_rate(
                bau_i["bau_main"], i_u
            )
            result["E_interim_C"] = target_from_rate(
                bau_i["bau_main"], i_c
            )

            if bau_i["bau_source"] == "PMRCPBIE":
                result["E_interim_rcp45"] = target_from_rate(
                    bau_i["bau_rcp45"], i_main
                )
                result["E_interim_rcp85"] = target_from_rate(
                    bau_i["bau_rcp85"], i_main
                )
                result["E_interim_oecd"] = target_from_rate(
                    bau_i["bau_oecd"], i_main
                )
                result["E_interim_pik"] = target_from_rate(
                    bau_i["bau_pik"], i_main
                )
                result["E_interim_iiasa"] = target_from_rate(
                    bau_i["bau_iiasa"], i_main
                )

        built_any = any(pd.notna(result[k]) for k in [
            "E_target_main", "E_target_U", "E_target_C"
        ])
        result["target_status"] = (
            "NOT_YET_OBSERVABLE"
            if usability == "NOT_YET_OBSERVABLE" and built_any
            else ("OK" if built_any else "NEEDS_REVIEW")
        )
        if not built_any:
            result["target_note"] = (
                "BAU architecture identified but no usable percentage "
                "or absolute reduction amount was available"
            )
        return result

    # D. Per-capita target
    if metric == "PER_CAPITA_TARGET" or (
        PER_CAPITA_FLAG_COL is not None
        and int(to_num(row.get(PER_CAPITA_FLAG_COL)) or 0) == 1
    ):
        # Case 1: an explicit target level in tCO2e/person is reported.
        target_pc = (
            to_num(row.get(PER_CAPITA_TARGET_VALUE_COL))
            if PER_CAPITA_TARGET_VALUE_COL is not None
            else np.nan
        )

        if pd.notna(target_pc):
            pop_T_vals = socio_values(pop_lookup_full, iso3, T)
            pop_T = pop_T_vals.get("main_median", np.nan)

            if pd.notna(pop_T):
                result["E_target_main"] = target_pc * pop_T / 1000.0
                # A target expressed directly in tCO2e/person still depends
                # on the population projection used to convert it into an
                # absolute emissions level. Preserve that uncertainty in the
                # same RCP/model variants as for relative per-capita targets.
                result["E_target_rcp45"] = (
                    target_pc * pop_T_vals.get("rcp45_median", np.nan) / 1000.0
                )
                result["E_target_rcp85"] = (
                    target_pc * pop_T_vals.get("rcp85_median", np.nan) / 1000.0
                )
                result["E_target_oecd"] = (
                    target_pc * pop_T_vals.get("main_oecd", np.nan) / 1000.0
                )
                result["E_target_pik"] = (
                    target_pc * pop_T_vals.get("main_pik", np.nan) / 1000.0
                )
                result["E_target_iiasa"] = (
                    target_pc * pop_T_vals.get("main_iiasa", np.nan) / 1000.0
                )
                result["bau_source"] = "PMRCPBIE_POP_LEVEL_SSP2"
                result["bau_n_models"] = pop_T_vals.get("n_models_main", np.nan)
                result["target_method"] = "PER_CAPITA_LEVEL_PMR_SSP2"
                result["target_status"] = (
                    "NOT_YET_OBSERVABLE"
                    if usability == "NOT_YET_OBSERVABLE"
                    else "OK"
                )
                return result

        # Case 2: percentage reduction in emissions per capita relative
        # to a declared base year.
        if pd.isna(b):
            result["target_status"] = "NEEDS_REVIEW"
            result["target_method"] = "PER_CAPITA_MISSING_BASE_YEAR"
            result["target_note"] = (
                "Per-capita target identified but no usable base year exists"
            )
            return result

        E_base = observed_emissions_for_cycle(iso3, row.get(NDC_CYCLE_COL), b)
        scales = socio_scale_bundle(
            pop_lookup_full, iso3, b, T
        )

        result["target_method"] = "PER_CAPITA_RELATIVE_PMR_SSP2"

        result["E_target_main"] = target_from_scaled_relative(
            E_base, main_rate, scales["main"]
        )
        result["E_target_U"] = target_from_scaled_relative(
            E_base, u_rate, scales["main"]
        )
        result["E_target_C"] = target_from_scaled_relative(
            E_base, c_rate, scales["main"]
        )

        # Main-target uncertainty/sensitivity variants.
        result["E_target_rcp45"] = target_from_scaled_relative(
            E_base, main_rate, scales["rcp45"]
        )
        result["E_target_rcp85"] = target_from_scaled_relative(
            E_base, main_rate, scales["rcp85"]
        )
        result["E_target_oecd"] = target_from_scaled_relative(
            E_base, main_rate, scales["oecd"]
        )
        result["E_target_pik"] = target_from_scaled_relative(
            E_base, main_rate, scales["pik"]
        )
        result["E_target_iiasa"] = target_from_scaled_relative(
            E_base, main_rate, scales["iiasa"]
        )
        result["bau_n_models"] = scales["n_models"]
        result["bau_source"] = "PMRCPBIE_POP_SSP2"

        built_any = any(pd.notna(result[k]) for k in [
            "E_target_main", "E_target_U", "E_target_C"
        ])

        result["target_status"] = (
            "NOT_YET_OBSERVABLE"
            if usability == "NOT_YET_OBSERVABLE" and built_any
            else ("OK" if built_any else "NEEDS_REVIEW")
        )

        if not built_any:
            result["target_note"] = (
                "Per-capita target identified but EDGAR base emissions "
                "or PMRCPBIE population scale is missing"
            )
        return result

    # E. Intensity target
    if metric == "INTENSITY_TARGET":
        if pd.isna(b):
            result["target_status"] = "NEEDS_REVIEW"
            result["target_method"] = "INTENSITY_MISSING_BASE_YEAR"
            result["target_note"] = (
                "Intensity target identified but no usable base year exists"
            )
            return result

        # The percentage may be stored in the standard headline field or
        # in an explicit intensity field. Standard NDC coding has priority.
        intensity_rate = main_rate
        if pd.isna(intensity_rate) and INTENSITY_REDUCTION_COL is not None:
            intensity_rate = pct_to_rate(
                row.get(INTENSITY_REDUCTION_COL)
            )

        E_base = observed_emissions_for_cycle(iso3, row.get(NDC_CYCLE_COL), b)
        scales = socio_scale_bundle(
            gdp_lookup_full, iso3, b, T
        )

        result["target_method"] = "INTENSITY_RELATIVE_PMR_GDPPPP_SSP2"

        result["E_target_main"] = target_from_scaled_relative(
            E_base, intensity_rate, scales["main"]
        )
        result["E_target_U"] = target_from_scaled_relative(
            E_base, u_rate, scales["main"]
        )
        result["E_target_C"] = target_from_scaled_relative(
            E_base, c_rate, scales["main"]
        )

        result["E_target_rcp45"] = target_from_scaled_relative(
            E_base, intensity_rate, scales["rcp45"]
        )
        result["E_target_rcp85"] = target_from_scaled_relative(
            E_base, intensity_rate, scales["rcp85"]
        )
        result["E_target_oecd"] = target_from_scaled_relative(
            E_base, intensity_rate, scales["oecd"]
        )
        result["E_target_pik"] = target_from_scaled_relative(
            E_base, intensity_rate, scales["pik"]
        )
        result["E_target_iiasa"] = target_from_scaled_relative(
            E_base, intensity_rate, scales["iiasa"]
        )
        result["bau_n_models"] = scales["n_models"]
        result["bau_source"] = "PMRCPBIE_GDPPPP_SSP2"

        built_any = any(pd.notna(result[k]) for k in [
            "E_target_main", "E_target_U", "E_target_C"
        ])

        result["target_status"] = (
            "NOT_YET_OBSERVABLE"
            if usability == "NOT_YET_OBSERVABLE" and built_any
            else ("OK" if built_any else "NEEDS_REVIEW")
        )

        if not built_any:
            result["target_note"] = (
                "Intensity target identified but EDGAR base emissions, "
                "target rate, or PMRCPBIE GDP scale is missing"
            )
        return result

    # F. Autres architectures
    result["target_status"] = "NEEDS_REVIEW"
    result["target_note"] = f"Unhandled target metric: {metric}"
    return result

target_results = ndc.apply(
    lambda r: pd.Series(build_targets(r)),
    axis=1
)

# 6.1 COMMON SUPPORT FOR UNCERTAINTY SCENARIOS
# The alternative Gap variables must cover not only scenario-dependent BAU,
# intensity and per-capita targets, but also countries whose target conversion
# is scenario-invariant (base-year reductions, fixed/absolute levels, audited
# official BAU values, and analogous sector-matched targets).
#
# For a scenario-invariant architecture, the correct counterfactual under each
# alternative scenario is exactly the central target. We therefore replicate
# E_target_main (and any main interim target) across the five alternatives.
# For PMRCPBIE-dependent architectures, genuine alternative values are kept;
# a missing model projection remains missing and is never replaced by main.
_scenario_target_cols = [
    "E_target_rcp45", "E_target_rcp85", "E_target_oecd",
    "E_target_pik", "E_target_iiasa",
]
_scenario_interim_cols = [
    "E_interim_rcp45", "E_interim_rcp85", "E_interim_oecd",
    "E_interim_pik", "E_interim_iiasa",
]

_source = target_results["bau_source"].astype("string").fillna("")
_scenario_sensitive = _source.str.startswith("PMRCPBIE")
_scenario_invariant = (
    target_results["E_target_main"].notna() & ~_scenario_sensitive
)

for _col in _scenario_target_cols:
    target_results.loc[_scenario_invariant, _col] = (
        target_results.loc[_scenario_invariant, "E_target_main"]
    )

_invariant_with_interim = (
    _scenario_invariant & target_results["E_interim_main"].notna()
)
for _col in _scenario_interim_cols:
    target_results.loc[_invariant_with_interim, _col] = (
        target_results.loc[_invariant_with_interim, "E_interim_main"]
    )

target_results["uncertainty_scenario_status"] = np.select(
    [
        target_results["E_target_main"].isna(),
        _scenario_sensitive,
        _scenario_invariant,
    ],
    [
        "NO_MAIN_TARGET",
        "SCENARIO_SENSITIVE",
        "SCENARIO_INVARIANT_REPLICATED",
    ],
    default="UNCLASSIFIED",
)

# Prevent duplicate derived columns. Some final-audit fields (for example
# interim_target_year) already exist in the harmonized NDC input and are also
# returned by build_targets(). Duplicate names make row.get() return a Series.
_overlap_target_cols = [
    c for c in target_results.columns
    if c in ndc.columns
]
if _overlap_target_cols:
    ndc = ndc.drop(columns=_overlap_target_cols)

ndc_targets = pd.concat(
    [ndc.reset_index(drop=True), target_results.reset_index(drop=True)],
    axis=1
)

# Final safeguard: all downstream trajectory code assumes unique column names.
if ndc_targets.columns.duplicated().any():
    _dups = ndc_targets.columns[ndc_targets.columns.duplicated()].tolist()
    warnings.warn(
        "Duplicate columns detected after target construction; "
        f"keeping first occurrence only: {_dups}"
    )
    ndc_targets = ndc_targets.loc[:, ~ndc_targets.columns.duplicated()].copy()

# 7. CONSTRUCTION DES TRAJECTOIRES ANNUELLES

def build_path_rows(row):
    """
    Construit un calendrier OBSERVE commun 2015-2024 pour chaque pays/cycle.

    Convention méthodologique :
    - s = première année de mise en oeuvre de la NDC;
    - b = s - 1 = année d'ancrage de la trajectoire;
    - la trajectoire relie E_b à E_target(T);
    - le Gap est calculé uniquement pour s <= t <= T;
    - les lignes 2015-2024 existent même avant s, mais les Gap y restent
      manquants;
    - aucune ligne postérieure à 2024 n'est créée dans le panel observé.
    """
    iso3 = clean_text(row["iso3"]).upper()
    s = year_num(row["start_year"])
    b = year_num(row["trajectory_baseline_year"])
    T = year_num(row["target_year"])

    if not iso3:
        return []

    # E_b : émission observée l'année précédant la mise en oeuvre.
    E_b = (
        observed_emissions_for_cycle(iso3, row.get(NDC_CYCLE_COL), b)
        if pd.notna(b)
        else np.nan
    )

    target_variants = {
        "main": row.get("E_target_main", np.nan),
        "unconditional": row.get("E_target_U", np.nan),
        "conditional": row.get("E_target_C", np.nan),
        "rcp45": row.get("E_target_rcp45", np.nan),
        "rcp85": row.get("E_target_rcp85", np.nan),
        "oecd": row.get("E_target_oecd", np.nan),
        "pik": row.get("E_target_pik", np.nan),
        "iiasa": row.get("E_target_iiasa", np.nan),
    }

    interim_year = year_num(row.get("interim_target_year"))
    interim_variants = {
        "main": row.get("E_interim_main", np.nan),
        "unconditional": row.get("E_interim_U", np.nan),
        "conditional": row.get("E_interim_C", np.nan),
        "rcp45": row.get("E_interim_rcp45", np.nan),
        "rcp85": row.get("E_interim_rcp85", np.nan),
        "oecd": row.get("E_interim_oecd", np.nan),
        "pik": row.get("E_interim_pik", np.nan),
        "iiasa": row.get("E_interim_iiasa", np.nan),
    }

    harmonization_method = clean_text(
        row.get("harmonization_method", "")
    ).upper()

    # Ambition (%): required reduction from the trajectory baseline
    # emission E_b to the final NDC target.
    ambition_variants = {
        label: (
            100.0 * (E_b - E_target) / E_b
            if pd.notna(E_b) and E_b != 0 and pd.notna(E_target)
            else np.nan
        )
        for label, E_target in target_variants.items()
    }

    out = []

    for t in range(ANALYSIS_START, ANALYSIS_END + 1):
        E_t = observed_emissions_for_cycle(iso3, row.get(NDC_CYCLE_COL), t)

        base = {
            "iso3": iso3,
            "Country": row.get(COUNTRY_COL),
            "NDC_cycle": row.get(NDC_CYCLE_COL),
            "year": t,
            "start_year": int(s) if pd.notna(s) else np.nan,
            "trajectory_baseline_year": int(b) if pd.notna(b) else np.nan,
            "target_year": int(T) if pd.notna(T) else np.nan,
            "E_baseline": E_b,
            "E_obs": E_t,
            "target_metric_std": row.get(TARGET_METRIC_COL),
            "target_scope_std": (
                row.get(TARGET_SCOPE_COL)
                if TARGET_SCOPE_COL is not None else np.nan
            ),
            "quantified_target_coverage_std": (
                row.get(TARGET_COVERAGE_COL)
                if TARGET_COVERAGE_COL is not None else np.nan
            ),
            "target_usability_status": (
                row.get(TARGET_USABILITY_COL)
                if TARGET_USABILITY_COL is not None else np.nan
            ),
            "target_method": row.get("target_method"),
            "target_status": row.get("target_status"),
            "audit_decision": row.get("audit_decision"),
            "audit_reason": row.get("audit_reason"),
            "harmonization_method": row.get("harmonization_method"),
            "interim_target_year": interim_year,
            "bau_source": row.get("bau_source"),
            "bau_n_models": row.get("bau_n_models"),
        }

        # La trajectoire est définie depuis b=s-1, mais la performance NDC
        # n'est évaluée qu'à partir de s.
        path_active = (
            pd.notna(s)
            and pd.notna(b)
            and pd.notna(T)
            and pd.notna(E_b)
            and (b < T)
            and (t >= s)
            and (t <= T)
        )

        for label, E_target in target_variants.items():
            base[f"E_target_{label}"] = E_target
            base[f"E_interim_{label}"] = interim_variants.get(label, np.nan)
            base[f"Ambition_{label}"] = ambition_variants[label]

            if (not path_active) or pd.isna(E_target):
                base[f"E_path_{label}"] = np.nan
                base[f"Gap_abs_{label}"] = np.nan
                base[f"Gap_{label}"] = np.nan
                base[f"Gap_asinh_{label}"] = np.nan
                continue

            # Special cumulative-reduction trajectory:
            # use the annual BAU and the validated linear ramp in abatement.
            if harmonization_method == "CUMULATIVE_REDUCTION_LINEAR_RAMP":
                total_amount = (
                    to_num(row.get(SPECIAL_TARGET_AMOUNT_COL))
                    if SPECIAL_TARGET_AMOUNT_COL is not None else np.nan
                )
                n_years = int(T - s + 1)
                denom = n_years * (n_years + 1) / 2.0
                step = int(t - s + 1)
                annual_reduction = (
                    total_amount * step / denom
                    if pd.notna(total_amount) and step >= 1 else np.nan
                )

                if label == "main":
                    bau_t = bau_reference(row, iso3, t)["bau_main"]
                elif label == "rcp45":
                    bau_t = bau_reference(row, iso3, t)["bau_rcp45"]
                elif label == "rcp85":
                    bau_t = bau_reference(row, iso3, t)["bau_rcp85"]
                elif label == "oecd":
                    bau_t = bau_reference(row, iso3, t)["bau_oecd"]
                elif label == "pik":
                    bau_t = bau_reference(row, iso3, t)["bau_pik"]
                elif label == "iiasa":
                    bau_t = bau_reference(row, iso3, t)["bau_iiasa"]
                else:
                    bau_t = np.nan

                E_path = (
                    bau_t - annual_reduction
                    if pd.notna(bau_t) and pd.notna(annual_reduction)
                    else np.nan
                )

            # Piecewise path when an audited intermediate target exists:
            # E_b -> E_interim -> E_target.
            elif (
                pd.notna(interim_year)
                and b < interim_year < T
                and pd.notna(interim_variants.get(label, np.nan))
            ):
                E_interim = interim_variants[label]
                if t <= interim_year:
                    E_path = (
                        E_b
                        + ((t - b) / (interim_year - b))
                        * (E_interim - E_b)
                    )
                else:
                    E_path = (
                        E_interim
                        + ((t - interim_year) / (T - interim_year))
                        * (E_target - E_interim)
                    )

            else:
                E_path = (
                    E_b
                    + ((t - b) / (T - b))
                    * (E_target - E_b)
                )

            base[f"E_path_{label}"] = E_path

            if pd.notna(E_t):
                gap_abs = E_t - E_path
                base[f"Gap_abs_{label}"] = gap_abs
                base[f"Gap_{label}"] = (
                    100.0 * gap_abs / E_path if E_path != 0 else np.nan
                )
                base[f"Gap_asinh_{label}"] = np.arcsinh(gap_abs)

                # IMPLEMENTATION PROGRESS SCORE — DISABLED
                # The following score was considered in an earlier version:
                #
                # required_reduction = E_b - E_path
                # actual_reduction = E_b - E_t
                # Implementation = (
                #     100.0 * actual_reduction / required_reduction
                #     if required_reduction > 0 else np.nan
                # )
                #
                # It is not constructed or exported in the current analysis.
                # When the required reduction is close to zero, the ratio is
                # mechanically unstable and can take extreme values that are
                # difficult to interpret. The Gap-to-Target remains the main
                # annual alignment indicator. This commented formula is kept
                # only to document the earlier specification.
            else:
                base[f"Gap_abs_{label}"] = np.nan
                base[f"Gap_{label}"] = np.nan
                base[f"Gap_asinh_{label}"] = np.nan

        out.append(base)

    return out

panel_rows = []

if ndc_targets.columns.duplicated().any():
    _dups = ndc_targets.columns[ndc_targets.columns.duplicated()].tolist()
    warnings.warn(
        "Duplicate columns detected before annual path construction; "
        f"keeping first occurrence only: {_dups}"
    )
    ndc_targets = ndc_targets.loc[:, ~ndc_targets.columns.duplicated()].copy()

for _, row in ndc_targets.iterrows():
    panel_rows.extend(build_path_rows(row))

gap_panel = pd.DataFrame(panel_rows)

# Strictly historical observed panel: no post-2024 rows and no rows without
# observed EDGAR emissions.
if not gap_panel.empty:
    gap_panel = gap_panel[
        gap_panel["year"].between(ANALYSIS_START, ANALYSIS_END)
        & gap_panel["E_obs"].notna()
    ].copy()

# 7.1 PANEL ANNUEL AVEC COLONNES SPECIFIQUES PAR CYCLE
#
# Gap_Panel reste le format long détaillé.
# cycle_panel ajoute une colonne spécifique pour NDC1 et NDC2 afin de
# faciliter Stata, les tableaux et les graphiques.
#
# Exemple :
#   Gap_main_n1, Gap_unconditional_n1, Gap_conditional_n1
#   Gap_main_n2, Gap_unconditional_n2, Gap_conditional_n2

CYCLE_SUFFIX = {
    "First_NDC": "n1",
    "Second_NDC": "n2",
}

cycle_value_cols = [
    "E_obs", "start_year", "trajectory_baseline_year", "target_year", "E_baseline",
    "target_metric_std", "target_scope_std",
    "quantified_target_coverage_std", "target_usability_status",
    "target_method", "target_status",
    "audit_decision", "audit_reason", "harmonization_method",
    "interim_target_year",
    "bau_source", "bau_n_models",
]

for variant in [
    "main", "unconditional", "conditional",
    "rcp45", "rcp85", "oecd", "pik", "iiasa"
]:
    cycle_value_cols.extend([
        f"E_target_{variant}",
        f"E_interim_{variant}",
        f"E_path_{variant}",
        f"Gap_abs_{variant}",
        f"Gap_{variant}",
        f"Gap_asinh_{variant}",
        f"Ambition_{variant}",
        # f"Implementation_{variant}",  # disabled in the current analysis
    ])

cycle_value_cols = [
    c for c in cycle_value_cols if c in gap_panel.columns
]

# Universe analytique : tous les pays First/Second NDC avec ISO3 résolu.
country_universe = (
    ndc_targets[["iso3", COUNTRY_COL]]
    .dropna(subset=["iso3"])
    .drop_duplicates("iso3")
    .rename(columns={COUNTRY_COL: "Country"})
    .copy()
)

calendar = pd.DataFrame({
    "year": np.arange(ANALYSIS_START, ANALYSIS_END + 1, dtype=int)
})

country_universe["_key"] = 1
calendar["_key"] = 1

cycle_base = (
    country_universe
    .merge(calendar, on="_key", how="inner")
    .drop(columns="_key")
)

cycle_base["E_obs"] = cycle_base.apply(
    lambda r: edgar_lookup.get(
        (clean_text(r["iso3"]).upper(), int(r["year"])),
        np.nan
    ),
    axis=1
)

# Common national EDGAR series retained only as a diagnostic/backward-compatible column.
# Gap calculations use E_obs_n1 / E_obs_n2 from the cycle-specific harmonised scope.
cycle_base = cycle_base[
    cycle_base["E_obs"].notna()
].copy()

cycle_panel = cycle_base.copy()

for cycle_name, suffix in CYCLE_SUFFIX.items():
    tmp = gap_panel[
        gap_panel["NDC_cycle"].astype(str).eq(cycle_name)
    ].copy()

    if tmp.empty:
        continue

    tmp = tmp[
        ["iso3", "year"] + cycle_value_cols
    ].drop_duplicates(["iso3", "year"], keep="first")

    rename_cycle = {
        c: f"{c}_{suffix}" for c in cycle_value_cols
    }
    tmp = tmp.rename(columns=rename_cycle)

    cycle_panel = cycle_panel.merge(
        tmp,
        on=["iso3", "year"],
        how="left"
    )

cycle_panel = cycle_panel.sort_values(
    ["iso3", "year"]
).reset_index(drop=True)

# 7A. CLASSIFICATION DE LA COUVERTURE SECTORIELLE DES CIBLES
#
# Classification validée pour les analyses d'hétérogénéité :
#
#   1 = Economy-wide
#       Cible explicitement nationale / economy-wide.
#
#   2 = Broad multi-sector
#       Cible couvrant un ensemble large de secteurs d'émissions
#       sans être explicitement economy-wide.
#
#   3 = Limited sectoral
#       Cible couvrant plusieurs secteurs, mais avec une portée
#       sectorielle partielle / limitée.
#
#   4 = Single-sector
#       Cible portant sur un seul secteur ou sous-secteur.
#
# IMPORTANT :
# La classification décrit la couverture de la CIBLE QUANTIFIEE,
# pas simplement la liste des secteurs dans lesquels le pays annonce
# des politiques ou mesures. Elle est construite séparément pour NDC1
# et NDC2 afin de permettre à la couverture d'évoluer entre les cycles.

COVERAGE_CODE_LABELS = {
    1: "Economy-wide",
    2: "Broad multi-sector",
    3: "Limited sectoral",
    4: "Single-sector",
}

def classify_coverage_cluster(coverage, scope):
    """
    Map the harmonised quantified-target coverage to the four validated
    empirical clusters.

    Priority is given to quantified_target_coverage_std because it describes
    the coverage of the quantified GHG target itself. target_scope_std is used
    only as a conservative fallback when the quantified-coverage field is
    missing or unclear.
    """
    coverage = clean_text(coverage).upper()
    scope = clean_text(scope).upper()
    if scope in {"NO_QUANTIFIED_GHG_TARGET", "NO_SUBMISSION", "OUTSIDE_ANALYSIS_WINDOW"}:
        return np.nan

    # Explicit economy-wide / national GHG target.
    if coverage == "NATIONAL_GHG":
        return 1

    # Broad national coverage without explicit economy-wide wording.
    if coverage == "BROAD_NATIONAL_GHG":
        return 2

    # Audited multi-sector / partial-sector GHG target.
    if coverage == "PARTIAL_SECTORAL_GHG":
        return 3

    # One quantified sector / sub-sector only.
    if coverage == "SINGLE_SECTOR_GHG":
        return 4

    # Conservative fallbacks from the descriptive scope variable.
    if scope == "ECONOMY_WIDE":
        return 1
    if scope == "ECONOMY_WIDE_OR_BROAD":
        return 2
    if scope == "MULTISECTORAL":
        return 3
    if scope == "SECTORAL":
        return 4

    return np.nan

def coverage_code_to_text(code):
    if pd.isna(code):
        return np.nan
    return COVERAGE_CODE_LABELS.get(int(code), np.nan)

for suffix in ["n1", "n2"]:
    coverage_col = f"quantified_target_coverage_std_{suffix}"
    scope_col = f"target_scope_std_{suffix}"

    if coverage_col in cycle_panel.columns or scope_col in cycle_panel.columns:
        coverage_series = (
            cycle_panel[coverage_col]
            if coverage_col in cycle_panel.columns
            else pd.Series("", index=cycle_panel.index)
        )
        scope_series = (
            cycle_panel[scope_col]
            if scope_col in cycle_panel.columns
            else pd.Series("", index=cycle_panel.index)
        )

        cycle_panel[f"coverage_cluster_{suffix}"] = [
            classify_coverage_cluster(cov, scp)
            for cov, scp in zip(coverage_series, scope_series)
        ]

        cycle_panel[f"coverage_class_{suffix}"] = (
            cycle_panel[f"coverage_cluster_{suffix}"]
            .map(coverage_code_to_text)
        )

        # Store the numeric cluster as nullable integer where possible.
        cycle_panel[f"coverage_cluster_{suffix}"] = (
            pd.to_numeric(
                cycle_panel[f"coverage_cluster_{suffix}"],
                errors="coerce"
            )
            .astype("Int64")
        )

# Convenience aliases with the older n1/n2 naming logic.
# These aliases do NOT change the indicator: they simply expose the
# cycle-specific variables directly.
alias_map = {
    "Gap_main_n1": "gap_n1",
    "Gap_unconditional_n1": "gap_n1_uncond",
    "Gap_conditional_n1": "gap_n1_cond",
    "Gap_main_n2": "gap_n2",
    "Gap_unconditional_n2": "gap_n2_uncond",
    "Gap_conditional_n2": "gap_n2_cond",

    "Gap_abs_main_n1": "gap_abs_n1",
    "Gap_abs_unconditional_n1": "gap_abs_n1_uncond",
    "Gap_abs_conditional_n1": "gap_abs_n1_cond",
    "Gap_abs_main_n2": "gap_abs_n2",
    "Gap_abs_unconditional_n2": "gap_abs_n2_uncond",
    "Gap_abs_conditional_n2": "gap_abs_n2_cond",

    "Gap_asinh_main_n1": "gap_asinh_n1",
    "Gap_asinh_unconditional_n1": "gap_asinh_n1_uncond",
    "Gap_asinh_conditional_n1": "gap_asinh_n1_cond",
    "Gap_asinh_main_n2": "gap_asinh_n2",
    "Gap_asinh_unconditional_n2": "gap_asinh_n2_uncond",
    "Gap_asinh_conditional_n2": "gap_asinh_n2_cond",
}

for src_col, alias_col in alias_map.items():
    if src_col in cycle_panel.columns:
        cycle_panel[alias_col] = cycle_panel[src_col]

# Aliases Stata-friendly pour incertitude de scénario et sensibilité
# aux variantes OECD / PIK / IIASA.
for cyc in ["n1", "n2"]:
    for variant in ["rcp45", "rcp85", "oecd", "pik", "iiasa"]:
        mappings = {
            f"Gap_{variant}_{cyc}": f"gap_{cyc}_{variant}",
            f"Gap_abs_{variant}_{cyc}": f"gap_abs_{cyc}_{variant}",
            f"Gap_asinh_{variant}_{cyc}": f"gap_asinh_{cyc}_{variant}",
        }
        for src_col, alias_col in mappings.items():
            if src_col in cycle_panel.columns:
                cycle_panel[alias_col] = cycle_panel[src_col]

# Aliases compacts des indicateurs d'ambition (%).
for cyc in ["n1", "n2"]:
    for variant in [
        "main", "unconditional", "conditional",
        "rcp45", "rcp85", "oecd", "pik", "iiasa"
    ]:
        src_col = f"Ambition_{variant}_{cyc}"
        suffix = {
            "main": "",
            "unconditional": "_uncond",
            "conditional": "_cond",
            "rcp45": "_rcp45",
            "rcp85": "_rcp85",
            "oecd": "_oecd",
            "pik": "_pik",
            "iiasa": "_iiasa",
        }[variant]
        alias_col = f"ambition_{cyc}{suffix}"
        if src_col in cycle_panel.columns:
            cycle_panel[alias_col] = cycle_panel[src_col]

# Implementation Progress Score aliases are intentionally not created.
# The score is excluded from the current empirical and graphical analysis.

# Variables compactes utilisées dans les estimations principales.
analysis_keep = [
    "iso3", "Country", "year", "E_obs",

    # ---------------- First NDC ----------------
    "start_year_n1", "trajectory_baseline_year_n1", "target_year_n1",
    "E_baseline_n1",

    # Main
    "E_target_main_n1", "E_path_main_n1",
    "gap_abs_n1", "gap_n1", "gap_asinh_n1",

    # Conditionality
    "E_target_unconditional_n1", "E_path_unconditional_n1",
    "gap_abs_n1_uncond", "gap_n1_uncond", "gap_asinh_n1_uncond",
    "E_target_conditional_n1", "E_path_conditional_n1",
    "gap_abs_n1_cond", "gap_n1_cond", "gap_asinh_n1_cond",

    # Scenario uncertainty: RCP4.5 / RCP8.5, SSP2 held constant
    "E_target_rcp45_n1", "E_path_rcp45_n1",
    "gap_abs_n1_rcp45", "gap_n1_rcp45", "gap_asinh_n1_rcp45",
    "E_target_rcp85_n1", "E_path_rcp85_n1",
    "gap_abs_n1_rcp85", "gap_n1_rcp85", "gap_asinh_n1_rcp85",

    # Downscaling/GDP-source sensitivity under RCP6-SSP2
    "E_target_oecd_n1", "E_path_oecd_n1",
    "gap_abs_n1_oecd", "gap_n1_oecd", "gap_asinh_n1_oecd",
    "E_target_pik_n1", "E_path_pik_n1",
    "gap_abs_n1_pik", "gap_n1_pik", "gap_asinh_n1_pik",
    "E_target_iiasa_n1", "E_path_iiasa_n1",
    "gap_abs_n1_iiasa", "gap_n1_iiasa", "gap_asinh_n1_iiasa",

    "target_metric_std_n1", "target_scope_std_n1",
    "target_method_n1", "target_status_n1",
    "bau_source_n1", "bau_n_models_n1",

    # ---------------- Second NDC ----------------
    "start_year_n2", "trajectory_baseline_year_n2", "target_year_n2",
    "E_baseline_n2",

    # Main
    "E_target_main_n2", "E_path_main_n2",
    "gap_abs_n2", "gap_n2", "gap_asinh_n2",

    # Conditionality
    "E_target_unconditional_n2", "E_path_unconditional_n2",
    "gap_abs_n2_uncond", "gap_n2_uncond", "gap_asinh_n2_uncond",
    "E_target_conditional_n2", "E_path_conditional_n2",
    "gap_abs_n2_cond", "gap_n2_cond", "gap_asinh_n2_cond",

    # Scenario uncertainty
    "E_target_rcp45_n2", "E_path_rcp45_n2",
    "gap_abs_n2_rcp45", "gap_n2_rcp45", "gap_asinh_n2_rcp45",
    "E_target_rcp85_n2", "E_path_rcp85_n2",
    "gap_abs_n2_rcp85", "gap_n2_rcp85", "gap_asinh_n2_rcp85",

    # Downscaling/GDP-source sensitivity
    "E_target_oecd_n2", "E_path_oecd_n2",
    "gap_abs_n2_oecd", "gap_n2_oecd", "gap_asinh_n2_oecd",
    "E_target_pik_n2", "E_path_pik_n2",
    "gap_abs_n2_pik", "gap_n2_pik", "gap_asinh_n2_pik",
    "E_target_iiasa_n2", "E_path_iiasa_n2",
    "gap_abs_n2_iiasa", "gap_n2_iiasa", "gap_asinh_n2_iiasa",

    "target_metric_std_n2", "target_scope_std_n2",
    "target_method_n2", "target_status_n2",
    "bau_source_n2", "bau_n_models_n2",

    # Quantified-target sector coverage clusters
    "coverage_cluster_n1", "coverage_class_n1",
    "coverage_cluster_n2", "coverage_class_n2",
]

analysis_keep = [
    c for c in analysis_keep if c in cycle_panel.columns
]

analysis_panel = cycle_panel[analysis_keep].copy()

# 7.2 FINAL PANEL — ISO3 + YEAR + GAP + AMBITION
# Keep every Gap-family variable (relative, absolute, asinh,
# conditionality, BAU scenario uncertainty and model sensitivity).
# No emissions, target, path, country-name or technical variables are
# retained in the final econometric dataset.

final_gap_cols = [
    c for c in cycle_panel.columns
    if c.startswith("gap_") and c == c.lower()
]
final_ambition_cols = [
    c for c in cycle_panel.columns
    if c.startswith("ambition_") and c == c.lower()
]
# Implementation scores are deliberately excluded from the final panel.
final_impl_cols = []

# Stable ordering: main relative gaps first, then the remaining Gap family.
preferred_gap_order = [
    "gap_n1", "gap_n2",
    "gap_n1_uncond", "gap_n1_cond",
    "gap_n2_uncond", "gap_n2_cond",
    "gap_n1_rcp45", "gap_n1_rcp85",
    "gap_n2_rcp45", "gap_n2_rcp85",
    "gap_n1_oecd", "gap_n1_pik", "gap_n1_iiasa",
    "gap_n2_oecd", "gap_n2_pik", "gap_n2_iiasa",
]
final_gap_cols = (
    [c for c in preferred_gap_order if c in final_gap_cols]
    + sorted(c for c in final_gap_cols if c not in preferred_gap_order)
)

preferred_ambition_order = [
    "ambition_n1", "ambition_n2",
    "ambition_n1_uncond", "ambition_n1_cond",
    "ambition_n2_uncond", "ambition_n2_cond",
    "ambition_n1_rcp45", "ambition_n1_rcp85",
    "ambition_n2_rcp45", "ambition_n2_rcp85",
    "ambition_n1_oecd", "ambition_n1_pik", "ambition_n1_iiasa",
    "ambition_n2_oecd", "ambition_n2_pik", "ambition_n2_iiasa",
]
final_ambition_cols = (
    [c for c in preferred_ambition_order if c in final_ambition_cols]
    + sorted(c for c in final_ambition_cols if c not in preferred_ambition_order)
)

final_coverage_cols = [
    c for c in [
        "coverage_cluster_n1", "coverage_class_n1",
        "coverage_cluster_n2", "coverage_class_n2",
    ]
    if c in cycle_panel.columns
]

final_panel = cycle_panel[
    ["iso3", "year"]
    + final_gap_cols
    + final_ambition_cols
    + final_impl_cols
    + final_coverage_cols
].copy()
final_panel["audit_schema_version"] = AUDIT_SCHEMA_VERSION

# Country-level exclusion rule:
# A country is excluded only if it has NO usable main Gap observation
# for BOTH NDC1 and NDC2 over 2015-2024.
main_gap_cols = [
    c for c in ["gap_n1", "gap_n2"]
    if c in final_panel.columns
]

if not main_gap_cols:
    raise ValueError(
        "Neither gap_n1 nor gap_n2 exists; final dataset cannot be built."
    )

country_has_main_gap = (
    final_panel
    .groupby("iso3")[main_gap_cols]
    .apply(lambda g: bool(g.notna().any().any()))
)

eligible_iso3 = country_has_main_gap[
    country_has_main_gap
].index.tolist()

# Final raw-NDC audit safeguard embedded directly in the Builder. No external
# audit dataset is needed. These countries are removed even if an upstream
# generic rule accidentally produces a numerical target.
FINAL_MANUAL_EXCLUSIONS = {
    "ATG", "BHR", "BOL", "CUB", "TLS", "ECU", "GUY", "IRN", "IRQ", "LBY",
    "NRU", "NPL", "PNG", "SAU", "SSD", "SUR", "SYR", "URY", "VUT", "YEM",
}

if len(FINAL_MANUAL_EXCLUSIONS) != 20:
    raise AssertionError("The final raw-NDC exclusion list must contain 20 countries.")
if "EGY" in FINAL_MANUAL_EXCLUSIONS:
    raise AssertionError("Egypt must be retained because its NDC2 Gap is constructible.")

excluded_no_gap_iso3 = country_has_main_gap[
    ~country_has_main_gap
].index.tolist()

eligible_iso3 = [
    iso for iso in eligible_iso3
    if iso not in FINAL_MANUAL_EXCLUSIONS
]

excluded_iso3 = sorted(
    set(excluded_no_gap_iso3) | FINAL_MANUAL_EXCLUSIONS
)

final_panel = (
    final_panel[
        final_panel["iso3"].isin(eligible_iso3)
        & final_panel["year"].between(ANALYSIS_START, ANALYSIS_END)
    ]
    .sort_values(["iso3", "year"])
    .reset_index(drop=True)
)

# Final country-audit invariants. Bosnia and Herzegovina remains eligible
# through its Second NDC. Its First NDC cannot be reconstructed because the
# 1990 FAOSTAT Forest Land base-year component is unavailable. FSM and Tuvalu
# are not required inclusions: their quantified sectoral targets cannot be
# matched to scope-compatible EDGAR energy components.
REQUIRED_AUDITED_INCLUSIONS = {"BIH"}
_retained_iso3 = set(final_panel["iso3"].dropna().astype(str).str.upper())
_missing_required = REQUIRED_AUDITED_INCLUSIONS - _retained_iso3
if _missing_required:
    raise AssertionError(
        "Audited countries unexpectedly absent from the final panel: "
        + ", ".join(sorted(_missing_required))
    )

_required_cycle_gaps = {
    "BIH": ("gap_n2",),
}
for _iso3, _gap_cols in _required_cycle_gaps.items():
    _country_rows = final_panel[final_panel["iso3"].eq(_iso3)]
    for _gap_col in _gap_cols:
        if _gap_col not in _country_rows.columns or not _country_rows[_gap_col].notna().any():
            raise AssertionError(
                f"Audited target was not constructed: {_iso3} {_gap_col}. "
                "Check the country-cycle scope bridge and base-year emissions."
            )

# Integrity checks for the final analytical sample. Missing Gap values before
# implementation are intentionally retained; only countries with both main
# Gap series missing over the whole 2015-2024 window are removed.
expected_years = ANALYSIS_END - ANALYSIS_START + 1
country_year_counts = final_panel.groupby("iso3")["year"].nunique()
if not country_year_counts.eq(expected_years).all():
    bad = country_year_counts[country_year_counts.ne(expected_years)]
    raise ValueError(
        "Unbalanced final panel: some eligible countries do not have "
        f"{expected_years} years.\n{bad.to_string()}"
    )

# Final analytical sample expected after applying the embedded country audit and
# the scope-compatible EDGAR/FAOSTAT availability rules.
EXPECTED_FINAL_COUNTRIES = 164
EXPECTED_FINAL_ROWS = EXPECTED_FINAL_COUNTRIES * expected_years
_actual_final_countries = final_panel["iso3"].nunique()
_actual_final_rows = len(final_panel)
if (
    _actual_final_countries != EXPECTED_FINAL_COUNTRIES
    or _actual_final_rows != EXPECTED_FINAL_ROWS
):
    raise AssertionError(
        "Unexpected final analytical sample after the Paris audit: "
        f"{_actual_final_countries} countries and {_actual_final_rows} rows; "
        f"expected {EXPECTED_FINAL_COUNTRIES} countries and "
        f"{EXPECTED_FINAL_ROWS} rows. Check the Harmonizer output and the "
        "country-level exclusion audit."
    )

print("\n" + "=" * 80)
print("FINAL ANALYTICAL SAMPLE")
print("=" * 80)
print(f"Countries retained: {final_panel['iso3'].nunique()}")
print(f"Years: {ANALYSIS_START}-{ANALYSIS_END} ({expected_years} years)")
print(f"Observations: {len(final_panel)}")

# Country-level distribution of coverage clusters.
# Because coverage is cycle-specific and constant within country/cycle, count
# each country once for NDC1 and once for NDC2 where the class is available.
for _suffix, _label in [("n1", "NDC1"), ("n2", "NDC2")]:
    _ccol = f"coverage_cluster_{_suffix}"
    _tcol = f"coverage_class_{_suffix}"
    if _ccol in final_panel.columns:
        _country_cov = (
            final_panel[["iso3", _ccol, _tcol]]
            .drop_duplicates("iso3")
            .dropna(subset=[_ccol])
        )
        print(f"\nCoverage classes — {_label}:")
        if _country_cov.empty:
            print("  No classified countries.")
        else:
            _counts = (
                _country_cov[_tcol]
                .value_counts(dropna=False)
                .reindex([
                    "Economy-wide",
                    "Broad multi-sector",
                    "Limited sectoral",
                    "Single-sector",
                ])
                .dropna()
            )
            for _class, _n in _counts.items():
                print(f"  {_class}: {int(_n)}")

print(f"Excluded because GapN1 and GapN2 are both missing throughout: {len(excluded_no_gap_iso3)}")
print(f"Manual audit exclusions: {', '.join(sorted(FINAL_MANUAL_EXCLUSIONS))}")
print("Retained ISO3:")
print(" ".join(sorted(final_panel["iso3"].dropna().unique())))

# IMPORTANT: do NOT drop pre-implementation rows merely because their
# Gap values are missing. This preserves the common 2015-2024 calendar
# for every eligible country.

# Dataset graphique : volontairement compact.
graph_keep = [
    "iso3", "Country", "year", "E_obs",
    "E_path_main_n1", "E_path_unconditional_n1",
    "E_path_conditional_n1",
    "E_path_main_n2", "E_path_unconditional_n2",
    "E_path_conditional_n2",
    "E_target_main_n1", "E_target_unconditional_n1",
    "E_target_conditional_n1",
    "E_target_main_n2", "E_target_unconditional_n2",
    "E_target_conditional_n2",
]
graph_keep = [
    c for c in graph_keep if c in cycle_panel.columns
]
country_graphs = cycle_panel[graph_keep].copy()

# 8. AUDIT QUALITE

audit_cols = [
    "iso3", NDC_CYCLE_COL,
    "iso3_match_source",
    COUNTRY_COL,
    NDC_CYCLE_COL,
    TARGET_METRIC_COL,
    TARGET_COVERAGE_COL,
    TARGET_USABILITY_COL,
    "target_year",
    "start_year",
    "target_method",
    "target_status",
    "target_note",
    "audit_decision",
    "audit_reason",
    "harmonization_method",
    "interim_target_year",
    "bau_source",
    "bau_main",
    "bau_n_models",
    "E_target_main",
    "E_target_U",
    "E_target_C",
]

# Keep only existing columns AND remove duplicates while preserving order.
# NDC_CYCLE_COL may otherwise appear twice (explicitly and through aliases),
# which makes row access return a pandas Series instead of a scalar.
audit_cols = list(dict.fromkeys(
    c for c in audit_cols if c in ndc_targets.columns
))

audit = ndc_targets[audit_cols].copy()

audit["has_scope_emissions_start"] = audit.apply(
    lambda r: int(
        pd.notna(
            observed_emissions_for_cycle(
                clean_text(r["iso3"]).upper(),
                r.get(NDC_CYCLE_COL),
                r["start_year"]
            )
        )
    )
    if pd.notna(r["start_year"]) else 0,
    axis=1
)

audit["target_built_flag"] = (
    audit["E_target_main"].notna()
    | audit["E_target_U"].notna()
    | audit["E_target_C"].notna()
).astype(int)

audit["three_bau_variants_flag"] = np.where(
    audit["bau_source"].eq("PMRCPBIE"),
    (audit["bau_n_models"] == 3).astype(int),
    np.nan
)

# 9. RESUME CONSOLE

print("\n" + "=" * 80)
print("PARIS GAP-TO-TARGET — RESUME")
print("=" * 80)

print(f"NDC analysées : {len(ndc_targets):,}")
print(
    "Cibles construites : "
    f"{audit['target_built_flag'].sum():,}"
)
print(
    "Cibles à revoir / non construites : "
    f"{(audit['target_built_flag'] == 0).sum():,}"
)

print("\nStatut de construction :")
print(
    ndc_targets["target_status"]
    .value_counts(dropna=False)
    .to_string()
)

print("\nMéthode de construction :")
print(
    ndc_targets["target_method"]
    .value_counts(dropna=False)
    .to_string()
)

if "bau_n_models" in ndc_targets.columns:
    print("\nNombre de variantes BAU disponibles (PMR, RCP6-SSP2) :")
    print(
        ndc_targets.loc[
            ndc_targets["bau_source"].eq("PMRCPBIE"),
            "bau_n_models"
        ]
        .value_counts(dropna=False)
        .sort_index()
        .to_string()
    )

print(f"\nFenêtre analytique observée : {ANALYSIS_START}-{ANALYSIS_END}")
print("Ancrage des trajectoires : année précédant le début de mise en oeuvre (b=s-1)")
print("Gap relatif et Ambition : exprimés en pourcentage (x100)")
print("Implementation Progress Score : désactivé et non exporté")
print("Intensity targets: E_base*(1-r)*(GDP_T/GDP_base), PMRCPBIE SSP2")
print("Per-capita targets: E_base*(1-r)*(POP_T/POP_base), PMRCPBIE SSP2")
print(f"Panel annuel long créé : {len(gap_panel):,} lignes")
print(f"Panel annuel NDC1/NDC2 en colonnes : {len(cycle_panel):,} lignes")
print(
    "Pays avec émissions EDGAR dans le panel : "
    f"{cycle_panel['iso3'].nunique():,}"
)
print(
    "Années présentes : "
    f"{int(cycle_panel['year'].min()) if len(cycle_panel) else 'NA'}-"
    f"{int(cycle_panel['year'].max()) if len(cycle_panel) else 'NA'}"
)

_counts_country = cycle_panel.groupby("iso3")["year"].nunique()
print(
    "Pays avec les 10 années 2015-2024 : "
    f"{(_counts_country == 10).sum():,} / {_counts_country.size:,}"
)
if (_counts_country != 10).any():
    print("\nPays avec moins de 10 années observées :")
    print(_counts_country[_counts_country != 10].sort_values().to_string())

if "gap_n1" in analysis_panel.columns:
    print(
        "Observations annuelles avec Gap NDC1 principal : "
        f"{analysis_panel['gap_n1'].notna().sum():,}"
    )
if "gap_n2" in analysis_panel.columns:
    print(
        "Observations annuelles avec Gap NDC2 principal : "
        f"{analysis_panel['gap_n2'].notna().sum():,}"
    )

# 10. EXPORT EXCEL

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl",
    mode="w"
) as writer:

    ndc_targets.to_excel(
        writer,
        sheet_name="NDC_Targets",
        index=False
    )

    gap_panel.to_excel(
        writer,
        sheet_name="Gap_Panel",
        index=False
    )

    cycle_panel.to_excel(
        writer,
        sheet_name="Cycle_Panel",
        index=False
    )

    audit.to_excel(
        writer,
        sheet_name="Audit",
        index=False
    )

    bau_extract = ghg_wide[
        [
            "country", "year",
            "main_oecd", "main_pik", "main_iiasa",
            "bau_rcp6_median", "bau_n_models_rcp6",
            "bau_rcp45_median", "bau_rcp85_median"
        ]
    ].copy()

    # BAU_Extract peut contenir des années futures car il sert uniquement
    # à reconstruire les niveaux cibles à T. Il ne constitue pas le panel
    # d'émissions observées.
    bau_extract.to_excel(
        writer,
        sheet_name="BAU_Extract",
        index=False
    )

    # V10: reconstruct the socioeconomic audit tables from the
    # scenario/model-resolved lookup dictionaries used for intensity
    # and per-capita target conversion.
    pop_country_year = (
        pd.DataFrame.from_dict(pop_lookup_full, orient="index")
        .reset_index()
        .rename(columns={"level_0": "country", "level_1": "year"})
    )
    pop_country_year["unit"] = "persons"

    gdp_country_year = (
        pd.DataFrame.from_dict(gdp_lookup_full, orient="index")
        .reset_index()
        .rename(columns={"level_0": "country", "level_1": "year"})
    )
    gdp_country_year["unit"] = "Million2011GKD"

    pop_country_year.to_excel(
        writer,
        sheet_name="Population_SSP2",
        index=False
    )

    gdp_country_year.to_excel(
        writer,
        sheet_name="GDP_SSP2",
        index=False
    )

    # Mise en forme simple
    wb = writer.book

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
            cell.alignment = cell.alignment.copy(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        for col_cells in ws.columns:
            letter = col_cells[0].column_letter
            header = clean_text(col_cells[0].value)
            ws.column_dimensions[letter].width = min(
                max(len(header) + 2, 12), 28
            )

# 10.1 LABELS DU DATASET FINAL

final_labels = {
    "iso3": "ISO3 country code",
    "year": "Year",
}

def make_final_gap_label(var):
    """Generate a concise, complete Stata label for every final Gap variable."""
    v = var.lower()

    if "n1" in v:
        cycle = "First NDC"
    elif "n2" in v:
        cycle = "Second NDC"
    else:
        cycle = "NDC"

    if v.startswith("gap_abs_"):
        measure = "Absolute Gap-to-Target"
    elif v.startswith("gap_asinh_"):
        measure = "Asinh absolute Gap-to-Target"
    else:
        measure = "Relative Gap-to-Target (%)"

    qualifiers = []
    if "uncond" in v:
        qualifiers.append("unconditional")
    if "cond" in v and "uncond" not in v:
        qualifiers.append("conditional")
    if "rcp45" in v:
        qualifiers.append("RCP4.5-SSP2 BAU median")
    if "rcp85" in v:
        qualifiers.append("RCP8.5-SSP2 BAU median")
    if "oecd" in v:
        qualifiers.append("RCP6-SSP2 OECD BAU")
    if "pik" in v:
        qualifiers.append("RCP6-SSP2 PIK BAU")
    if "iiasa" in v:
        qualifiers.append("RCP6-SSP2 IIASA BAU")

    label = f"{measure}, {cycle}"
    if qualifiers:
        label += ", " + ", ".join(qualifiers)

    # Stata variable labels are limited to 80 characters.
    return label[:80]

for col in final_gap_cols:
    final_labels[col] = make_final_gap_label(col)

def make_final_ambition_label(var):
    v = var.lower()
    cycle = "First NDC" if "n1" in v else ("Second NDC" if "n2" in v else "NDC")
    q = []
    if "uncond" in v: q.append("unconditional")
    elif "cond" in v: q.append("conditional")
    if "rcp45" in v: q.append("RCP4.5-SSP2 BAU median")
    if "rcp85" in v: q.append("RCP8.5-SSP2 BAU median")
    if "oecd" in v: q.append("RCP6-SSP2 OECD BAU")
    if "pik" in v: q.append("RCP6-SSP2 PIK BAU")
    if "iiasa" in v: q.append("RCP6-SSP2 IIASA BAU")
    label = f"Ambition (%), {cycle}"
    if q: label += ", " + ", ".join(q)
    return label[:80]

for col in final_ambition_cols:
    final_labels[col] = make_final_ambition_label(col)

# Stata labels for implementation scores are intentionally omitted because
# no Implementation_* or impl_* variable is constructed or exported.

# Coverage-cluster variables used for heterogeneity / clustering analysis.
final_labels.update({
    "audit_schema_version": "Shared Harmonizer/Builder/Figures audit revision",
    "coverage_cluster_n1": "NDC1 target coverage: 1 economy-wide, 2 broad, 3 limited, 4 single",
    "coverage_class_n1": "First NDC quantified-target sector coverage class",
    "coverage_cluster_n2": "NDC2 target coverage: 1 economy-wide, 2 broad, 3 limited, 4 single",
    "coverage_class_n2": "Second NDC quantified-target sector coverage class",
})

unlabeled_final = [
    c for c in final_panel.columns
    if c not in final_labels
]
if unlabeled_final:
    raise ValueError(
        "Unlabeled variables in final dataset: "
        + ", ".join(unlabeled_final)
    )

too_long_final_labels = {
    k: v for k, v in final_labels.items()
    if k in final_panel.columns and len(v) > 80
}
if too_long_final_labels:
    raise ValueError(
        "Stata labels exceeding 80 characters: "
        + "; ".join(f"{k}={len(v)}" for k, v in too_long_final_labels.items())
    )

# 11. EXPORTS COMPLEMENTAIRES EXCEL / CSV / STATA

# 11.1 Panel par cycle

cycle_panel.to_excel(
    OUT_CYCLE_PANEL_XLSX,
    index=False,
    engine="openpyxl"
)

# 11.2 NDC targets

ndc_targets.to_excel(
    OUT_TARGETS_XLSX,
    index=False,
    engine="openpyxl"
)
ndc_targets.to_csv(
    OUT_TARGETS_CSV,
    index=False,
    encoding="utf-8-sig"
)

targets_stata = prepare_for_stata_export(ndc_targets)
targets_stata, _ = stata_safe_varnames(targets_stata)
targets_stata.to_stata(
    OUT_TARGETS_DTA,
    write_index=False,
    version=118
)

# 11.3 Audit

audit.to_excel(
    OUT_AUDIT_XLSX,
    index=False,
    engine="openpyxl"
)

audit_stata = prepare_for_stata_export(audit)
audit_stata, _ = stata_safe_varnames(audit_stata)
audit_stata.to_stata(
    OUT_AUDIT_DTA,
    write_index=False,
    version=118
)

# 11.4 Master Stata : toutes les variables NDC1/NDC2

master_stata = prepare_for_stata_export(cycle_panel)
master_stata, _ = stata_safe_varnames(master_stata)

master_labels = {
    "iso3": "ISO3 country code",
    "year": "Year",
    "E_obs": "National EDGAR GHG emissions (diagnostic common series), Gg CO2eq",
    "gap_n1": "Gap-to-Target ratio, First NDC",
    "gap_n1_uncond": "Gap-to-Target ratio, First NDC unconditional",
    "gap_n1_cond": "Gap-to-Target ratio, First NDC conditional",
    "gap_n2": "Gap-to-Target ratio, Second NDC",
    "gap_n2_uncond": "Gap-to-Target ratio, Second NDC unconditional",
    "gap_n2_cond": "Gap-to-Target ratio, Second NDC conditional",
    "gap_abs_n1": "Absolute Gap-to-Target, First NDC",
    "gap_abs_n2": "Absolute Gap-to-Target, Second NDC",
    "gap_asinh_n1": "Asinh absolute Gap-to-Target, First NDC",
    "gap_asinh_n2": "Asinh absolute Gap-to-Target, Second NDC",
}

master_labels_complete = dict(master_labels)
for col in master_stata.columns:
    if col not in master_labels_complete:
        readable = (
            col.replace("_", " ")
               .replace("n1", "First NDC")
               .replace("n2", "Second NDC")
               .replace("rcp45", "RCP4.5 SSP2")
               .replace("rcp85", "RCP8.5 SSP2")
               .replace("oecd", "OECD")
               .replace("pik", "PIK")
               .replace("iiasa", "IIASA")
        )
        master_labels_complete[col] = readable[:80]

master_stata.to_stata(
    OUT_MASTER_DTA,
    write_index=False,
    version=118,
    variable_labels=safe_stata_labels(
        master_labels_complete, master_stata.columns
    )
)

# 11.5 Analysis Stata : jeu compact pour économétrie

analysis_stata = prepare_for_stata_export(analysis_panel)
analysis_stata, _ = stata_safe_varnames(analysis_stata)

analysis_labels = {
    "iso3": "ISO3 country code",
    "year": "Year",
    "E_obs": "National EDGAR GHG emissions (diagnostic common series), Gg CO2eq",

    "gap_n1": "Relative Gap-to-Target, First NDC",
    "gap_n1_uncond": "Relative Gap-to-Target, First NDC unconditional",
    "gap_n1_cond": "Relative Gap-to-Target, First NDC conditional",

    "gap_n2": "Relative Gap-to-Target, Second NDC",
    "gap_n2_uncond": "Relative Gap-to-Target, Second NDC unconditional",
    "gap_n2_cond": "Relative Gap-to-Target, Second NDC conditional",

    "gap_abs_n1": "Absolute Gap-to-Target, First NDC",
    "gap_abs_n1_uncond": "Absolute Gap-to-Target, First NDC unconditional",
    "gap_abs_n1_cond": "Absolute Gap-to-Target, First NDC conditional",

    "gap_abs_n2": "Absolute Gap-to-Target, Second NDC",
    "gap_abs_n2_uncond": "Absolute Gap-to-Target, Second NDC unconditional",
    "gap_abs_n2_cond": "Absolute Gap-to-Target, Second NDC conditional",

    "gap_asinh_n1": "Asinh absolute Gap-to-Target, First NDC",
    "gap_asinh_n1_uncond": "Asinh absolute Gap-to-Target, First NDC unconditional",
    "gap_asinh_n1_cond": "Asinh absolute Gap-to-Target, First NDC conditional",

    "gap_asinh_n2": "Asinh absolute Gap-to-Target, Second NDC",
    "gap_asinh_n2_uncond": "Asinh absolute Gap-to-Target, Second NDC unconditional",
    "gap_asinh_n2_cond": "Asinh absolute Gap-to-Target, Second NDC conditional",

    "gap_n1_rcp45": "Relative Gap First NDC, RCP4.5-SSP2 median",
    "gap_n1_rcp85": "Relative Gap First NDC, RCP8.5-SSP2 median",
    "gap_n1_oecd": "Relative Gap First NDC, RCP6-SSP2 OECD",
    "gap_n1_pik": "Relative Gap First NDC, RCP6-SSP2 PIK",
    "gap_n1_iiasa": "Relative Gap First NDC, RCP6-SSP2 IIASA",

    "gap_n2_rcp45": "Relative Gap Second NDC, RCP4.5-SSP2 median",
    "gap_n2_rcp85": "Relative Gap Second NDC, RCP8.5-SSP2 median",
    "gap_n2_oecd": "Relative Gap Second NDC, RCP6-SSP2 OECD",
    "gap_n2_pik": "Relative Gap Second NDC, RCP6-SSP2 PIK",
    "gap_n2_iiasa": "Relative Gap Second NDC, RCP6-SSP2 IIASA",

    "E_obs_n1": "Observed GHG emissions harmonised to First NDC sector/LULUCF scope, Gg CO2eq",
    "E_obs_n2": "Observed GHG emissions harmonised to Second NDC sector/LULUCF scope, Gg CO2eq",
    "coverage_cluster_n1": "NDC1 coverage: 1 economy-wide, 2 broad, 3 limited, 4 single",
    "coverage_class_n1": "First NDC quantified-target sector coverage class",
    "coverage_cluster_n2": "NDC2 coverage: 1 economy-wide, 2 broad, 3 limited, 4 single",
    "coverage_class_n2": "Second NDC quantified-target sector coverage class",
}

analysis_labels_complete = dict(analysis_labels)
for col in analysis_stata.columns:
    if col not in analysis_labels_complete:
        readable = (
            col.replace("_", " ")
               .replace("n1", "First NDC")
               .replace("n2", "Second NDC")
               .replace("rcp45", "RCP4.5 SSP2")
               .replace("rcp85", "RCP8.5 SSP2")
               .replace("oecd", "OECD")
               .replace("pik", "PIK")
               .replace("iiasa", "IIASA")
        )
        analysis_labels_complete[col] = readable[:80]

analysis_value_labels = {}
for _var in ["coverage_cluster_n1", "coverage_cluster_n2"]:
    if _var in analysis_stata.columns:
        analysis_stata[_var] = pd.to_numeric(
            analysis_stata[_var], errors="coerce"
        ).astype(float)
        analysis_value_labels[_var] = COVERAGE_CODE_LABELS

analysis_stata.to_stata(
    OUT_ANALYSIS_DTA,
    write_index=False,
    version=118,
    variable_labels=safe_stata_labels(
        analysis_labels_complete, analysis_stata.columns
    ),
    value_labels=analysis_value_labels
)

# 11.6 FINAL dataset : ISO3 + year + Gap + Ambition variables

final_stata = prepare_for_stata_export(final_panel)
final_stata, _ = stata_safe_varnames(final_stata)

# Rebuild label mapping after Stata-safe renaming if necessary.
final_stata_labels = {
    c: final_labels[c]
    for c in final_stata.columns
    if c in final_labels
}

# Safety check: every final Stata variable must have a label.
missing_stata_labels = [
    c for c in final_stata.columns
    if c not in final_stata_labels
]
if missing_stata_labels:
    raise ValueError(
        "Final Stata variables without labels: "
        + ", ".join(missing_stata_labels)
    )

# Stata value labels for the numeric coverage clusters.
final_value_labels = {}
for _var in ["coverage_cluster_n1", "coverage_cluster_n2"]:
    if _var in final_stata.columns:
        # pandas/Stata requires a regular numeric dtype rather than nullable Int64.
        final_stata[_var] = pd.to_numeric(
            final_stata[_var], errors="coerce"
        ).astype(float)
        final_value_labels[_var] = COVERAGE_CODE_LABELS

final_stata.to_stata(
    OUT_FINAL_DTA,
    write_index=False,
    version=118,
    variable_labels=final_stata_labels,
    value_labels=final_value_labels
)

final_panel.to_excel(
    OUT_FINAL_XLSX,
    index=False,
    engine="openpyxl"
)

final_panel.to_csv(
    OUT_FINAL_CSV,
    index=False,
    encoding="utf-8-sig"
)

# 11.7 Country graphs Stata

graphs_stata = prepare_for_stata_export(country_graphs)
graphs_stata, _ = stata_safe_varnames(graphs_stata)
graphs_stata.to_stata(
    OUT_GRAPHS_DTA,
    write_index=False,
    version=118
)

print("\n" + "=" * 80)
print("FICHIERS CREES")
print("=" * 80)

print("\nWorkbook complet conservé :")
print("  ", OUTPUT_FILE.resolve())

print("\nExports Excel complémentaires :")
print("  ", OUT_CYCLE_PANEL_XLSX.resolve())
print("  ", OUT_TARGETS_XLSX.resolve())
print("  ", OUT_AUDIT_XLSX.resolve())

print("\nExports Stata :")
print("  ", OUT_MASTER_DTA.resolve())
print("  ", OUT_ANALYSIS_DTA.resolve())
print("  ", OUT_GRAPHS_DTA.resolve())
print("  ", OUT_TARGETS_DTA.resolve())
print("  ", OUT_AUDIT_DTA.resolve())

print("\nDATASET FINAL — ISO3 + YEAR + GAP + AMBITION + IMPLEMENTATION :")
print(f"  Pays uniques éligibles (au moins un Gap NDC1 ou NDC2) : {final_panel['iso3'].nunique()}")
print(f"  Observations finales : {len(final_panel)}")
print("  ", OUT_FINAL_DTA.resolve())
print("  ", OUT_FINAL_XLSX.resolve())
print("  ", OUT_FINAL_CSV.resolve())
print(f"  Pays retenus : {final_panel['iso3'].nunique():,}")
print(f"  Observations  : {len(final_panel):,}")
print(f"  Variables     : {len(final_panel.columns):,}")
print(
    "  Pays exclus sans aucun Gap NDC1/NDC2 : "
    f"{len(excluded_iso3):,}"
)
if excluded_iso3:
    print("  ISO3 exclus : " + ", ".join(sorted(excluded_iso3)))
print("  Labels Stata : complets pour toutes les variables finales")

print("\nExport CSV des cibles :")
print("  ", OUT_TARGETS_CSV.resolve())

print(
    "\nIMPORTANT : intensity and per-capita targets are converted with "
    "PMRCPBIE SSP2 socioeconomic scale ratios (target/base). "
    "Any remaining NEEDS_REVIEW case lacks a required input or needs "
    "manual interpretation."
)
