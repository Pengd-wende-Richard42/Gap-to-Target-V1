"""Run the complete Gap-to-Target replication pipeline.

All paths are resolved relative to the repository root. The program writes the
full output of each stage to ``logs/replication.log`` while keeping the command
window concise and informative.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOG_FILE = PROJECT_ROOT / "logs" / "replication.log"

RAW_FILES = (
    "data/raw/kyoto/Annual_Accounting_quantity_Total_Firstperiod_Kyoto.xlsx",
    "data/raw/paris/NDCV1.xlsx",
    "data/raw/paris/EDGAR_AR5_GHG_1970_2024.xlsx",
    "data/raw/paris/Emissions from forests (Global, National - Annual) - FAOSTAT.xlsx",
    "data/raw/paris/ICR.xlsx",
    "data/raw/paris/PMRCPBIE_04Feb20.zip",
    "data/raw/geospatial/ne_110m_admin_0_countries.zip",
)

STAGES = (
    (
        "Extracting Kyoto ITL transactions",
        "code/kyoto/01_extract_itl_transactions.py",
    ),
    (
        "Building Kyoto Gap indicators",
        "code/kyoto/02_build_kyoto_indicators.py",
    ),
    (
        "Generating Kyoto figures",
        "code/kyoto/03_make_kyoto_figures.py",
    ),
    (
        "Generating Kyoto analytical framework",
        "code/kyoto/04_make_kyoto_framework.py",
    ),
    (
        "Harmonising Paris NDC targets and emissions coverage",
        "code/paris/01_harmonize_ndc_targets.py",
    ),
    (
        "Building Paris Gap indicators",
        "code/paris/02_build_paris_indicators.py",
    ),
    (
        "Generating Paris figures and sensitivity analysis",
        "code/paris/03_make_paris_figures.py",
    ),
    (
        "Generating Paris analytical framework",
        "code/paris/04_make_paris_framework.py",
    ),
    (
        "Building country-level audit tables",
        "code/05_build_audit_tables.py",
    ),
)

EXPECTED_OUTPUTS = (
    "data/processed/kyoto/Gap_Kyoto.dta",
    "data/processed/paris/Paris_GapToTarget_Final.dta",
    "data/processed/paris/Paris_NDC_Targets.xlsx",
    "outputs/figures/kyoto/kyoto_annual_medians.png",
    "outputs/figures/kyoto/kyoto_variant_distributions.png",
    "outputs/figures/paris/Main/Figure_2_Annual_Median_Gaps_Uncertainty.png",
    "outputs/figures/paris/Main/Figure_8_Transitions_in_Ambition_and_Alignment.png",
    "outputs/figures/paris/Appendix/Appendix_Aggregate_Alignment_Positions_and_Transitions.png",
    "outputs/figures/paris/Appendix/Appendix_Ambition_and_Alignment.png",
    "outputs/figures/paris/Appendix/Appendix_Gap_by_Detailed_Region.png",
    "outputs/figures/paris/Diagnostics/ambition_alignment_transition_counts.csv",
    "outputs/figures/paris/Diagnostics/ambition_alignment_transition_audit.csv",
    "outputs/figures/paris/Diagnostics/detailed_region_classification.csv",
    "outputs/figures/paris/Diagnostics/detailed_region_summary.csv",
    "outputs/audit/country_commitment_audit_workbook.xlsx",
    "outputs/audit/paris_country_cycle_audit_appendix.tex",
)


def check_inputs() -> None:
    missing = [name for name in RAW_FILES if not (PROJECT_ROOT / name).is_file()]
    missing.extend(
        f"data/raw/kyoto/itl_reports/ITL_{year}.pdf"
        for year in range(2008, 2016)
        if not (PROJECT_ROOT / f"data/raw/kyoto/itl_reports/ITL_{year}.pdf").is_file()
    )
    if missing:
        formatted = "\n".join(f"  - {name}" for name in missing)
        raise FileNotFoundError(f"Required input files are missing:\n{formatted}")


def reset_generated_directories() -> None:
    for relative in ("data/processed", "outputs"):
        path = PROJECT_ROOT / relative
        if path.exists():
            shutil.rmtree(path)
    for relative in (
        "data/processed/kyoto",
        "data/processed/paris",
        "outputs/figures/kyoto",
        "outputs/figures/paris",
        "outputs/audit",
        "logs",
    ):
        (PROJECT_ROOT / relative).mkdir(parents=True, exist_ok=True)


def log_tail(lines: int = 30) -> str:
    try:
        content = LOG_FILE.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return ""
    return "\n".join(content[-lines:])


def run_stage(index: int, label: str, script: str, log) -> None:
    print(f"[{index:02d}/{len(STAGES):02d}] {label}...", flush=True)
    log.write(f"\n{'=' * 78}\n{label}\nScript: {script}\n{'=' * 78}\n")
    log.flush()
    result = subprocess.run(
        [sys.executable, str(PROJECT_ROOT / script)],
        cwd=PROJECT_ROOT,
        stdout=log,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if result.returncode:
        raise RuntimeError(
            f"Stage failed: {label}\n\nLast log lines:\n{log_tail()}"
        )
    print(f"         Completed: {label}", flush=True)


def validate_outputs() -> None:
    missing = [name for name in EXPECTED_OUTPUTS if not (PROJECT_ROOT / name).is_file()]
    pdf_files = list((PROJECT_ROOT / "outputs" / "figures").rglob("*.pdf"))
    if missing:
        formatted = "\n".join(f"  - {name}" for name in missing)
        raise RuntimeError(f"Expected outputs were not created:\n{formatted}")
    if pdf_files:
        formatted = "\n".join(f"  - {p.relative_to(PROJECT_ROOT)}" for p in pdf_files)
        raise RuntimeError(f"Unexpected PDF figures were created:\n{formatted}")


def main() -> int:
    try:
        check_inputs()
        reset_generated_directories()
        started = datetime.now(timezone.utc).isoformat()
        with LOG_FILE.open("w", encoding="utf-8") as log:
            log.write(f"Gap-to-Target replication started: {started}\n")
            log.write(f"Python executable: {sys.executable}\n")
            log.write(f"Project root: {PROJECT_ROOT}\n")
            for index, (label, script) in enumerate(STAGES, start=1):
                run_stage(index, label, script, log)
            log.write("\nAll replication stages completed successfully.\n")
        print("[CHECK] Validating final datasets, figures, and audit files...", flush=True)
        validate_outputs()
        print("        Completed: Output validation", flush=True)
        return 0
    except Exception as exc:
        print(f"\n[ERROR] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
