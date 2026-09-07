# Gap-to-Target replication package

This repository reproduces the data construction, indicators, sensitivity
analysis, figures, and country-level audit material for the Kyoto Protocol and
the first two Paris NDC cycles. The workflow is designed for a Windows command
prompt and runs from a single entry point.

## Quick start on Windows

1. Install Python 3.12 and enable the Python launcher during setup.
2. Clone or download this repository.
3. Open Command Prompt in the repository directory.
4. Run:

```cmd
run_all.bat
```

No personal path needs to be entered. The launcher determines the repository
location automatically, creates an isolated `.venv` environment, installs the
required Python packages, and executes every stage in the correct order.

The first run takes longer because the environment must be created and the
packages installed from the tested `requirements.lock` file. Later runs reuse
the same environment. Delete `.venv` at
any time to rebuild it without affecting the rest of the computer.

## Replication stages

The command window reports the current operation while detailed messages are
written to `logs/replication.log`.

1. **Extracting Kyoto ITL transactions** — reads the transaction annexes from
   the annual ITL reports and constructs a country-period panel.
2. **Building Kyoto Gap indicators** — merges assigned amounts, Annex A
   emissions, LULUCF accounting, and ITL transactions; then constructs the four
   Kyoto accounting specifications.
3. **Generating Kyoto figures** — produces the annual medians, distributions,
   country dumbbell plot, and map.
4. **Generating Kyoto analytical framework** — creates the Kyoto framework
   diagram used in the Appendix.
5. **Harmonising Paris NDC targets and emissions coverage** — standardises NDC
   architectures and matches observed EDGAR/FAOSTAT emissions to each target
   perimeter.
6. **Building Paris Gap indicators** — converts targets into emissions outcomes,
   constructs annual target-consistent paths, and calculates the main and
   alternative Gap-to-Target indicators.
7. **Generating Paris figures and sensitivity analysis** — creates the main and
   Appendix figures, including BAU scenario ranges, maps, the four
   ambition--alignment quadrants, and country-level transitions between NDC
   cycles.
8. **Generating Paris analytical framework** — creates the Paris framework
   diagram.
9. **Building country-level audit tables** — produces the complete CSV/XLSX
   audit files and the LaTeX Appendix tables.

The pipeline stops immediately if a required input is missing or a stage fails.
It then reports the error and preserves the detailed log for diagnosis.

## Directory structure

```text
code/                 Ordered Python scripts and the master runner
data/raw/             Source inputs required by the final methodology
data/processed/       Reconstructed analytical datasets
outputs/figures/      Publication figures in PNG format only
outputs/audit/        Country-level audit files and LaTeX tables
logs/                 Detailed replication log
```

Generated datasets and outputs are not required in version control because they
are rebuilt by `run_all.bat`. The `.gitignore` file excludes them by default,
except for the new detailed regional figure and its two audit files, which are
retained as documented publication additions.

## Main outputs

After successful execution:

- `data/processed/kyoto/Gap_Kyoto.dta` contains the four Kyoto Gap variants;
- `data/processed/paris/Paris_GapToTarget_Final.dta` contains the final Paris
  country-year indicators;
- `outputs/figures/kyoto/` contains all Kyoto PNG figures;
- `outputs/figures/paris/Main/` contains the main Paris figures;
- `outputs/figures/paris/Appendix/` contains the complementary Paris figures;
- `outputs/figures/paris/Main/Figure_8_Transitions_in_Ambition_and_Alignment.png`
  identifies country-level transitions for the common sample of 122 countries;
- `outputs/figures/paris/Appendix/Appendix_Aggregate_Alignment_Positions_and_Transitions.png`
  summarises the corresponding aggregate positions and transition shares;
- `outputs/figures/paris/Appendix/Appendix_Gap_by_Detailed_Region.png`
  complements the seven-region result by separating Europe from Central Asia
  and the Middle East from North Africa, for a total of nine subregions;
- `outputs/figures/paris/Diagnostics/ambition_alignment_transition_audit.csv`
  documents the transition and target-architecture status of every country in
  the common sample;
- `outputs/figures/paris/Diagnostics/detailed_region_classification.csv` and
  `detailed_region_summary.csv` document the nine-region country assignment
  and the values plotted in the detailed regional figure;
- `outputs/audit/country_commitment_audit_workbook.xlsx` contains the Paris and
  Kyoto country-level audit tables;
- `outputs/audit/*.tex` contains the corresponding LaTeX longtables.

## Reproducibility choices

- Every project path is relative to the repository root.
- Python dependencies are isolated from system packages.
- The prospective-pathway CSV is read directly from its compressed archive, so
  no avoidable uncompressed intermediate file is created.
- Geographic boundaries are stored locally for reproducible offline maps.
- Publication figures are saved only as PNG files.
- Generated directories are cleared at the start of each complete run to avoid
  mixing new results with stale outputs.
- Required final files are checked after the last stage.
- The local `.venv` folder is created on the user's computer and is never part
  of the distributed replication archive or the GitHub repository.

The package was tested from a clean Python 3.12 environment. The numerical and
structural checks performed before release are reported in `VALIDATION.md`.

## Troubleshooting

If the pipeline stops, consult `logs/replication.log`. To rebuild the Python
environment, close any program using it, delete `.venv`, and run `run_all.bat`
again. File synchronization software may lock Excel or image files; close those
files before rerunning the pipeline.

The ITL extraction uses Camelot's stream parser and therefore does not require
Stata. Stata-format datasets are written directly by pandas/pyreadstat. Stata is
only needed if the resulting `.dta` files are analysed separately after the
replication.

## Copyright

© 2026 NIKIEMA Pengd Wende Richard — All rights reserved.

No part of this repository, dataset, documentation, analytical framework, or
associated materials may be reproduced, distributed, modified, or used for
commercial purposes without prior written permission from the author. See the
[copyright notice](LICENSE).

## Citation

```bibtex
@article{nikiema6761581measuring,
  title   = {Measuring Alignment with National Climate Commitments: A Gap-to-Target Indicator for Kyoto and Paris NDC Cycles},
  author  = {NIKIEMA, Pengd Wend{\'e} Richard},
  journal = {Available at SSRN 6761581}
}
```
