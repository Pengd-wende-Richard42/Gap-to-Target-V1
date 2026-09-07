# Data inputs

The replication uses only the source files required by the final Kyoto and
Paris pipelines. Files created by earlier exploratory or comparison exercises
are intentionally excluded.

## Kyoto Protocol

- `raw/kyoto/Annual_Accounting_quantity_Total_Firstperiod_Kyoto.xlsx` contains
  the first-commitment-period accounting quantities used to construct the
  country-year benchmark and accounting variants.
- `raw/kyoto/itl_reports/ITL_2008.pdf` through `ITL_2015.pdf` are the annual
  International Transaction Log reports. The pipeline extracts the relevant
  transaction annexes before constructing the indicators.

## Paris Agreement

- `raw/paris/NDCV1.xlsx` is the reviewed country-level NDC coding input.
- `raw/paris/EDGAR_AR5_GHG_1970_2024.xlsx` contains EDGAR emissions through
  2024, including the sector detail used to match observed emissions to each
  target perimeter.
- `raw/paris/Emissions from forests (Global, National - Annual) - FAOSTAT.xlsx`
  provides the Forest Land proxy used only when required by the audited NDC
  perimeter.
- `raw/paris/PMRCPBIE_04Feb20.zip` contains the country-resolved emissions and
  socioeconomic pathways used for targets requiring prospective inputs. The
  CSV is stored as a single-file ZIP to remain below GitHub's individual-file
  size limit; pandas reads it directly without creating an intermediate file.
- `raw/paris/ICR.xlsx` contains the country metadata used in the figures.

## Geographic boundaries

`raw/geospatial/ne_110m_admin_0_countries.zip` contains the Natural Earth
country boundaries used by the maps. Keeping this archive in the repository
makes map generation independent of network availability.

Users remain responsible for observing the licences and attribution conditions
of the original data providers. The repository documents and reproduces the
transformations; it does not alter ownership of the underlying source data.
