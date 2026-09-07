# Validation report

## Clean-run test

The complete workflow was executed from a newly created Python 3.12 virtual
environment using `requirements.lock`. All nine stages completed successfully,
including the final checks for required datasets, figures, audit tables, and
the absence of PDF files from the publication-figure directories.

## Paris indicators

The reconstructed Paris panel contains 1,640 country-year observations for the
164 countries retained in the analytical sample. Relative to the previously
archived final panel:

- `gap_n1` is reproduced exactly, including missing observations;
- `gap_n2` is reproduced exactly, including missing observations; and
- the country coverage remains 138 countries for NDC1 and 148 for NDC2.

## Kyoto indicators

The reconstructed Kyoto panel contains 180 country-year observations. The
benchmark and LULUCF-adjusted Gap variants are reproduced exactly. The current
PDF extractor additionally recovers three 2010 ITL reporting observations for
Ireland, Italy, and Japan that were absent from the earlier stored transaction
intermediate. Consequently, the transaction-adjusted and fully adjusted
variants differ slightly for these observations. The largest country-year
absolute differences are 1.0834 and 1.0441 percentage points, respectively.

This is an intentional source-level correction rather than a stochastic
replication discrepancy: the regenerated values correspond to the entries in
the included ITL reports. At the annual-median level, the resulting difference
is confined to a small change in the 2011 transaction-adjusted median
(-9.0106 instead of -9.0973 percent); the remaining compared annual medians are
unchanged.

## Repository checks

- No user-specific absolute path is required by the workflow.
- All publication figures are generated in PNG format only.
- The prospective pathway input remains compressed and is read directly.
- Generated directories are rebuilt from scratch on each complete run.
- The Paris Appendix audit table is restricted to the final 164-country sample.

## Ambition--alignment transition checks

The country-level transition figure uses the common sample of 122 countries
with both Gap-to-Target and ambition indicators under NDC1 and NDC2. Its audit
tables verify the following mutually exclusive transitions: 21 countries
remain below their annual pathways, 24 move from below to above, 6 move from
above to below, and 71 remain above. The four categories sum to the complete
common sample. Major target-architecture changes are flagged separately for 39
countries and do not alter the transition counts.

The ambition--alignment figures use zero as the vertical ambition reference,
display ISO3 labels for all observations, and label the non-parametric curve as
``LOWESS fit''. The former aggregate alignment and transition figure is retained
as an Appendix output.

## Detailed regional figure

The original seven-region Appendix figure is retained. An additional figure
reports nine geographic groups by separating Europe from Central Asia and the
Middle East from North Africa. Afghanistan and Pakistan are classified in
South Asia, Djibouti in Sub-Saharan Africa, and Malta in Europe. Sample sizes
are displayed separately for NDC1 and NDC2, and the interval bars retain the
range across the main estimate and five alternative specifications.
