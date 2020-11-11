# kth_report

Code to produce plots for the yearly SciLifeLab report to KTH.

Requires Python 3. Uses the Plotly library, the Python version.

Input files are located in a directory, and output files are created
in another.

Input files are:

1) Aggregated reports files (XLSX), combining data from each reporting
   unit (facility) in the Reporting Portal
   https://reportings.scilifelab.se/. Use the "Aggregate" feature for
   the relevant form in the portal.
2) Volume data XLSX (or XLSM) files from each reporting unit
   (facility) from within each report in the Reporting Portal.
