# kth_report

Code to produce plots for the yearly SciLifeLab report to KTH.

## Code requirements

- Python 3
- Plotly Python package
- kaleido Python package
- Openpyxl Python package
- XlsxWriter Python package

## Input files

The Reporting Portal https://reportings.scilifelab.se/ is the primary
data source. The "Aggregate" feature is used to create XLSX files
combining the data from the reports based on the relevant form.

In addition, the Volume Data files attached to each report is also used.

### Create the aggregate files

1. Login using an admin or staff account at https://reportings.scilifelab.se/

2. Go to the form for this year's reports, e.g.
   https://reporting.scilifelab.se/form/c37b50f0a16c4ad2ab277e79a8902f43
   for 2019.

3. Click on the button "Aggregate".

4. Check the following fields for output.
   - **Report status filter**: Submitted
   - **Report metadata**: identifier, title
   - **Report history**: [none]
   - **Report owner**: [none]
   - **Report fields**: [all]
   - **Table field**: [none]
   - **File format**: Excel (XLSX)

5. Create and download the aggregate file by clicking the button "Aggregate".

6. Check only the "facility" value in the **Report fields**, and then
   perform an aggregate operation for each value of **Table field**;
   "facility_director", "facility_head", "additional_funding" and
   "immaterial_property_rights".

7. Move these XLSX files to the subdirectory `aggregate_files` for
   this years report data.

### Download the Volume data files


### Create the XLSX file for OO

1. Check and set the parameters for the filepaths in the source code
   file `facility_data.py`.

2. Check that the list of field identifiers for each 