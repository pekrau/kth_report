# kth_report

Code to produce plots for the yearly SciLifeLab report to KTH.

## Code requirements

- Python 3
- Plotly Python package
- kaleido Python package
- Openpyxl Python package
- XlsxWriter Python package

## Source code organisation

The source code for each year is in its own subdirectory.

The code for 2019 was written to recreate the merged files for that year.
This code did not exist in this form when the actual reports were processed.
See https://github.com/senthil10/dc_reporting_scripts instead.

The source code for 2020 is new and used for production.

## Input files

The Reporting Portal https://reportings.scilifelab.se/ is the primary
data source. The "Aggregate" feature is used to create XLSX files
combining the data from the reports based on the relevant form.

In addition, the Volume Data files attached to each report is also used.

To perform the operations below, you need to be logged in as an admin
or staff account at https://reportings.scilifelab.se/

### Create the aggregate files

1. Go to the form for this year's reports in the Reporting Portal, e.g.
   https://reporting.scilifelab.se/form/c37b50f0a16c4ad2ab277e79a8902f43
   for 2019.

2. Click on the button "Aggregate".

3. Check the following fields for output.
   - **Report status filter**: Submitted
   - **Report metadata**: identifier, title
   - **Report history**: [none]
   - **Report owner**: [none]
   - **Report fields**: [all]
   - **Table field**: [none]
   - **File format**: Excel (XLSX)

4. Create and download the aggregate file by clicking the button "Aggregate".

5. Check *only* the "facility" value in the **Report fields**, and then
   perform an aggregate operation for *each value* of **Table field**;
   "facility_director", "facility_head", "additional_funding" and
   "immaterial_property_rights".

6. Move these XLSX files to the subdirectory `aggregate_files` for
   this year's report data.

### Download the Volume data files

1. Go to the form for this year's reports in the Reporting Portal, e.g.
   https://reporting.scilifelab.se/form/c37b50f0a16c4ad2ab277e79a8902f43
   for 2019.

2. Click on the number of Reports to go to the list of all reports for
   that form.

3. Download manually the volume data files for all **submitted** reports
   to the subdirectory `volume_data_files` for this year's report
   data. Start from the list of all reports for the relevant form.

### Create the merged files

1. Check and set the parameters for the filepaths and other data in
   the source code file `facility_data.py`.

2. The script `merge_A.py` produces the file
   `A_Infrastructure Single Data Reported {year}.xlsx` from the
   contents of file `orders_Facility_report_{year}.xlsx`.

3. The script `merge_B.py` produces the file
   `B_Infrastructure FD and HF {year}.xlsx` from the contents of the
   files `orders_Facility_report_{year}_facility_director.xlsx` and
   `orders_Facility_report_{year}_facility_head.xlsx`.

4. The script `merge_C.py` produces the file
   `C_Infrastructure Other Funding {year}.xlsx` from the contents of
   the file `orders_Facility_report_{year}_additional_funding.xlsx`.

5. The script `merge_D.py` produces the file
   `D_Infrastructure Immaterial Property Rights {year}.xlsx` from the
   contents of the file
   `orders_Facility_report_{year}_immaterial_property_rights.xlsx`.
