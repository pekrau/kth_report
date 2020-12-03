# kth_report

Produce aggregate files and plots for the yearly SciLifeLab report to KTH.
For both facility and fellows data.


## Code requirements

- Python 3
- Openpyxl Python package (read XLSX and XLSM files)
- XlsxWriter Python package (create XLSX files)
- Plotly Python package (create plots)
- kaleido Python package (create PNG of plots)


## Source code organisation

The source code for each year is in its own subdirectory.

The code for 2019 was written to recreate the merged files for that year.
This code did not exist in this form when the actual reports were processed.
See https://github.com/senthil10/dc_reporting_scripts instead.

The source code for 2020 is new and used for production.


## Facilities

### Specifications

The specifications for the output files A-H are given in the document
`Data files for KTH and Infra Reports.xlsx` for each year.


## Input files

The Reporting Portal https://reportings.scilifelab.se/ is the primary
data source. The "Aggregate" feature is used to create XLSX files
combining the data from the reports based on the relevant form.

In addition, the Volume Data files attached to each report is also used.

To perform the operations below, you need to be logged in as an admin
or staff account at https://reportings.scilifelab.se/


## Create the aggregate files

1. Go to the form for this year's Facility reports in the Reporting Portal.

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

1. Go to the list of all reports.

2. Filter by the appropriate form, and the status "Submitted".

3. Download manually the volume data files for all reports to the
   subdirectory `volume_data_files` for this year's report data.


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

6. The script `merge_E.py` produces the file
   `E_Infrastructure Users {year}.xlsx` from the contents of the
   sheet `A. Users` in all the volume data files.

   NOTE: Some of the `XLSX`/`XLSM` files cause "UserWarning" when read
   by `openpyxl`. This can be ignored.

7. The script `merge_F.py` produces the file
   `F_Infrastructure Courses {year}.xlsx` from the contents of the
   sheet `B. Courses` in all the volume data files.

   NOTE: Some of the `XLSX`/`XLSM` files cause "UserWarning" when read
   by `openpyxl`. This can be ignored.

8. The script `merge_G.py` produces the file
   `G_Infrastructure Conferences Symposia Seminars {year}.xlsx` from
   the contents of the sheet `C. Conf, symp, semin` in all the volume
   data files.

   NOTE: Some of the `XLSX`/`XLSM` files cause "UserWarning" when read
   by `openpyxl`. This can be ignored.

9. The script `merge_H.py` produces the file
   `H_Infrastructure External Collaborations {year}.xlsx` from
   the contents of the sheet `D. External Collab ` (yes, there is a
   trailing white-space in the name!)  in all the volume data files.

   NOTE: Some of the `XLSX`/`XLSM` files cause "UserWarning" when read
   by `openpyxl`. This can be ignored.


## Fellows

### Input files

The Reporting Portal https://reportings.scilifelab.se/ is the primary
data source. The "Aggregate" feature is used to create XLSX files
combining the data from the reports based on the relevant form.

In addition, the Volume Data files attached to each report is also used.

To perform the operations below, you need to be logged in as an admin
or staff account at https://reportings.scilifelab.se/


## Create the aggregate files

1. Go to the form for this year's Fellow reports in the Reporting Portal.

2. Click on the button "Aggregate".

3. Check the following fields for output.
   - **Report status filter**: Submitted
   - **Report metadata**: identifier, title
   - **Report history**: [none]
   - **Report owner**: [all]
   - **Report fields**: [all except volume_data]
   - **File format**: Excel (XLSX)

4. Create and download the aggregate file by clicking the button "Aggregate".

5. Move these XLSX files to the subdirectory `aggregate_files` for
   this year's report data.


### Download the Volume data files

1. Go to the list of all reports.

2. Filter by the appropriate form, and the status "Submitted".

3. Download manually the volume data files for all reports to the
   subdirectory `volume_data_files` for this year's report data.


### Create the merged files

1. Check and set the parameters for the filepaths and other data in
   the source code file `merge_fellows.py`.

2. Run the script `merge_fellows.py`, which produces the files
   `Fellows {year} Teaching.xlsx`, 'Fellows {year} Grants.xlsx'
   and 'Fellows {year} Collaborations.xlsx' from the volume data files.

   NOTE: Some of the `XLSX`/`XLSM` files cause "UserWarning" when read
   by `openpyxl`. This can be ignored.

