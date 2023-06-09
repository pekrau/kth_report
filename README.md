# kth_report

Produce merged files and plots for the yearly SciLifeLab report to KTH.
For both unit (facility) and fellows data.


## Code requirements

- Python 3
- Openpyxl Python package (read XLSX and XLSM files)
- XlsxWriter Python package (create XLSX files)
- Plotly Python package (create plots)
- kaleido Python package (create PNG of plots)

Tip: Create a virtual Python environment and install the packages using pip
and the `requirements.txt` file which can be found in the source code for each year:

    $ pip install -r requirements.txt


## Source code organisation

The source code for each year is in its own subdirectory.

### 2022

The source code for 2022 is an adjusted copy of 2021 and is used for production.


### 2021

The source code for 2021 is an adjusted copy of 2020 and is used for production.

### 2020

The source code for 2020 is new and is used for production.
**NOTE** that the new source code does not use HTML or explicit
JavaScript. It is pure Python.

### 2019

The code for 2019 was written to recreate the merged files for that year.
This code did not exist in this form when the actual reports were processed.
See https://github.com/senthil10/dc_reporting_scripts instead.


## Units (Facilities)

### Specifications

The specifications for the output files A-H are given in the document
`Data files for KTH and Infra Reports.xlsx` for each year.


## Input files

The Reporting Portal https://reporting.scilifelab.se/ is the primary
data source. The "Aggregate" feature is used to create XLSX files
combining the data from the reports based on the relevant form.

In addition, the Volume Data files attached to each report is also used.

To perform the operations below, you need to be logged in as an admin
or staff account at `https://reporting.scilifelab.se/`

My location for the input files are `~/Nextcloud/Årsrapport 2022/`

## Create the aggregate files

1. Go to the form for this year's Infrastructure Unit (Facility)
   reports in the Reporting Portal.

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
   perform one aggregate operation for *each value* of **Table field**;
   "facility_director", "facility_head", "additional_funding" and
   "immaterial_property_rights".

6. Move these XLSX files to the subdirectory `Units reports/aggregate_files` for
   this year's report data.


### Download the Volume data files

1. Go to the page listing all reports
   [https://reporting.scilifelab.se/orders](https://reporting.scilifelab.se/orders)

2. Filter by the appropriate form and status "Submitted".

3. Download manually the volume data files for all reports to the
   subdirectory `volume_data_files` for this year's report data.

  1. Right-click on the report link, to bring up a new tab (keeping
     the list as is in its tab).
  2. Click on the link to the Volume Data file; it is visible under "Files" in the
     top panel, and also further down in the field "Volume data by Excel..."
     It is the same file.


### Create the merged files

1. Check and set the parameters for the filepaths and other data in
   the source code file `facility_data.py`. This file is imported as a module
   by the other scripts and its variables are used there.

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


### Create the plot named "Figure 5" (Affiliations of users of SciLifeLab units)

The Python script `make_fig5.py` uses the Python library of Plotly and
a few other packages to produce a large PNG file of a scatterplot of
affiliations versus facilities, where the size of each circle shows
the number of unique users.

The input file is that produced by the script `merge_E.py` (step 6 above).
(This is an improvement from 2020; no manual cut-and-paste is required now.)

This script also creates a CSV file containing the counts.


## SciLifeLab Fellows

### Input files

The Reporting Portal https://reporting.scilifelab.se/ is the primary
data source. The "Aggregate" feature is used to create XLSX files
combining the data from the reports based on the relevant form.

In addition, the Volume Data files attached to each report is also used.

To perform the operations below, you need to be logged in as an admin
or staff account at https://reporting.scilifelab.se/


## Create the aggregate file.

1. Go to the form for this year's SciLifeLab Fellow reports in the Reporting Portal.

2. Click on the button "Aggregate".

3. Check the following fields for output.
   - **Report status filter**: Submitted
   - **Report metadata**: identifier, title
   - **Report history**: [none]
   - **Report owner**: [all]
   - **Report fields**: [all except volume_data]
   - **File format**: Excel (XLSX)

4. Create and download the aggregate file by clicking the button "Aggregate".

5. Move this XLSX file to the subdirectory `aggregate_files` for
   this year's report data.


### Download the Volume data files

1. Go to the list of all reports
   [https://reporting.scilifelab.se/orders](https://reporting.scilifelab.se/orders)

2. Filter by the appropriate form, and the status "Submitted".

3. Download manually the volume data files for all reports to the
   subdirectory `volume_data_files` for this year's report data.

  1. Right-click on the report link, to bring up a new tab (keeping
     the list as is in its tab).
  2. Click on the link to the Volume Data file; it is visible under "Files" in the
     top panel, and also further down in the field "Volume data by Excel..."
     It is the same file.


### Create the merged files

1. Check and set the parameters for the filepaths and other data in
   the source code file `merge_scilifelab_fellows.py`. This file is
   imported as a module by the other scripts and its variables are
   used there.

2. Run the script `merge_scilifelab_fellows.py`, which produces the files
   `Fellows {year} Teaching.xlsx`, `Fellows {year} Grants.xlsx`
   and `Fellows {year} Collaborations.xlsx` from the volume data files.

   NOTE: Some of the `XLSX`/`XLSM` files cause "UserWarning" when read
   by `openpyxl`. This can be ignored.


## DDLS Fellows

This procedure was copied 2022 from the SciLifeLab Fellows above.

### Input files

The Reporting Portal https://reporting.scilifelab.se/ is the primary
data source. The "Aggregate" feature is used to create XLSX files
combining the data from the reports based on the relevant form.

In addition, the Volume Data files attached to each report is also used.

To perform the operations below, you need to be logged in as an admin
or staff account at https://reporting.scilifelab.se/


## Create the aggregate file.

1. Go to the form for this year's DDLS Fellow reports in the Reporting Portal.

2. Click on the button "Aggregate".

3. Check the following fields for output.
   - **Report status filter**: Submitted
   - **Report metadata**: identifier, title
   - **Report history**: [none]
   - **Report owner**: [all]
   - **Report fields**: [all except volume_data]
   - **File format**: Excel (XLSX)

4. Create and download the aggregate file by clicking the button "Aggregate".

5. Move this XLSX file to the subdirectory `aggregate_files` for
   this year's report data.


### Download the Volume data files

1. Go to the list of all reports.

2. Filter by the appropriate form, and the status "Submitted".

3. Download manually the volume data files for all reports to the
   subdirectory `volume_data_files` for this year's report data.


### Create the merged files

1. Check and set the parameters for the filepaths and other data in
   the source code file `merge_ddls_fellows.py`.

2. Run the script `merge_ddls_fellows.py`, which produces the files
   `DDLS Fellows {year} Teaching.xlsx`, `DDLS Fellows {year} Grants.xlsx`
   and `DDLS Fellows {year} Collaborations.xlsx` from the volume data files.

   NOTE: Some of the `XLSX`/`XLSM` files cause "UserWarning" when read
   by `openpyxl`. This can be ignored.
