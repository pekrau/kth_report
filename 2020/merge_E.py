"""Facility Reports 2020.

Create the file 'E_Infrastructure Users 2020.xlsx'
"""

import json
import os.path

import xlsxwriter

import facility_data

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(facility_data.BASEDIRPATH, "merged_files")

### Full file name for the E file.
FILENAME = "E_Infrastructure Users 2020.xlsx"


def merge_E(filepath):
    """Create the E file, containing all facility users.
    """
    wb = xlsxwriter.Workbook(filepath)

    head_text_format = wb.add_format({'bold':True,
                                      'text_wrap':True,
                                      'bg_color':'#9ECA7F',
                                      'font_size':15,
                                      'align':'center',
                                      'border':1})
    normal_text_format = wb.add_format({'font_size':14,
                                        'align':'left',
                                        'valign':'vcenter'})
    long_text_format = wb.add_format({'text_wrap':True,
                                      'font_size':14,
                                      'align':'left',
                                      'valign':'vcenter'})

    ws = wb.add_worksheet()
    ws.freeze_panes(1, 2)
    ws.set_row(0, None, head_text_format)
    ws.set_column(0, 2, 40, normal_text_format)
    ws.set_column(3, 4, 20, normal_text_format)
    ws.set_column(5, 7, 40, normal_text_format)

    ws.write_row(0, 0,
                 ["1. Name of reporting unit*",
                  "2. Platform",
                  "3. Your e-mail address*",
                  "4a. First name of the responsible PI*",
                  "4b. Surname of the responsible PI*",
                  "5. E-mail address of responsible PI*",
                  "6a. Affiliation of PI: Specific university or category*",
                  "6b. For non-specific universities and categories"
                  " in 5a, name the organization"])

    records = facility_data.get_volume_data("A. Users")
    for row, record in enumerate(records, 1):
        try:
            facility = record["1.  Name of reporting unit*"
                              " (choose from drop-down menu)"].strip()
            try:
                platform = facility_data.PLATFORM_LOOKUP[facility]
            # Bizarrely, sometimes the facility name has wrong character case.
            except KeyError:
                facility, platform = facility_data.PLATFORM_LOOKUP_LOWER[facility.lower()]
            rowdata = [facility,
                       platform,
                       record["2. Your e-mail address*"].lower(),
                       record["3a. First name of the responsible PI*"],
                       record["3b. Surname of the responsible PI*"],
                       record["4. E-mail address of responsible PI*"].lower(),
                       record["5a. Affiliation of PI: Specific university or"
                              " category (choose from drop-down menu)*"],
                       record["5b. For non-specific universities and categories"
                              " in 5a, name the organization (free text)"]]
        except KeyError as error:
            print(row)
            print(json.dumps(record, indent=2))
            raise
        ws.write_row(row, 0, rowdata)

    wb.close()

    
if __name__ == "__main__":
    merge_E(os.path.join(DIRPATH, FILENAME))
