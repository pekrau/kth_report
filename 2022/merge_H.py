"""Infrastructure Units Reports 2022.

Create the file 'H_Infrastructure External Collaborations 2022.xlsx'

This code is identical to the 2021 code.
"""

import datetime
import json
import os.path

import xlsxwriter

import facility_data

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(facility_data.BASEDIRPATH, "merged_files")

### Full file name for the H file.
FILENAME = "H_Infrastructure External Collaborations 2022.xlsx"


def merge_H(filepath):
    """Create the H file, containing all external collaborations.
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
    ws.set_column(0, 5, 40, long_text_format)
    ws.set_column(6, 6, 50, long_text_format)

    ws.write_row(0, 0,
                 ["1. Name of reporting unit*",
                  "2. Platform",
                  "3. Your e-mail address*",
                  "4. Name of external organization*",
                  "5. Type of organization* (choose from drop-down menu)",
                  "6. Reference person",
                  "7. Purpose of collabaration/alliance*"])

    ### Madness! This sheetname has a trailing blank!
    records = facility_data.get_volume_data("D. External Collab ")
    key = "1. Name of reporting unit* (choose from drop-down menu)"

    for row, record in enumerate(records, 1):
        try:
            facility = record[key].strip()
            try:
                platform = facility_data.PLATFORM_LOOKUP[facility]
            # Bizarrely, sometimes the facility name has wrong character case.
            except KeyError:
                facility, platform = facility_data.PLATFORM_LOOKUP_LOWER[facility.lower()]
            rowdata = [facility,
                       platform,
                       record["2. Your e-mail address*"].lower(),
                       record["3. Name of external organization*"],
                       record["4. Type of organization* (choose from drop-down menu)"],
                       record["5. Reference person"],
                       record["6. Purpose of collabaration/alliance*"]]
        except (ValueError, KeyError) as error:
            print(row)
            print(json.dumps(record, indent=2))
            raise
        ws.write_row(row, 0, rowdata)

    wb.close()

    
if __name__ == "__main__":
    merge_H(os.path.join(DIRPATH, FILENAME))
