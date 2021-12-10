"""Facility Reports 2020.

Create the file 'F_Infrastructure Courses 2020.xlsx'
"""

import datetime
import json
import os.path

import xlsxwriter

import facility_data

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(facility_data.BASEDIRPATH, "merged_files")

### Full file name for the F file.
FILENAME = "F_Infrastructure Courses 2020.xlsx"


def merge_F(filepath):
    """Create the F file, containing all facility courses.
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
    ws.set_column(3, 3, 60, long_text_format)
    ws.set_column(4, 4, 20, normal_text_format)
    ws.set_column(5, 5, 30, long_text_format)
    ws.set_column(6, 8, 20, long_text_format)
    ws.set_column(9, 9, 40, long_text_format)

    ws.write_row(0, 0,
                 ["1. Name of reporting unit*",
                  "2. Platform",
                  "3. Your e-mail address*",
                  "4. Full name of the course*",
                  "5a. Did the reporting unit organize or co-organize the course?*",
                  "5b. If co-organized, with whom?",
                  "6. Start date*",
                  "7. End date*",
                  "8. Location (city) of the course*",
                  "9. Comment"])

    records = facility_data.get_volume_data("B. Courses")
    key = "1. Name of reporting unit* (choose from drop-down menu)"
    iso = "%Y-%m-%d"

    for row, record in enumerate(records, 1):
        try:
            facility = record[key].strip()
            try:
                platform = facility_data.PLATFORM_LOOKUP[facility]
            # Bizarrely, sometimes the facility name has wrong character case.
            except KeyError:
                facility, platform = facility_data.PLATFORM_LOOKUP_LOWER[facility.lower()]
            # Arggh, another strangeness; most of the time 'datetime' instance,
            # but sometimes an 'int'. Try to handle it...
            start = record["5. Start date* (yyyy-mm-dd)"]
            if isinstance(start, datetime.datetime):
                start = start.strftime(iso)
            elif isinstance(start, int):
                year = start // 10000
                month = start // 100 % 100
                day = start % 100
                start = f"{year:4d}-{month:02d}-{day:02d}"
            else:
                raise ValueError("unknown start type '%s'" % type(start))

            end = record["6. End date* (yyyy-mm-dd)"]
            if isinstance(end, datetime.datetime):
                end = end.strftime(iso)
            elif isinstance(end, int):
                year = end // 10000
                month = end // 100 % 100
                day = end % 100
                end = f"{year:4d}-{month:02d}-{day:02d}"
            else:
                raise ValueError("unknown end type '%s'" % type(end))

            rowdata = [facility,
                       platform,
                       record["2. Your e-mail address*"].lower(),
                       record["3. Full name of the course*"],
                       record["4a. Did the reporting unit organize or co-organize the course?*"],
                       record["4b. If co-organized, with whom?"],
                       start,
                       end,
                       record["7. Location (city) of the course*"],
                       record["8. Comment"]]
        except (ValueError, KeyError) as error:
            print(row)
            print(json.dumps(record, indent=2))
            raise
        ws.write_row(row, 0, rowdata)

    wb.close()

    
if __name__ == "__main__":
    merge_F(os.path.join(DIRPATH, FILENAME))
