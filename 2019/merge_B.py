"""Facility Reports 2019.

Create the file 'B_Infrastructure FD and HF 2019.xlsx'
"""

import os.path

import xlsxwriter

import facility_data

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(facility_data.BASEDIRPATH, "merged_files")

### Full file name for the B file.
FILENAME = "B_Infrastructure FD and HF 2019.xlsx"


def merge_B(filepath):
    """Create the B file, containing fields for the Facility director
    and Head of Facility, collected from the table fields in the
    infrastructure facility reports.
    """
    director_data = facility_data.get_facility_director_data()
    head_data = facility_data.get_facility_head_data()

    # Reformat data into one row per facility
    report_data = []
    for facility, platform in facility_data.PLATFORM_LOOKUP.items():
        rowdata = [facility, platform]

        # Facility director data first.
        first_names = [r["facility_director: First name"]
                       for r in director_data if r["facility"] == facility]
        rowdata.append("\n".join(first_names))
        last_names = [r["facility_director: Last name"]
                      for r in director_data if r["facility"] == facility]
        rowdata.append("\n".join(last_names))
        emails = [r["facility_director: Email address"]
                  for r in director_data if r["facility"] == facility]
        rowdata.append("\n".join(emails))
        affiliations = [r["facility_director: Affiliation (University)"]
                        for r in director_data if r["facility"] == facility]
        rowdata.append("\n".join(affiliations))

        # Facility head data second.
        first_names = [r["facility_head: First name"]
                       for r in head_data if r["facility"] == facility]
        rowdata.append("\n".join(first_names))
        last_names = [r["facility_head: Last name"]
                      for r in head_data if r["facility"] == facility]
        rowdata.append("\n".join(last_names))
        emails = [r["facility_head: Email address"]
                  for r in head_data if r["facility"] == facility]
        rowdata.append("\n".join(emails))
        affiliations = [r["facility_head: Affiliation (University)"]
                        for r in head_data if r["facility"] == facility]
        rowdata.append("\n".join(affiliations))

        report_data.append(rowdata)

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
    ws.freeze_panes(2, 2)
    ws.set_row(0, None, head_text_format)
    ws.set_row(1, None, head_text_format)
    ws.merge_range(0, 0, 1, 0, "Facility")
    ws.merge_range(0, 1, 1, 1, "Platform")
    ws.merge_range(0, 2, 0, 5, "Facility director")
    ws.merge_range(0, 6, 0, 9, "Facility heads")
    ws.write_row(1, 2, ["First Name", "Last Name", "Email", "Affliation"])
    ws.write_row(1, 6, ["First Name", "Last Name", "Email", "Affliation"])

    ws.set_column(0, 1, 40, long_text_format)
    ws.set_column(2, 3, 20, long_text_format)
    ws.set_column(4, 4, 40, long_text_format)
    ws.set_column(5, 7, 20, long_text_format)
    ws.set_column(8, 8, 40, long_text_format)
    ws.set_column(9, 9, 20, long_text_format)

    for row, rowdata in enumerate(report_data, 2):
        ws.write_row(row, 0, rowdata)

    wb.close()

    
if __name__ == "__main__":
    merge_B(os.path.join(DIRPATH, FILENAME))
