"""Facility Reports 2019.

Create the file 'C_Infrastructure Other Funding 2019.xlsx'
"""

import os.path

import xlsxwriter

import facility_data

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(facility_data.BASEDIRPATH, "merged_files")

### Standard full file name for the C file.
FILENAME = "C_Infrastructure Other Funding 2019.xlsx"


def merge_C(filepath):
    """Create the C file, containing fields for additional funding.
    """
    funding_data = facility_data.get_additional_funding_data()

    # Reformat funding data
    facility_funding = []
    for facility, platform in facility_data.PLATFORM_LOOKUP.items():
        grants = []
        for rowdata in funding_data:
            if rowdata["facility"] == facility:
                grants.append(
                    (rowdata["additional_funding: Category of financier"],
                     rowdata["additional_funding: Name/type of financier"],
                     rowdata["additional_funding: Amount (kSEK)"])
                )
        facility_funding.append((facility, platform, grants))

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
    ws.set_column(0, 3, 40, long_text_format)
    ws.set_column(4, 4, 20, long_text_format)

    ws.write_row(0, 0, ("Facility", 
                        "Platform",
                        "Category of financier",
                        "Name/type of financier",
                        "Amount (kSEK)"))
    row = 1
    for facility, platform, grants in facility_funding:
        if len(grants) < 1:
            print("None for", facility, platform)
        elif len(grants) == 1:
            ws.write_row(row, 0, (facility, platform) + grants[0])
            row += 1
        else:
            ws.merge_range(row, 0, row + len(grants)-1, 0, facility)
            ws.merge_range(row, 1, row + len(grants)-1, 1, platform)
            for grant in grants:
                ws.write_row(row, 2, grant)
                row += 1

    wb.close()

    
if __name__ == "__main__":
    merge_C(os.path.join(DIRPATH, FILENAME))
