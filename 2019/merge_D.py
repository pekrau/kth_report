"""Facility Reports 2019.

Create the file 'D_Infrastructure Immaterial Property Rights 2019.xlsx'
"""

import os.path

import xlsxwriter

import facility_data

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(facility_data.BASEDIRPATH, "merged_files")

### Full file name for the D file.
FILENAME = "D_Infrastructure Immaterial Property Rights 2019.xlsx"


def merge_D(filepath):
    """Create the D file, containing fields for immaterial property rights.
    """
    ip_data = facility_data.get_ip_rights_data()

    # Reformat IP data
    facility_ip = []
    for facility, platform in facility_data.PLATFORM_LOOKUP.items():
        patents = []
        for record in ip_data:
            if record["facility"] == facility:
                patents.append(
                    (record["immaterial_property_rights: Patent title"],
                     record["immaterial_property_rights: Patent application number"],
                     # 2018! Must have forgotten to update the form...
                     record["immaterial_property_rights: Filed or granted during 2018?"],
                     record["immaterial_property_rights: Registered designs"],
                     record["immaterial_property_rights: Registered trademarks"])
                )
        facility_ip.append((facility, platform, patents))

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
    ws.set_column(0, 2, 40, long_text_format)
    ws.set_column(3, 6, 20, long_text_format)

    ws.write_row(0, 0, ("Facility", 
                        "Platform",
                        "Patent title",
                        "Patent application number",
                        "Filed or granted during 2018?",
                        "Registered designs",
                        "Registered trademarks"))
    row = 1
    for facility, platform, patents in facility_ip:
        if len(patents) < 1:
            print("None for", facility)
        elif len(patents) == 1:
            ws.write_row(row, 0, (facility, platform) + patents[0])
            row += 1
        else:
            ws.merge_range(row, 0, row + len(patents)-1, 0, facility)
            ws.merge_range(row, 1, row + len(patents)-1, 1, platform)
            for patent in patents:
                ws.write_row(row, 2, patent)
                row += 1

    wb.close()

    
if __name__ == "__main__":
    merge_D(os.path.join(DIRPATH, FILENAME))
