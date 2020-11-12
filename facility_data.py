"""
Read the facility data XLSX files and return as lists of dictionaries.
"""

import os.path

import openpyxl

YEAR = "2019"
DIRPATH = os.path.expanduser(f"~/Nextcloud/Årsrapport {YEAR}/Facility reports/input_files")


def get_single_data(filename=f"A_Infrastructure Single Data Reported {YEAR}.xlsx"):
    "Return the data reported in single-valued fields the OrderPortal form."
    wb = openpyxl.load_workbook(filename=os.path.join(DIRPATH, filename))
    return get_rows(wb)

def get_rows(wb):
    """Get the data for the active sheet as a list of dictionaries.
    The first row of the sheet defines the keys, and each subsequent
    row is stored in a separate dictionary. All dictionaries have the
    same keys.
    """
    ws = wb.active
    rows = list(ws.rows)
    headers = [cell.value for cell in rows[0] if cell.value]
    for cell in rows[0]:
        print(cell, cell.column, cell.row, cell.coordinate, cell.parent)
        if cell.coordinate == "D1":
            print(dir(cell))
    result = []
    for row in rows[1:]:
        data = dict(zip(headers, [cell.value for cell in row]))
        result.append(data)
    return result


if __name__ == "__main__":
    data = get_fd_hf_data()
