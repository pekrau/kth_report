"""
Read the facility data XLSX files and return as lists of dictionaries,
one dictionary for each row.
"""

import os.path

import openpyxl

BASEDIRPATH = os.path.expanduser("~/Nextcloud")
YEAR = "2019"
BASEFILENAME = f"orders_Facility_report_{YEAR}"
DIRPATH = os.path.join(BASEDIRPATH,
                       f"Årsrapport {YEAR}/Facility reports/aggregate_files")

REPORT_FILEPATH = os.path.join(DIRPATH, f"{BASEFILENAME}.xlsx")
def get_report_data(filepath=REPORT_FILEPATH):
    "Get the data reported in single-valued fields the OrderPortal form."
    return read_file(filepath)

FACILITY_HEAD_FILEPATH = os.path.join(
    DIRPATH, f"{BASEFILENAME}_facility_head.xlsx")
def get_facility_head_data(filepath=FACILITY_HEAD_FILEPATH):
    "Get the facility head data."
    return read_file(filepath)

FACILITY_DIRECTOR_FILEPATH = os.path.join(
    DIRPATH, f"{BASEFILENAME}_facility_director.xlsx")
def get_facility_director_data(filepath=FACILITY_DIRECTOR_FILEPATH):
    "Get the facility director data."
    return read_file(filepath)

ADDITIONAL_FUNDING_FILEPATH = os.path.join(
    DIRPATH, f"{BASEFILENAME}_additional_funding.xlsx")
def get_additional_funding_data(filepath=ADDITIONAL_FUNDING_FILEPATH):
    "Get the additional funding data."
    return read_file(filepath)

IMMATERIAL_PROPERTY_RIGHTS_FILEPATH = os.path.join(
    DIRPATH, f"{BASEFILENAME}_immaterial_property_rights.xlsx")
def get_immaterial_property_rights_data(filepath=IMMATERIAL_PROPERTY_RIGHTS_FILEPATH):
    "Get the immaterial_propery_rights data."
    return read_file(filepath)

def read_file(filepath):
    """Open the Excel file given by the path.
    Return the list of dictionaries.
    """
    wb = openpyxl.load_workbook(filename=filepath)
    return get_rows(wb)

def get_rows(wb):
    """Get the data for the active sheet as a list of dictionaries.
    The first row of the sheet defines the keys, and each subsequent
    row (the data) is stored in a separate dictionary. All
    dictionaries have the same keys.
    """
    ws = wb.active
    rows = list(ws.rows)
    headers = [cell.value for cell in rows[0] if cell.value]
    result = []
    for row in rows[1:]:
        data = dict(zip(headers, [cell.value for cell in row]))
        result.append(data)
    return result


if __name__ == "__main__":
    for f in [get_report_data,
              get_facility_head_data,
              get_facility_director_data,
              get_additional_funding_data,
              get_immaterial_property_rights_data]:
        data = f()
        print(len(data))
