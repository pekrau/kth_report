"""Facility Reports 2019.

Read the aggregated facility data XLSX files and return as lists of
dictionaries, one dictionary for each row.

The files were created in and downloaded from the Reporting Portal 
Reporting Portal https://reporting.scilifelab.se/
"""

import glob
import os.path

import openpyxl

BASEDIRPATH = os.path.expanduser("~/Nextcloud/Årsrapport 2019/Facility reports")
BASEFILENAME = "orders_Facility_report_2019"

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(BASEDIRPATH, "aggregate_files")

### Path to directory containing the downloaded volume data files.
VOLDIRPATH = os.path.join(BASEDIRPATH, "volume_data_files")


# Lookup from facility name to platform name.
PLATFORM_LOOKUP = {
    "Advanced Light Microscopy (ALM)" : "Cellular and Molecular Imaging",
    "Ancient DNA" : "Genomics",
    "Autoimmunity Profiling" : "Proteomics and Metabolomics",
    "BioImage Informatics" : "Cellular and Molecular Imaging",
    "Cell Profiling" : "Cellular and Molecular Imaging",
    "Chemical Biology Consortium Sweden" : "Chemical Biology and Genome Engineering",
    "Chemical Proteomics and Proteogenomics (MBB)" : "Proteomics and Metabolomics",
    "Chemical Proteomics and Proteogenomics (OnkPat)" : "Proteomics and Metabolomics",
    "Clinical Genomics Göteborg" : "Diagnostics Development",
    "Clinical Genomics Lund" : "Diagnostics Development",
    "Clinical Genomics Stockholm" : "Diagnostics Development",
    "Clinical Genomics Uppsala" : "Diagnostics Development",
    "Compute and Storage" : "Bioinformatics",
    "Cryo-EM (SU)" : "Cellular and Molecular Imaging",
    "Cryo-EM (UmU)" : "Cellular and Molecular Imaging",
    "Drug Discovery and Development" : "Drug Discovery and Development",
    # *** NOTE *** This is renamed to 'In Situ Sequencing', see below.
    # "Eukaryotic Single Cell Genomics" : "Genomics",
    "Genome Engineering Zebrafish" : "Chemical Biology and Genome Engineering",
    "High Throughput Genome Engineering" : "Chemical Biology and Genome Engineering",
    "In Situ Sequencing" : "Genomics",
    "Long-term Support (WABI)" : "Bioinformatics",
    "Mass Cytometry (KI)" : "Proteomics and Metabolomics",
    "Mass Cytometry (LiU)" : "Proteomics and Metabolomics",
    "Microbial Single Cell Genomics" : "Genomics",
    "NGI Stockholm" : "Genomics",
    "NGI Uppsala SNP&SEQ" : "Genomics",
    "NGI Uppsala UGC" : "Genomics",
    "PLA and Single Cell Proteomics" : "Proteomics and Metabolomics",
    "Plasma Profiling" : "Proteomics and Metabolomics",
    "Protein Science Facility" : "Cellular and Molecular Imaging",
    "Support and Infrastructure" : "Bioinformatics",
    "Swedish Metabolomics Centre" : "Proteomics and Metabolomics",
    "Swedish NMR Centre" : "Cellular and Molecular Imaging",
    "Systems Biology" : "Bioinformatics"
}

REPORT_FILEPATH = os.path.join(DIRPATH, f"{BASEFILENAME}.xlsx")

def get_report_data(filepath=REPORT_FILEPATH):
    """Get the data reported in single-valued fields of the
    OrderPortal form.
    """
    return read_file(filepath)

FACILITY_HEAD_FILEPATH = os.path.join(
    DIRPATH, f"{BASEFILENAME}_facility_head.xlsx")

def get_facility_head_data(filepath=FACILITY_HEAD_FILEPATH):
    """Get the facility head data.
    """
    return read_file(filepath)

FACILITY_DIRECTOR_FILEPATH = os.path.join(
    DIRPATH, f"{BASEFILENAME}_facility_director.xlsx")

def get_facility_director_data(filepath=FACILITY_DIRECTOR_FILEPATH):
    """Get the facility director data.
    """
    return read_file(filepath)

ADDITIONAL_FUNDING_FILEPATH = os.path.join(
    DIRPATH, f"{BASEFILENAME}_additional_funding.xlsx")

def get_additional_funding_data(filepath=ADDITIONAL_FUNDING_FILEPATH):
    """Get the additional funding data.
    """
    return read_file(filepath)

IP_RIGHTS_FILEPATH = os.path.join(
    DIRPATH, f"{BASEFILENAME}_immaterial_property_rights.xlsx")

def get_ip_rights_data(filepath=IP_RIGHTS_FILEPATH):
    """Get the immaterial_propery_rights data."""
    return read_file(filepath)

def get_users_data(dirpath=VOLDIRPATH):
    """Get all users for each facility.
    Returns dictionary with key=facility name, value=list of row data
    in dictionaries.
    """
    count = 0
    for filepath in sorted(glob.glob(f"{VOLDIRPATH}/*.xls[mx]")):
        try:
            data = read_file(filepath,
                             sheetname="A. Users", 
                             skip_rows_until="Name of reporting unit")
        except KeyError as error:
            print(os.path.basename(filepath), error)
            data = []
    result = {}
    # XXX
    return result

def read_file(filepath):
    """Open the Excel file given by the path and read the first sheet.
    Return a list of dictionaries, one for each row.
    """
    wb = openpyxl.load_workbook(filename=filepath)
    rows = list(wb.active)
    header = [c.value for c in rows[0]]
    result = []
    for row in rows[1:]:
        record = dict(list(zip(header, [c.value for c in row])))
        # *** NOTE *** Special case; correct the name of one facility.
        if record.get("facility") == "Eukaryotic Single Cell Genomics":
            record["facility"] = "In Situ Sequencing"
        result.append(record)
    return result

# def get_rows(ws, skip_rows_until=None):
#     """Get the data from the given sheet as a list of dictionaries.
#     If given, the rows whose first cell does not contain
#     'skip_rows_until' are discarded before using the rest of the rows.
#     The first row defines the keys, and each subsequent row (the data)
#     is stored in a separate dictionary. All dictionaries have the same
#     keys.
#     NOTE: Special case; correct the name of one facility.
#     """
#     # Awful kludge to avoid weird behaviour when reading
#     # all rows from a sheet in one file after another...
#     rows = []
#     while True:
#         for row in ws.iter_rows(min_row=len(rows)+1, max_row=len(rows)+100):
#             rows.append([cell.value for cell in row])
#         if not rows[-1] or rows[-1][0] is None: break
#     while rows[-1][0] is None:
#         rows.pop()
#     if skip_rows_until is None:
#         first = 0
#     else:
#         for first, row in enumerate(rows):
#             if row[0] and skip_rows_until in row[0]: break
#     headers = [cell for cell in rows[first] if cell]
#     result = []
#     for row in rows[first+1:]:
#         data = dict(zip(headers, row))
#         # *** NOTE *** Special case; correct the name of one facility.
#         if data.get("facility") == "Eukaryotic Single Cell Genomics":
#             data["facility"] = "In Situ Sequencing"
#         result.append(data)
#     return result

def read_volume_file(filepath, sheetname):
    """Open the Excel Volume data file given by the path and read the
    sheet with the given name, or the first sheet. Return a tuple
    (facility, list-of-records), where a record is a dictionary.
    """
    wb = openpyxl.load_workbook(filename=filepath)
    wb.active = wb.get_sheet_by_name(sheetname)
    result = get_rows(wb.active, skip_rows_until=skip_rows_until)
    wb.close()
    return result

def get_volume_rows(ws, skip_rows_until=None):
    """Get the data from the given sheet as a list of dictionaries.
    If given, the rows whose first cell does not contain
    'skip_rows_until' are discarded before using the rest of the rows.
    The first row defines the keys, and each subsequent row (the data)
    is stored in a separate dictionary. All dictionaries have the same
    keys.
    NOTE: Special case; correct the name of one facility.
    """
    # Awful kludge to avoid weird behaviour when reading
    # all rows from a sheet in one file after another...
    rows = []
    while True:
        for row in ws.iter_rows(min_row=len(rows)+1, max_row=len(rows)+100):
            rows.append([cell.value for cell in row])
        if not rows[-1] or rows[-1][0] is None: break
    while rows[-1][0] is None:
        rows.pop()
    if skip_rows_until is None:
        first = 0
    else:
        for first, row in enumerate(rows):
            if row[0] and skip_rows_until in row[0]: break
    headers = [cell for cell in rows[first] if cell]
    result = []
    for row in rows[first+1:]:
        data = dict(zip(headers, row))
        # *** NOTE *** Special case; correct the name of one facility.
        if data.get("facility") == "Eukaryotic Single Cell Genomics":
            data["facility"] = "In Situ Sequencing"
        result.append(data)
    return result


if __name__ == "__main__":
    for f in [get_report_data,
              get_facility_head_data,
              get_facility_director_data,
              get_additional_funding_data,
              get_ip_rights_data]:
        data = f()
        print(len(data))
