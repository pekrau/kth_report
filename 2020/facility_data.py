"""Facility Reports 2020.

Read the aggregated facility data XLSX files and return as lists of
dictionaries, one dictionary for each row.

The files were created in and downloaded from the Reporting Portal 
Reporting Portal https://reporting.scilifelab.se/
"""

import os.path

import openpyxl

BASEDIRPATH = os.path.expanduser("~/Nextcloud/Årsrapport 2020/Facility reports")
BASEFILENAME = "orders_Facility_report_2020"

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(BASEDIRPATH, "aggregate_files")


# Lookup from facility name to platform name.
PLATFORM_LOOKUP = {
    "Advanced Light Microscopy" : "Cellular and Molecular Imaging",
    "Ancient DNA" : "Genomics",
    "Autoimmunity and Serology Profiling" : "Proteomics and Metabolomics",
    "BioImage Informatics" : "Cellular and Molecular Imaging",
    "Cell Profiling" : "Cellular and Molecular Imaging",
    "Chemical Biology Consortium Sweden" : "Chemical Biology and Genome Engineering",
    "Chemical Proteomics and Proteogenomics (MBB)" : "Proteomics and Metabolomics",
    "Chemical Proteomics and Proteogenomics (OnkPat)" : "Proteomics and Metabolomics",
    "Clinical Genomics Gothenburg" : "Diagnostics Development",
    "Clinical Genomics Linköping" : "Diagnostics Development",
    "Clinical Genomics Lund" : "Diagnostics Development",
    "Clinical Genomics Stockholm" : "Diagnostics Development",
    "Clinical Genomics Umeå" : "Diagnostics Development",
    "Clinical Genomics Uppsala" : "Diagnostics Development",
    "Clinical Genomics Örebro" : "Diagnostics Development",
    "Compute and Storage" : "Bioinformatics",
    "Cryo-EM" : "Cellular and Molecular Imaging",
    "Drug Discovery and Development" : "Drug Discovery and Development",
    "Eukaryotic Single Cell Genomics" : "Genomics",
    "Genome Engineering Zebrafish" : "Chemical Biology and Genome Engineering",
    "High Throughput Genome Engineering" : "Chemical Biology and Genome Engineering",
    "Long-term Support (WABI)" : "Bioinformatics",
    "Mass Cytometry (KI)" : "Proteomics and Metabolomics",
    "Mass Cytometry (LiU)" : "Proteomics and Metabolomics",
    "Microbial Single Cell Genomics" : "Genomics",
    "National Genomics Infrastructure" : "Genomics",
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

def read_file(filepath):
    """Open the Excel file given by the path.
    Return the list of dictionaries.
    """
    wb = openpyxl.load_workbook(filename=filepath)
    return get_rows(wb)

def get_rows(wb):
    """Get the data for the active sheet as a list of
    dictionaries. The first row of the sheet defines the keys, and
    each subsequent row (the data) is stored in a separate
    dictionary. All dictionaries have the same keys.
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
              get_ip_rights_data]:
        data = f()
        print(len(data))
