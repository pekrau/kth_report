"""Facility Reports 2019.

Create the file 'E_Infrastructure Users 2019.xlsx'
"""

import os.path

import xlsxwriter

import facility_data

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(facility_data.BASEDIRPATH, "merged_files")

### Full file name for the E file.
FILENAME = "E_Infrastructure Users 2019.xlsx"


def merge_E(filepath):
    """Create the E file, containing all facility users.
    """
    users_data = facility_data.get_users_data()
    
if __name__ == "__main__":
    merge_E(os.path.join(DIRPATH, FILENAME))
