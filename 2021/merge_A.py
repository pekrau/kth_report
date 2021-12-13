"""Infrastructure Units Reports 2021.

Create the file 'A_Infrastructure Single Data Reported 2021.xlsx'

This code has added a few fields for Covid-19 related information
compared to 2019. Otherwise it is identical.
"""

import os.path

import xlsxwriter

import facility_data

### Path to directory containing the downloaded aggregate files.
DIRPATH = os.path.join(facility_data.BASEDIRPATH, "merged_files")

### Standard full file name for the A file.
FILENAME = "A_Infrastructure Single Data Reported 2021.xlsx"

# List of single-valued field identifiers to collect from the aggregate file.
# Cut-and-paste from file 'Data files for KTH and Infra Reports.xlsx'.
FIELD_IDENTIFIERS = """personnel_count
personnel_count_male
personnel_count_phd
personnel_count_phd_male
fte
fte_scilifelab
eln_usage
resource_academic_national
resource_academic_international
resource_internal
resource_industry
resource_healthcare
resource_other
total_user_fees
user_fee_models
user_fees
user_fees_academic_sweden
user_fees_academic_international
user_fees_industry
user_fees_healthcare
user_fees_other
cost_reagents
cost_instrument
cost_salaries
cost_rents
cost_other
number_projects
number_projects_covid19
fte_covid19
impact_covid19
user_feedback
innovation_utilization
technology_development
scientific_achievements""".split("\n")

# The column headers must match the order of the field identifiers above,
# except for the two first.
COLUMN_HEADERS = ["Facility",
                  "Platform",
                  "Personnel count",
                  "Personnel count male",
                  "Personnel count Phd",
                  "Personnel count Phd male",
                  "FTE",
                  "FTE Scilifelab",
                  "ELN usage",
                  "Resource academic national",
                  "Resource academic international",
                  "Resource internal",
                  "Resource industry",
                  "Resource healthcare",
                  "Resource other",
                  "Total user fees",
                  "User fee models",
                  "User fees",
                  "User fees academic Sweden",
                  "User fees academic international",
                  "User fees industry",
                  "User fees healthcare",
                  "User fees other",
                  "Cost reagents",
                  "Cost instrument",
                  "Cost salaries",
                  "Cost rents",
                  "Cost other",
                  "# Projects",
                  "# Covid-19 projects",
                  "# Covid-19 FTE resources",
                  "Impact Covid-19",
                  "User feedback",
                  "Innovation utilization",
                  "Technology development",
                  "Scientific achievements"]


def create_A(filepath):
    """Create the A file, containing single-valued fields from the
    infrastructure facility reports.
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
    ws.set_column(0, 1, 40, normal_text_format)
    ws.set_column(2, 29, 20, normal_text_format)
    ws.set_column(33, 35, 100, long_text_format)

    ws.write_row(0, 0, COLUMN_HEADERS)
    for rownum, report in enumerate(facility_data.get_report_data()):
        facility = report["facility"]
        platform = facility_data.PLATFORM_LOOKUP[facility]
        rowdata = [facility, platform]
        rowdata.extend([report.get(fid, "") for fid in FIELD_IDENTIFIERS])
        ws.write_row(rownum+1, 0, rowdata)
    wb.close()

    
if __name__ == "__main__":
    create_A(os.path.join(DIRPATH, FILENAME))
