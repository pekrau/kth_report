"""Create the figure 5 in the 2019 report, showing the spread of affiliations
of the users of the SciLifeLab facilities.

"Spridning av tillhörighet för SciLifeLAb-faciliteternas användare"
"""

import os.path

import plotly.graph_objects as go
import openpyxl

import facility_data

wb = openpyxl.load_workbook(os.path.join(facility_data.BASEDIRPATH,
                                         "merged_files",
                                         "E_Infrastructure Users 2019.xlsx"))
ws = wb.active
headers = ["facility", "platform", "email",
           "pi_first_name", "pi_last_name", "pi_email",
           "affiliation", "affiliation_other"]
rows = list(ws)
records = [dict(zip(headers, [c.value for c in row])) for row in rows[1:]]
print(len(records), "records")
wb.close()

counts = {}
for record in records:
    facility = counts.setdefault(record["facility"], dict())
    # Fix error in input; two cases.
    if not record["affiliation"]:
        record["affiliation"] = "Industry"
    try:
        facility[record["affiliation"] or ""] += 1
    except KeyError:
        facility[record["affiliation"]] = 0

all_affiliations = set()
for facility, affiliations in counts.items():
    all_affiliations.update(affiliations)
    # print(facility),
    # print(sorted(affiliations.items()))
all_affiliations = sorted(all_affiliations)
print(all_affiliations)
affiliation_ypos = dict([(a, y) for y, a in enumerate(all_affiliations)])
print(affiliation_ypos)

colours = ["#1E3F32", "#01646B", "#4f9b74", "#80C41C", "#1b918d", "378CAF",
           "#468365", "#AECE53", "#87B0AB", "#AEC69C", "#819e90", "#B1B0B1",
           "#1E3F32", "#01646B", "#4f9b74", "#80C41C", "#1b918d"]

trace = {"x": [], "y": [], "mode": "markers", "type": "scatter"}
for xpos, facility in enum(counts):
    ypos = 1

# go = go.Figure()

# go.show()
