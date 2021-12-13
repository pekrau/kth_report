"""Create the figure 5 (the number in the 2019 report), showing the spread 
of affiliations of the users of the SciLifeLab infrastructure units.

"Spridning av tillhörighet för SciLifeLab-enheternas användare"
"""

import csv
import math
import os.path

import openpyxl
import plotly.graph_objects as go

import facility_data
import scilifelab_brand_colors


INPUTFILENAME = os.path.join(facility_data.BASEDIRPATH,
                             "merged_files",
                             "E_Infrastructure Users 2021.xlsx")
OUTPUTFILENAME = os.path.join(facility_data.BASEDIRPATH,
                              "figures",
                              "fig_5_2021")


# Browser
### IMAGE = False

# PNG image file
IMAGE = True

if IMAGE:
    SCALE = 5.0
    IMAGE_WIDTH = 7685          # Aspect ratio 1:1.8
    IMAGE_HEIGHT = 4250
    TITLE_Y = 0.99
else:
    SCALE = 1.0
    BROWSER_WIDTH = 1537        # Aspect ratio 1:1.8
    BROWSER_HEIGHT = 850
    TITLE_Y = 0.95


def get_marker_size(number):
    """Same scaling as for year 2019. Produces more overlap between circles.
    But this was considered OK, since it does reflect the reality.
    """
    return SCALE * (5 * math.sqrt(number) + 5)


# SciLifeLab brand colors, 50% and 75% tint (saturation)
# The palette object allows cycling through the range of colors indefinitely.
colors = scilifelab_brand_colors.medium_color_palette

# Colors from 2019 JavaScript code.
# Cycle through the scale a few times...
# colors = ["#1E3F32", "#01646B", "#4f9b74", "#80C41C", "#1b918d", "#378CAF",
#           "#468365", "#AECE53", "#87B0AB", "#AEC69C", "#819e90", "#B1B0B1",
#           "#1E3F32", "#01646B", "#4f9b74", "#80C41C", "#1b918d", "#378CAF",
#           "#468365", "#AECE53", "#87B0AB", "#AEC69C", "#819e90", "#B1B0B1",
#           "#1E3F32", "#01646B", "#4f9b74", "#80C41C", "#1b918d", "#378CAF",
#           "#468365", "#AECE53", "#87B0AB", "#AEC69C", "#819e90", "#B1B0B1"]

# Alphabetical order. Forget about trying to keep approximately the
# same order as last year; too many changes.
FACILITIES = sorted(facility_data.PLATFORM_LOOKUP.keys())

# Affiliation in English this year, apparently.
AFFILIATIONS = [
    "Chalmers University of Technology",
    "Karolinska Institutet",
    "KTH Royal Institute of Technology",
    "Linköping University",
    "Lund University",
    "Stockholm University",
    "Swedish University of Agricultural Sciences",
    "Umeå University",
    "University of Gothenburg",
    "Uppsala University",
    "Örebro University",
    "Other Swedish University",
    "International University",
    "Healthcare",
    "Industry",
    "Naturhistoriska Riksmuséet",
    "Other Swedish organization",
    "Other international organization",
]

wb = openpyxl.load_workbook(INPUTFILENAME)
ws = wb.active
rows = list(ws)
# Skip first row; header
rows = rows[1:]
records = []
# Read off column positions manually from file.
FACILITY_COL = 0
PI_COL = 5
AFFILIATION_COL = 6
for row in rows:
    values = [c.value for c in row]
    affiliation = values[AFFILIATION_COL]
    # No affiliation specified: skip (or correct in the input file).
    if not affiliation:
        print("No affiliation for", values[FACILITY_COL], values[PI_COL])
        continue
    # A trailing blank in the input XLSX pull-down menu; remove it.
    affiliation = affiliation.strip()
    # Some value are lower-case first character?!
    affiliation = affiliation[0].upper() + affiliation[1:]
    records.append(dict(facility=values[FACILITY_COL],
                        pi=values[PI_COL],
                        affiliation=affiliation))
print(len(records), "records in file")
wb.close()

counts = {}
for record in records:
    facility = counts.setdefault(record["facility"], dict())
    try:
        facility[record["affiliation"]] += 1
    except KeyError:
        facility[record["affiliation"]] = 1

# Sanity check: The hardwired facilities matches the input.
all_facilities = counts.keys()
if set(FACILITIES) != set(all_facilities):
    print(sorted(set(FACILITIES).difference(all_facilities)),
          "\n\n",
          sorted(set(all_facilities).difference(FACILITIES)))
    raise ValueError("Hardwired facilities do not match input")

# Sanity check: The hardwired affiliations matches the input.
all_affiliations = set()
for facility, affiliations in counts.items():
    all_affiliations.update(affiliations)
if set(AFFILIATIONS) != all_affiliations:
    print(set(AFFILIATIONS).difference(all_affiliations),
          "\n\n",
          set(all_affiliations).difference(AFFILIATIONS))
    raise ValueError("Hardwired affiliations do not match input")
# print(sorted(all_affiliations))
affiliation_pos = dict([(a, y) for y, a in enumerate(AFFILIATIONS)])


data = []
for a, affiliation in enumerate(AFFILIATIONS):
    x = []
    y = []
    marker_size = []
    marker_text = []
    for f, facility in enumerate(FACILITIES):
        try:
            number = counts[facility][affiliation]
        except KeyError:
            pass
        else:
            x.append(f+1)
            y.append(a+1)
            marker_size.append(get_marker_size(number))
            marker_text.append(f"{affiliation} / {facility}")
    trace = {"mode": "markers",
             "type": "scatter",
             "x": x,
             "y": y,
             "marker": {"size": marker_size,
                        "color": colors[a]},
             "text": marker_text,
             "name": affiliation,
             "hoverinfo": "text",
    }
    data.append(trace)

fig = go.Figure(
    data=data,
    layout={
        "plot_bgcolor": "#fff",
        "showlegend": False,
        "xaxis": {
            "title": {"text": "Infrastrukturenheter",
                      "font": {"family": "Arial", "size": SCALE * 18}},
            "range": [0, len(FACILITIES) + 1],
            "gridcolor": "#eeeeee",
            "tickvals": list(range(1, len(FACILITIES) + 1)),
            "ticktext": FACILITIES,
            "tickfont": {"family": "Arial", "size": SCALE * 16},
            "tickangle": -40,
        },
        "yaxis": {
            "title": {"text": "Användartillhörighet",
                      "font": {"family": "Arial", "size": SCALE * 18}},
            "gridcolor": "#eeeeee",
            "tickvals": list(range(1, len(AFFILIATIONS) + 1)),
            "ticktext": AFFILIATIONS,
            "tickfont": {"family": "Arial", "size": SCALE * 16},
            "tickangle": -40,
            "zerolinecolor": "#6E6E6E",
        },
    })

with open(OUTPUTFILENAME + ".csv", "w") as outfile:
    writer = csv.writer(outfile)
    writer.writerow(["Infrastructure Unit"] + AFFILIATIONS)
    for facility in FACILITIES:
        row = [facility]
        for affiliation in AFFILIATIONS:
            try:
                row.append(counts[facility][affiliation])
            except KeyError:
                row.append(0)
        writer.writerow(row)


if IMAGE:
    fig.write_image(OUTPUTFILENAME + ".png",
                    width=IMAGE_WIDTH,
                    height=IMAGE_HEIGHT)
else:
    fig.show()

