"""Create the figure 5 in the 2019 report, showing the spread of affiliations
of the users of the SciLifeLab facilities.

"Spridning av tillhörighet för SciLifeLAb-faciliteternas användare"
"""

import math
import os.path

import openpyxl
import plotly.graph_objects as go

import facility_data

VERSION = "2"

INPUTFILENAME = os.path.join(facility_data.BASEDIRPATH,
                             "figures",
                             "Analyses Users 2020 for fig 5.xlsx")
OUTPUTFILENAME = os.path.join(facility_data.BASEDIRPATH,
                              "figures",
                              "fig_5_2020.png")


# Browser
# IMAGE = False

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
# Cycle through the scale a few times...
colors = ["#D3E4A3", "#BDD775",  # Lime
          "#82AEB2", "#43858B",  # Teal
          "#A6CBCF", "#79B1B7",  # Aqua
          "#A48FA9", "#77577E",  # Grape
          "#A6A6A6", "#3F3F3F",  # Gray
          "#D3E4A3", "#BDD775",   # Lime
          "#82AEB2", "#43858B",  # Teal
          "#A6CBCF", "#79B1B7",  # Aqua
          "#A48FA9", "#77577E",  # Grape
          "#A6A6A6", "#3F3F3F",  # Gray
]

# Colours from 2019 JavaScript code.
# Cycle through the scale a few times...
# colors = ["#1E3F32", "#01646B", "#4f9b74", "#80C41C", "#1b918d", "#378CAF",
#           "#468365", "#AECE53", "#87B0AB", "#AEC69C", "#819e90", "#B1B0B1",
#           "#1E3F32", "#01646B", "#4f9b74", "#80C41C", "#1b918d", "#378CAF",
#           "#468365", "#AECE53", "#87B0AB", "#AEC69C", "#819e90", "#B1B0B1",
#           "#1E3F32", "#01646B", "#4f9b74", "#80C41C", "#1b918d", "#378CAF",
#           "#468365", "#AECE53", "#87B0AB", "#AEC69C", "#819e90", "#B1B0B1"]

# Set explicitly to get approx same order as last year.
FACILITIES = [
    'Long-term Support (WABI)',
    'Support and Infrastructure',
    'Systems Biology',
    'Advanced Light Microscopy (ALM)',
    'BioImage Informatics',
    'Cell Profiling', 
    'Cryo-EM',
    'Swedish NMR Centre',
    'Chemical Biology Consortium Sweden (KI)',
    'Chemical Biology Consortium Sweden (UmU)',
    'Genome Engineering Zebrafish',
    'High Throughput Genome Engineering',
    'In Situ Sequencing',
    'Clinical Genomics Gothenburg',
    'Clinical Genomics Linköping',
    'Clinical Genomics Lund',
    'Clinical Genomics Stockholm', 
    'Clinical Genomics Umeå', 
    'Clinical Genomics Uppsala',
    'Clinical Genomics Örebro',
    'Drug Discovery and Development', 
    'Ancient DNA',
    'National Genomics Infrastructure',
    'Autoimmunity Profiling',
    'Chemical Proteomics and Proteogenomics (MBB)',
    'Chemical Proteomics and Proteogenomics (OncPat)',
    'PLA and Single Cell Proteomics',
    'Plasma Profiling',
    'Swedish Metabolomics Centre',
    'Eukaryotic Single Cell Genomics',
    'Mass Cytometry (KI)',
    'Mass Cytometry (LiU)',
    'Microbial Single Cell Genomics',
]
# FACILITIES = [
#     'Advanced Light Microscopy (ALM)',
#     'Ancient DNA',
#     'Autoimmunity Profiling',
#     'BioImage Informatics',
#     'Cell Profiling', 
#     'Chemical Biology Consortium Sweden (KI)',
#     'Chemical Biology Consortium Sweden (UmU)',
#     'Chemical Proteomics and Proteogenomics (MBB)',
#     'Chemical Proteomics and Proteogenomics (OncPat)',
#     'Clinical Genomics Gothenburg',
#     'Clinical Genomics Linköping',
#     'Clinical Genomics Lund',
#     'Clinical Genomics Stockholm', 
#     'Clinical Genomics Umeå', 
#     'Clinical Genomics Uppsala',
#     'Clinical Genomics Örebro',
#     'Cryo-EM',
#     'Drug Discovery and Development', 
#     'Eukaryotic Single Cell Genomics',
#     'Genome Engineering Zebrafish',
#     'High Throughput Genome Engineering',
#     'In Situ Sequencing',
#     'Long-term Support (WABI)',
#     'Mass Cytometry (KI)',
#     'Mass Cytometry (LiU)',
#     'Microbial Single Cell Genomics',
#     'National Genomics Infrastructure',
#     'PLA and Single Cell Proteomics',
#     'Plasma Profiling',
#     'Support and Infrastructure',
#     'Swedish Metabolomics Centre',
#     'Swedish NMR Centre',
#     'Systems Biology',
# ]

AFFILIATIONS = [
    'Chalmers',
    'KTH', 
    'Karolinska Institutet',
    'Linköpings universitet',
    'Lunds universitet',
    'Naturhistoriska Riksmuséet',
    'Stockholms universitet',
    'Sveriges lantbruksuniversitet',
    'Umeå universitet',
    'Göteborgs universitet',
    'Uppsala universitet', 
    'Andra svenska lärosäten', 
    'Andra svenska organisationer',
    'Internationella universitet',
    'Andra internationella organisationer',
    'Hälso- och sjukvård', 
    'Industri',
]
# AFFILIATIONS = [
#     'Andra internationella organisationer',
#     'Andra svenska lärosäten', 
#     'Andra svenska organisationer',
#     'Chalmers',
#     'Göteborgs universitet',
#     'Hälso- och sjukvård', 
#     'Industri',
#     'Internationella universitet',
#     'KTH', 
#     'Karolinska Institutet',
#     'Linköpings universitet',
#     'Lunds universitet',
#     'Naturhistoriska Riksmuséet',
#     'Stockholms universitet',
#     'Svenska lantbruksuniversitetet',
#     'Umeå universitet',
#     'Uppsala universitet', 
# ]

wb = openpyxl.load_workbook(INPUTFILENAME)
ws = wb.active
rows = list(ws)
# Skip first row; header
rows = rows[1:]
headers = ["facility", "platform",
           "pi_first_name", "pi_last_name", "pi_email",
           "affiliation", "affiliation_other"]
records = [dict(zip(headers, [c.value for c in row])) for row in rows]
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
    raise ValueError("Hardwired facilities do not match input:",
                     set(FACILITIES).difference(all_facilities),
                     set(all_facilities).difference(FACILITIES))
# print(sorted(all_facilities))
# Sanity check: The hardwired affiliations matches the input.
all_affiliations = set()
for facility, affiliations in counts.items():
    all_affiliations.update(affiliations)
if set(AFFILIATIONS) != all_affiliations:
    raise ValueError("Hardwired affiliations do not match input:",
                     set(AFFILIATIONS).difference(all_affiliations),
                     set(all_affiliations).difference(AFFILIATIONS))
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
            "title": {"text": "Faciliteter",
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

if IMAGE:
    fig.write_image(OUTPUTFILENAME,
                    width=IMAGE_WIDTH,
                    height=IMAGE_HEIGHT)
else:
    fig.show()

