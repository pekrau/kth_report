"""SciLifeLab brand colors, in usable Python form.
- Also outputs a JSON file of the data.
- Various tints (saturation).
- Uses web-friendly hexcode notation.
- Different datastructures are available for ease-of-use.
"""

# Base colors
LIME = "#A7C947"
TEAL = "#045C64"
AQUA = "#4C979F"
GRAPE = "#491F53"

# Grays
LIGHTGRAY = "#E5E5E5"
MEDIUMGRAY = "#A6A6A6"
DARKGRAY = "#3F3F3F"

# Tints (saturation, in percent)
LIME_TINTS = {25: "#E9F2D1", 50: "#D3E4A3", 75: "#BDD775", 100: LIME}
TEAL_TINTS = {25: "#C0D6D8", 50: "#82AEB2", 75: "#43858B", 100: TEAL}
AQUA_TINTS = {25: "#D2E5E7", 50: "#A6CBCF", 75: "#79B1B7", 100: AQUA}
GRAPE_TINTS = {25: "#D2C7D4", 50: "#A48FA9", 75: "#77577E", 100: GRAPE}
              
SIMPLE_COLORS = (LIME, TEAL, AQUA, GRAPE, DARKGRAY)
               

color_lookup = dict(lime=LIME,
                    teal=TEAL,
                    aqua=AQUA,
                    grape=GRAPE,
                    lightgray=LIGHTGRAY,
                    mediumgray=MEDIUMGRAY,
                    darkgray=DARKGRAY)

for saturation, color in LIME_TINTS.items():
    color_lookup[f"lime{saturation}"] = color
for saturation, color in TEAL_TINTS.items():
    color_lookup[f"teal{saturation}"] = color
for saturation, color in AQUA_TINTS.items():
    color_lookup[f"aqua{saturation}"] = color
for saturation, color in GRAPE_TINTS.items():
    color_lookup[f"grape{saturation}"] = color


if __name__ == "__main__":
    import json
    data = dict(scilifelab_brand_colors=dict(
        base_colors=dict(lime=LIME, teal=TEAL, aqua=AQUA, grape=GRAPE),
        grays=dict(lightgray=LIGHTGRAY,
                   mediumgray=MEDIUMGRAY,
                   darkgray=DARKGRAY),
        simple_colors=SIMPLE_COLORS,
        lookup=color_lookup))
    with open("scilifelab_brand_colors.json", "w") as outfile:
        json.dump(data, outfile, indent=2)
