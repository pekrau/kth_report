"""SciLifeLab brand colors, in usable Python form.
- Various tints (saturation).
- Uses web-friendly hexcode notation.
- Different data structures are available for convenience.
- Also outputs a JSON file of the data.
"""

class Palette:
    """Return a color from the given color range.
    Allows cycling through the range indefinitely.
    """
    def __init__(self, color_range):
        self.color_range = tuple(color_range)
    def __getitem__(self, i):
        return self.color_range[i % len(self.color_range)]

# Base colors
LIME = "#A7C947"
TEAL = "#045C64"
AQUA = "#4C979F"
GRAPE = "#491F53"

# Grays
LIGHTGRAY = "#E5E5E5"
MEDIUMGRAY = "#A6A6A6"
DARKGRAY = "#3F3F3F"
GRAYS = (LIGHTGRAY, MEDIUMGRAY, DARKGRAY)

# Tints (saturation, in percent)
LIME_TINTS = {25: "#E9F2D1", 50: "#D3E4A3", 75: "#BDD775", 100: LIME}
TEAL_TINTS = {25: "#C0D6D8", 50: "#82AEB2", 75: "#43858B", 100: TEAL}
AQUA_TINTS = {25: "#D2E5E7", 50: "#A6CBCF", 75: "#79B1B7", 100: AQUA}
GRAPE_TINTS = {25: "#D2C7D4", 50: "#A48FA9", 75: "#77577E", 100: GRAPE}
              
BASE_COLOR_RANGE = (LIME, TEAL, AQUA, GRAPE, DARKGRAY)
MEDIUM_COLOR_RANGE = (LIME_TINTS[50], LIME_TINTS[75],
                      TEAL_TINTS[50], TEAL_TINTS[75],
                      AQUA_TINTS[50], AQUA_TINTS[75],
                      GRAPE_TINTS[50], GRAPE_TINTS[75],
                      MEDIUMGRAY, DARKGRAY)
ALL_COLOR_RANGE = tuple([LIME_TINTS[t] for t in sorted(LIME_TINTS)] +
                        [TEAL_TINTS[t] for t in sorted(TEAL_TINTS)] +
                        [AQUA_TINTS[t] for t in sorted(AQUA_TINTS)] +
                        [GRAPE_TINTS[t] for t in sorted(GRAPE_TINTS)] +
                        list(GRAYS))

base_color_palette = Palette(BASE_COLOR_RANGE)
medium_color_palette = Palette(MEDIUM_COLOR_RANGE)
all_color_palette = Palette(ALL_COLOR_RANGE)

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
    data = dict(
        title="SciLifeLab brand colors",
        colors=dict(lime=LIME, teal=TEAL, aqua=AQUA, grape=GRAPE),
        grays=dict(lightgray=LIGHTGRAY,
                   mediumgray=MEDIUMGRAY,
                   darkgray=DARKGRAY),
        tints=dict(lime=LIME_TINTS, teal=TEAL_TINTS,
                   aqua=AQUA_TINTS, grape=GRAPE_TINTS),
        base_colors=BASE_COLOR_RANGE,
        medium_colors=MEDIUM_COLOR_RANGE,
        all_colors=ALL_COLOR_RANGE,
        lookup=color_lookup)
    with open("scilifelab_brand_colors.json", "w") as outfile:
        json.dump(data, outfile, indent=2)
