"""
LOB (Line of Business) Registry for NL-43 (Rural & Social Obligations).

LOBs appear as row labels in the PDF table (col 1), same orientation as some other forms.
NL-43 has 14 LOB rows + 1 total row.
"""

# Canonical LOB keys — ordered to match PDF row order.
LOB_ORDER = [
    "fire",
    "marine_cargo",
    "marine_hull",
    "motor_od",
    "motor_tp",
    "health",
    "personal_accident",
    "travel_insurance",
    "wc_el",
    "public_product_liability",
    "engineering",
    "aviation",
    "crop_insurance",
    "other_segments",
    "total_miscellaneous",
    "grand_total",              # "Total" row at the bottom
]

LOB_DISPLAY_NAMES = {
    "fire":                     "Fire",
    "marine_cargo":             "Marine Cargo",
    "marine_hull":              "Marine Other than Cargo",
    "motor_od":                 "Motor OD",
    "motor_tp":                 "Motor TP",
    "health":                   "Health",
    "personal_accident":        "Personal Accident",
    "travel_insurance":         "Travel",
    "wc_el":                    "WC / Employer's Liability",
    "public_product_liability": "Public / Product Liability",
    "engineering":              "Engineering",
    "aviation":                 "Aviation",
    "crop_insurance":           "Crop Insurance",
    "other_segments":           "Other Segments",
    "total_miscellaneous":      "Miscellaneous",
    "grand_total":              "Total",
}

LOB_ALIASES = {
    "fire":                                                     "fire",
    "marine cargo":                                             "marine_cargo",
    "marine other than cargo":                                  "marine_hull",
    "marine (other than cargo)":                                "marine_hull",
    "motor od":                                                 "motor_od",
    "motor ownerdamage":                                        "motor_od",
    "motor tp":                                                 "motor_tp",
    "motor thirdparty":                                         "motor_tp",
    "health":                                                   "health",
    "personal accident":                                        "personal_accident",
    "travel":                                                   "travel_insurance",
    "travel insurance":                                         "travel_insurance",
    "workmen's compensation/ employer's liability":             "wc_el",
    "workmen\u2019s compensation/ employer\u2019s liability":   "wc_el",
    "workmen's compensation/\nemployer's liability":            "wc_el",
    "workmen's compensation/employer's liability":              "wc_el",
    "workmen's compensation":                                   "wc_el",
    "wc/el":                                                    "wc_el",
    "public/ product liability":                                "public_product_liability",
    "public/product liability":                                 "public_product_liability",
    "engineering":                                              "engineering",
    "aviation":                                                 "aviation",
    "crop":                                                     "crop_insurance",
    "crop insurance":                                           "crop_insurance",
    "weather / crop insurance":                                 "crop_insurance",
    "other segment (a)":                                        "other_segments",
    "other segments (a)":                                       "other_segments",
    "other segment (crop)":                                     "other_segments",
    "other segments":                                           "other_segments",
    "miscellaneous":                                            "total_miscellaneous",
    "others":                                                   "total_miscellaneous",
    "total":                                                    "grand_total",
}

COMPANY_SPECIFIC_ALIASES: dict = {}
