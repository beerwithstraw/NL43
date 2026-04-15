"""
LOB Metadata — maps internal LOB keys to Master Sheet Column A/B values.

Column A: LOB_PARTICULARS (sorted prefix + key)
Column B: Grouped_LOB (display grouping)
"""

LOB_METADATA = {
    "fire":                     ("08-FIRE",                    "FIRE"),
    "marine_cargo":             ("09-MARINE_CARGO",            "MARINE CARGO"),
    "marine_hull":              ("10-MARINE_HULL",             "MARINE HULL"),
    "total_marine":             ("11-MARINE_TOTAL",            "MARINE TOTAL"),
    "motor_od":                 ("01-MOTOR_OD",                "MOTOR OD"),
    "motor_tp":                 ("02-MOTOR_TP",                "MOTOR TP"),
    "total_motor":              ("03-MOTOR_TOTAL",             "MOTOR TOTAL"),
    "health":                   ("04-HEALTH",                  "HEALTH"),
    "personal_accident":        ("05-PERSONAL_ACCIDENT",       "PERSONAL ACCIDENT"),
    "travel_insurance":         ("06-TRAVEL_INSURANCE",        "TRAVEL INSURANCE"),
    "total_health":             ("07-HEALTH_TOTAL",            "HEALTH TOTAL"),
    "wc_el":                    ("13-WC",                      "WC"),
    "public_product_liability": ("14-PUBLIC_PRODUCT_LIABILITY", "PUBLIC/PRODUCT LIABILITY"),
    "engineering":              ("12-ENGINEERING",              "ENGINEERING"),
    "aviation":                 ("16-AVIATION",                "AVIATION"),
    "crop_insurance":           ("21-CROP_INSURANCE",          "CROP INSURANCE"),
    "credit_insurance":         ("17-CREDIT_INSURANCE",        "CREDIT INSURANCE"),
    "other_liability":          ("15-OTHER_LIABILITY",         "MISCELLANEOUS"),
    "specialty":                ("20-SPECIALITY",              "MISCELLANEOUS"),
    "home":                     ("18-HOME",                    "MISCELLANEOUS"),
    "other_segments":           ("19-OTHER_SEGMENT",           "MISCELLANEOUS"),
    "other_miscellaneous":      ("22-MISCELLANEOUS_SEGMENT",   "MISCELLANEOUS"),
    "total_miscellaneous":      ("23-TOTAL_MISCELLANEOUS",     "MISCELLANEOUS TOTAL"),
    "grand_total":              ("24-GRAND_TOTAL",             "GRAND TOTAL"),
}


def get_lob_particulars(lob_key: str) -> str:
    """Return LOB_PARTICULARS (Column A) for a given LOB key."""
    entry = LOB_METADATA.get(lob_key)
    return entry[0] if entry else lob_key


def get_grouped_lob(lob_key: str) -> str:
    """Return Grouped_LOB (Column B) for a given LOB key."""
    entry = LOB_METADATA.get(lob_key)
    return entry[1] if entry else lob_key
