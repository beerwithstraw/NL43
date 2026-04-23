"""
Company Registry for NL-43.

NL-43 (Rural & Social Obligations) is filed by all general insurers.
No EXTRACTION_STRATEGY or DEDICATED_PARSER needed initially: the generic
NL-43 parser handles all companies (table structure is fixed across all filers).
"""

COMPANY_MAP = {
    "acko": "acko",
    "acko general": "acko",
    "acko insurance": "acko",
    "aditya birla": "aditya_birla_health",
    "aditya birla health": "aditya_birla_health",
    "aditya birla health insurance": "aditya_birla_health",
    "adityabirla": "aditya_birla_health",
    "agriculture insurance": "aic",
    "agriculture insurance company": "aic",
    "agriculture insurance company of india": "aic",
    "aic": "aic",
    "aicof": "aic",
    "bajaj": "bajaj_allianz",
    "bajaj allianz": "bajaj_allianz",
    "bajajgeneral": "bajaj_allianz",
    "bgil": "bajaj_allianz",
    "care": "care_health",
    "care health": "care_health",
    "carehealth": "care_health",
    "chola": "chola_ms",
    "chola general": "chola_ms",
    "chola ms": "chola_ms",
    "cholamandalam": "chola_ms",
    "digit general": "go_digit",
    "ecgc": "ecgc",
    "ecgc limited": "ecgc",
    "edelweiss": "zuno",
    "edelweiss general": "zuno",
    "export credit guarantee": "ecgc",
    "future generali": "future_generali",
    "futuregenerali": "future_generali",
    "galaxy": "galaxy_health",
    "galaxy health": "galaxy_health",
    "galaxy health and allied": "galaxy_health",
    "galaxy health insurance": "galaxy_health",
    "galaxyhealth": "galaxy_health",
    "generali": "future_generali",
    "generali central": "future_generali",
    "generalicentral": "future_generali",
    "go digit": "go_digit",
    "godigit": "go_digit",
    "hdfc": "hdfc_ergo",
    "hdfc ergo": "hdfc_ergo",
    "hdfcergo": "hdfc_ergo",
    "icici": "icici_lombard",
    "icici lombard": "icici_lombard",
    "icici lombard general": "icici_lombard",
    "iffco": "iffco_tokio",
    "iffco tokio": "iffco_tokio",
    "iffco-tokio": "iffco_tokio",
    "iffcotokio": "iffco_tokio",
    "indusind": "indusind_general",
    "indusind general": "indusind_general",
    "indusind general insurance": "indusind_general",
    "kotak": "zurich_kotak",
    "kotak mahindra": "zurich_kotak",
    "kotak mahindra general": "zurich_kotak",
    "kshema": "kshema_general",
    "kshema general": "kshema_general",
    "kshema general insurance": "kshema_general",
    "liberty": "liberty_general",
    "liberty general": "liberty_general",
    "liberty videocon": "liberty_general",
    "libertygeneral": "liberty_general",
    "lombard": "icici_lombard",
    "magma": "magma_general",
    "magma general": "magma_general",
    "magma hdi": "magma_general",
    "manipal": "manipal_cigna",
    "manipal cigna": "manipal_cigna",
    "manipalcigna": "manipal_cigna",
    "narayana": "narayana_health",
    "narayana health": "narayana_health",
    "narayanahealth": "narayana_health",
    "national insurance": "national_insurance",
    "nationalinsurance": "national_insurance",
    "navi": "navi_general",
    "navi general": "navi_general",
    "new india": "new_india",
    "newindia": "new_india",
    "nic": "national_insurance",
    "niva bupa": "niva_bupa",
    "nivabupa": "niva_bupa",
    "oriental": "oriental_insurance",
    "oriental insurance": "oriental_insurance",
    "orientalinsurance": "oriental_insurance",
    "raheja": "raheja_qbe",
    "raheja qbe": "raheja_qbe",
    "reliance": "indusind_general",
    "reliance general": "indusind_general",
    "royal sundaram": "royal_sundaram",
    "sbi": "sbi_general",
    "sbi general": "sbi_general",
    "sgi": "shriram_general",
    "shriram": "shriram_general",
    "shriram general": "shriram_general",
    "sriram general": "shriram_general",
    "star": "star_health",
    "star health": "star_health",
    "star health and allied": "star_health",
    "starhealth": "star_health",
    "tata aig": "tata_aig",
    "tataaig": "tata_aig",
    "united india": "united_india",
    "united": "united_india",
    "unitedindia": "united_india",
    "universal sompo": "universal_sompo",
    "universalsompo": "universal_sompo",
    "zuno": "zuno",
    "zurich kotak": "zurich_kotak",
}

# ---------------------------------------------------------------------------
# Display names
# ---------------------------------------------------------------------------
COMPANY_DISPLAY_NAMES = {
    "acko": "ACKO General Insurance Limited",
    "aditya_birla_health": "Aditya Birla Health Insurance Co. Limited",
    "aic": "Agriculture Insurance Company of India Limited",
    "bajaj_allianz": "Bajaj Allianz General Insurance Company Limited",
    "care_health": "Care Health Insurance Limited",
    "chola_ms": "Cholamandalam MS General Insurance Company Limited",
    "ecgc": "ECGC Limited",
    "future_generali": "Future Generali India Insurance Company Limited",
    "galaxy_health": "Galaxy Health and Allied Insurance Company Limited",
    "go_digit": "Go Digit General Insurance Limited",
    "hdfc_ergo": "HDFC ERGO General Insurance",
    "icici_lombard": "ICICI Lombard General Insurance Company Limited",
    "iffco_tokio": "IFFCO Tokio General Insurance Company Limited",
    "indusind_general": "IndusInd General Insurance Company Limited",
    "kshema_general": "Kshema General Insurance Limited",
    "liberty_general": "Liberty General Insurance Company Limited",
    "magma_general": "Magma General Insurance Limited",
    "manipal_cigna": "Manipal Cigna Health Insurance Company Limited",
    "narayana_health": "Narayana Health Insurance Limited",
    "national_insurance": "National Insurance Company",
    "navi_general": "Navi General Insurance Limited",
    "new_india": "The New India Assurance Company",
    "niva_bupa": "Niva Bupa Health Insurance Company Limited",
    "oriental_insurance": "The Oriental Insurance Company",
    "raheja_qbe": "Raheja QBE General Insurance Company Limited",
    "royal_sundaram": "Royal Sundaram General Insurance Co. Limited",
    "sbi_general": "SBI General Insurance Company Limited",
    "shriram_general": "Shriram General Insurance Company Limited",
    "star_health": "Star Health and Allied Insurance Co. Ltd.",
    "tata_aig": "Tata AIG General Insurance Company Limited",
    "united_india": "United India Insurance Company",
    "universal_sompo": "Universal Sompo General Insurance Company Limited",
    "zuno": "ZUNO General Insurance Limited",
    "zurich_kotak": "Zurich Kotak General Insurance Company (India) Limited",
}

# ---------------------------------------------------------------------------
# Dedicated parser function name (empty = use generic NL-39 parser for all)
# Add entries here as company-specific quirks are discovered.
# ---------------------------------------------------------------------------
DEDICATED_PARSER: dict = {
    "bajaj_allianz": "parse_bajaj_allianz",
}

# ---------------------------------------------------------------------------
# Completeness ignore: LOBs that are known to be blank for a specific company.
# These LOBs will be SKIPPED in the COMPLETENESS check (no WARN or FAIL).
# ---------------------------------------------------------------------------
COMPLETENESS_IGNORE: dict = {
    # Bajaj Allianz: Motor OD blank — no rural/social OD exposure.
    "bajaj_allianz":       {"motor_od"},
    # Health-only insurers: fire and motor LOBs are not applicable.
    "aditya_birla_health": {"fire", "motor_od", "motor_tp"},
    "care_health":         {"fire", "motor_od", "motor_tp"},
    "galaxy_health":       {"fire", "motor_od", "motor_tp"},
    "manipal_cigna":       {"fire", "motor_od", "motor_tp"},
    "narayana_health":     {"fire", "motor_od", "motor_tp"},
    "niva_bupa":           {"fire", "motor_od", "motor_tp"},
    # AIC: agriculture insurer — fire, motor and health are not applicable.
    "aic":                 {"fire", "motor_od", "motor_tp", "health"},
    # General insurers with confirmed nil rural/social data for these LOBs (verified from PDF).
    "acko":                {"fire", "motor_od", "motor_tp", "health"},
    "shriram_general":     {"fire", "motor_od", "motor_tp", "health"},
    "navi_general":        {"fire", "motor_od", "motor_tp", "health"},
    "indusind_general":    {"fire", "motor_od", "motor_tp", "health"},
}

# ---------------------------------------------------------------------------
# Total-sum ignore: (metric, segment) pairs to skip the TOTAL_SUM check for
# a company. Use when the PDF itself documents why the sum won't reconcile.
# ---------------------------------------------------------------------------
TOTAL_SUM_IGNORE: dict = {
    # Magma: Motor Comprehensive (174,666) is counted in both Motor OD and Motor TP
    # but only once in the Total Rural row — documented in PDF footnote.
    "magma_general": {("policies_issued", "rural")},
}
