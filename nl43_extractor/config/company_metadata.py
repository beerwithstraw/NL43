"""
Company Metadata Dictionary — maps internal company_key to Master Sheet fields.

Each entry provides:
  - company_name:  Short display name for Column C (Company_Name)
  - sorted_company: Sortable prefixed name for Column D (Company)
  - sector:        Insurer category for Column J (Sector)
  - competitors:   Competitive grouping for Column K (Industry Competitors)
"""

COMPANY_METADATA = {
    "bajaj_allianz": {
        "company_name": "BGIL",
        "sorted_company": "00-BGIL",
        "sector": "Private Sector Insurers",
        "competitors": "BAGIC",
    },
    "icici_lombard": {
        "company_name": "ICICILombard",
        "sorted_company": "01-ICICILombard",
        "sector": "Private Sector Insurers",
        "competitors": "Top 5 Private Players",
    },
    "tata_aig": {
        "company_name": "TATAAIG",
        "sorted_company": "02-TATAAIG",
        "sector": "Private Sector Insurers",
        "competitors": "Top 5 Private Players",
    },
    "hdfc_ergo": {
        "company_name": "HDFCERGO",
        "sorted_company": "03-HDFCERGO",
        "sector": "Private Sector Insurers",
        "competitors": "Top 5 Private Players",
    },
    "go_digit": {
        "company_name": "GODIGIT",
        "sorted_company": "04-GODIGIT",
        "sector": "Private Sector Insurers",
        "competitors": "Top 5 Private Players",
    },
    "sbi_general": {
        "company_name": "SBI",
        "sorted_company": "05-SBI",
        "sector": "Private Sector Insurers",
        "competitors": "Top 5 Private Players",
    },
    "star_health": {
        "company_name": "STAR",
        "sorted_company": "06-STAR",
        "sector": "SAHI",
        "competitors": "Others",
    },
    "future_generali": {
        "company_name": "GENERALICENTRAL",
        "sorted_company": "07-GENERALICENTRAL",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "new_india": {
        "company_name": "NEWINDIA",
        "sorted_company": "08-NEWINDIA",
        "sector": "Public Sector Insurers",
        "competitors": "Others",
    },
    "oriental_insurance": {
        "company_name": "ORIENTAL",
        "sorted_company": "09-ORIENTAL",
        "sector": "Public Sector Insurers",
        "competitors": "Others",
    },
    "indusind_general": {
        "company_name": "INDUSIND",
        "sorted_company": "10-INDUSIND",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "national_insurance": {
        "company_name": "NIC",
        "sorted_company": "11-NIC",
        "sector": "Public Sector Insurers",
        "competitors": "Others",
    },
    "united_india": {
        "company_name": "UNITEDINDIA",
        "sorted_company": "12-UNITEDINDIA",
        "sector": "Public Sector Insurers",
        "competitors": "Others",
    },
    "care_health": {
        "company_name": "CARE_HEALTH",
        "sorted_company": "13-CARE_HEALTH",
        "sector": "SAHI",
        "competitors": "Others",
    },
    "chola_ms": {
        "company_name": "CHOLA_MS",
        "sorted_company": "14-CHOLA_MS",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "niva_bupa": {
        "company_name": "NIVABUPA",
        "sorted_company": "15-NIVABUPA",
        "sector": "SAHI",
        "competitors": "Others",
    },
    "iffco_tokio": {
        "company_name": "IFFCOTokio",
        "sorted_company": "16-IFFCOTokio",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "aditya_birla_health": {
        "company_name": "ADITYA_BIRLA",
        "sorted_company": "17-ADITYA_BIRLA",
        "sector": "SAHI",
        "competitors": "Others",
    },
    "universal_sompo": {
        "company_name": "UNIVERSAL_SOMPO",
        "sorted_company": "18-UNIVERSAL_SOMPO",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "royal_sundaram": {
        "company_name": "ROYAL_SUNDARAM",
        "sorted_company": "19-ROYAL_SUNDARAM",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "magma_general": {
        "company_name": "MAGMA_HDI",
        "sorted_company": "20-MAGMA_HDI",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "liberty_general": {
        "company_name": "LIBERTY",
        "sorted_company": "21-LIBERTY",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "manipal_cigna": {
        "company_name": "MANIPAL_CIGNA",
        "sorted_company": "22-MANIPAL_CIGNA",
        "sector": "SAHI",
        "competitors": "Others",
    },
    "zurich_kotak": {
        "company_name": "KOTAK",
        "sorted_company": "23-KOTAK",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "zuno": {
        "company_name": "ZUNO",
        "sorted_company": "24-ZUNO",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "raheja_qbe": {
        "company_name": "RAHEJA_QBE",
        "sorted_company": "25-RAHEJA_QBE",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "aic": {
        "company_name": "AIC",
        "sorted_company": "26-AIC",
        "sector": "Specialized Insurers",
        "competitors": "Others",
    },
    "kshema_general": {
        "company_name": "KSHEMA",
        "sorted_company": "27-KSHEMA",
        "sector": "Specialized Insurers",
        "competitors": "Others",
    },
    "shriram_general": {
        "company_name": "SHRIRAM",
        "sorted_company": "28-SHRIRAM",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "acko": {
        "company_name": "ACKO",
        "sorted_company": "29-ACKO",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "ecgc": {
        "company_name": "ECGC",
        "sorted_company": "30-ECGC",
        "sector": "Specialized Insurers",
        "competitors": "Others",
    },
    "navi_general": {
        "company_name": "NAVI",
        "sorted_company": "31-NAVI",
        "sector": "Private Sector Insurers",
        "competitors": "Others",
    },
    "narayana_health": {
        "company_name": "NARAYANA_HEALTH",
        "sorted_company": "32-NARAYANA_HEALTH",
        "sector": "SAHI",
        "competitors": "Others",
    },
}

def get_metadata(company_key: str) -> dict:
    """Look up metadata for a company_key. Returns empty-string defaults if unknown."""
    return COMPANY_METADATA.get(company_key, {
        "company_name": company_key,
        "sorted_company": company_key,
        "sector": "",
        "competitors": "",
    })
