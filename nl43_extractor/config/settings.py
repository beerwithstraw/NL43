"""
Global constants and configuration for the NL-43 Extractor.
"""

# --- Versioning ---
EXTRACTOR_VERSION = "1.0.0"

# --- Default Paths ---
DEFAULT_INPUT_DIR = "inputs"
DEFAULT_OUTPUT_DIR = "outputs"

# --- FY Year String Helper ---

def make_fy_string(start_year: int, end_year: int) -> str:
    """Build the 6-character FY string. e.g. start=2025, end=2026 → '202526'"""
    return f"{start_year}{end_year % 100:02d}"

QUARTER_TO_FY = {
    "Q1": lambda y: make_fy_string(y, y + 1),
    "Q2": lambda y: make_fy_string(y, y + 1),
    "Q3": lambda y: make_fy_string(y, y + 1),
    "Q4": lambda y: make_fy_string(y - 1, y),
}

# --- Master Sheet Column Order ---
# NL-43: single period, one row per LOB per Segment (Rural / Social).
# "Segment" column holds "Rural" or "Social".
# Metric column names must .lower() to canonical keys in row_registry.py
# e.g. "Policies_Issued".lower() == "policies_issued" ✓
MASTER_COLUMNS = [
    "LOB_PARTICULARS",      # A
    "Grouped_LOB",          # B
    "Company_Name",         # C
    "Company",              # D
    "NL",                   # E — always "NL43"
    "Quarter",              # F
    "Year",                 # G — FY end year
    "Segment",              # H — "Rural" or "Social"
    "Sector",               # I
    "Industry_Competitors", # J
    "GI_Companies",         # K
    "Policies_Issued",      # L
    "Premium_Collected",    # M
    "Sum_Assured",          # N
    "Source_File",          # O
]

# --- Excel Formatting ---
NUMBER_FORMAT = "#,##0.00"
LOW_CONFIDENCE_FILL_COLOR = "FFFF99"

# --- Common Helpers ---

def company_key_to_pascal(company_key: str) -> str:
    """Convert snake_case company key to PascalCase. e.g. 'bajaj_allianz' → 'BajajAllianz'"""
    return company_key.replace("_", " ").title().replace(" ", "")
