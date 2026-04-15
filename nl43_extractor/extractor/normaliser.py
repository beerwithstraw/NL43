"""
Cell normalisation functions for NL-4 PDF extraction.

clean_number()   — converts any raw cell value to Python float or None
normalise_text() — normalises a row/column label for fuzzy matching

Source: approach document Section 9
Anti-hallucination rule #1: never invent values — return None on failure.
"""

import re


# Strings that represent "no value" in NL-4 PDFs
NIL_STRINGS = frozenset({
    "-", "--", "- -", "–", "—",      # ASCII dash, double-dash, en-dash, em-dash
    "n/a", "na", "nil",
})


def clean_number(raw):
    """
    Convert any raw cell value to Python float or None.

    Processing steps (order matters):
        1. None / empty / whitespace → None
        2. Nil strings ("-", "N/A", "nil", "NIL", "NA", "--", dashes) → None
        3. Strip leading/trailing whitespace
        4. If company_key in SPACE_BROKEN_NUMBERS: remove spaces within
           digit sequences (e.g. "3 4,193" → "34,193")
        5. Remove commas (handles both "1,234" and "1,24,941" Indian grouping)
        6. Unicode dashes: \\u2013 (–) and \\u2014 (—) → "-"
        7. Parentheses: "(500)" → "-500"
        8. Attempt float(cleaned) — return float or None on failure

    Always returns float or None. Never raises.

    Parameters
    ----------
    raw : str | None
        The raw cell text from pdfplumber or openpyxl.

    Returns
    -------
    float | None
    """
    # 1. None / empty / whitespace
    if raw is None:
        return None
    if not isinstance(raw, str):
        # If it's already a number (int/float), return as float
        try:
            return float(raw)
        except (TypeError, ValueError):
            return None

    stripped = raw.strip()
    if not stripped:
        return None

    # 2. Nil strings (case-insensitive)
    if stripped.lower() in NIL_STRINGS:
        return None

    # 3. Working copy (strip newlines and stray alphabets from PDF leakages e.g. "t\n3,78,997")
    cleaned = stripped.replace('\n', '').replace('\r', '')
    cleaned = re.sub(r'^[a-zA-Z]+', '', cleaned).strip()
    cleaned = re.sub(r'[a-zA-Z]+$', '', cleaned).strip()

    # 4. Remove commas (works for both Western "1,234" and Indian "1,24,941")
    cleaned = cleaned.replace(",", "")

    # 5. Space-broken number repair (universal)
    #    "3 4193" → "34193", "0 .16" → "0.16", "1 .75" → "1.75"
    #    Apply iteratively in case multiple space breaks exist
    prev = None
    while prev != cleaned:
        prev = cleaned
        cleaned = re.sub(r'(\d)\s+(\d)', r'\1\2', cleaned)
        cleaned = re.sub(r'(\d)\s+\.', r'\1.', cleaned)

    # 6. Unicode dashes → ASCII hyphen-minus
    cleaned = cleaned.replace("\u2013", "-")   # en-dash
    cleaned = cleaned.replace("\u2014", "-")   # em-dash

    # 7. Parentheses → negative: "(500)" → "-500"
    #    Also handles "(  31.20)" — strip inner whitespace first
    paren_match = re.match(r'^\(\s*(.+?)\s*\)$', cleaned)
    if paren_match:
        cleaned = "-" + paren_match.group(1)

    # 8. Parse float
    try:
        return float(cleaned)
    except ValueError:
        return None


def normalise_text(raw):
    """
    Normalise a row or column label for matching against alias registries.

    Steps:
        1. Strip whitespace and newlines
        2. Lowercase
        3. Remove all non-alphanumeric characters except spaces and forward slashes
        4. Collapse multiple spaces to one
        5. Strip

    Returns empty string if None or empty.

    Parameters
    ----------
    raw : str | None

    Returns
    -------
    str
    """
    if raw is None:
        return ""

    if not isinstance(raw, str):
        raw = str(raw)

    # Strip whitespace and newlines
    text = raw.strip().replace("\n", " ").replace("\r", " ")

    # Lowercase
    text = text.lower()

    # Remove non-alphanumeric except spaces, forward slashes, apostrophes,
    # periods, colons, hyphens, and parentheses (these appear in known aliases)
    # Actually, let's keep it simple: remove only characters that would never
    # appear in an alias. The alias dict already has the normalised forms.
    # Strategy: keep alphanumeric, spaces, and common punctuation in aliases.
    # Looking at ROW_ALIASES and LOB_ALIASES, they contain: letters, digits,
    # spaces, colons, slashes, apostrophes (both kinds), hyphens, periods,
    # parentheses.
    # So we keep all of those and remove everything else.
    text = re.sub(r"[^a-z0-9\s/:'\u2019()\-.]", "", text)

    # Collapse multiple spaces
    text = re.sub(r'\s+', ' ', text)

    return text.strip()
