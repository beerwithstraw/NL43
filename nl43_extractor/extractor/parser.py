"""
NL-43 Generic Parser (Rural & Social Obligations).

NL-43 table layout:
  - PDF rows   = LOBs (Fire, Marine Cargo, Motor OD, ...)
                 Each LOB has TWO rows: Rural and Social
  - PDF cols   = Sl.No. | Line of Business | Particular | Policies Issued |
                 Premium Collected | Sum Assured
  - Single page, single period ("Upto the Quarter")

Output model:
  data[lob_key][metric_key] = {"rural": <float|None>, "social": <float|None>}
  metric_key ∈ {"policies_issued", "premium_collected", "sum_assured"}

Only extract.current_year is populated; extract.prior_year is always None.
"""

import logging
from pathlib import Path

import pdfplumber

from config.company_registry import COMPANY_DISPLAY_NAMES, DEDICATED_PARSER
from config.lob_registry import LOB_ALIASES, COMPANY_SPECIFIC_ALIASES
from config.row_registry import ROW_ORDER
from extractor.models import CompanyExtract, PeriodData
from extractor.normaliser import clean_number, normalise_text

logger = logging.getLogger(__name__)

# Default column indices (0-based). Used when no header row is detected.
# Col 0: Sl.No.  Col 1: LOB  Col 2: Particular  Col 3: Policies  Col 4: Premium  Col 5: Sum Assured
_DEFAULT_LOB_COL = 1
_DEFAULT_PARTICULAR_COL = 2
_DEFAULT_METRIC_COL_MAP = {
    3: "policies_issued",
    4: "premium_collected",
    5: "sum_assured",
}


def _detect_column_layout(table):
    """
    Scan the first few rows for a header containing 'PARTICULAR'.
    Returns (lob_col, particular_col, metric_col_map).
    Falls back to defaults when no header is found.
    """
    for row in table[:6]:
        if not row:
            continue
        upper = [str(c or "").upper().strip() for c in row]
        if "PARTICULAR" not in upper:
            continue
        particular_col = upper.index("PARTICULAR")
        lob_col = particular_col - 1
        metric_col_map = {}
        for i, cell in enumerate(upper):
            if "POLICIES" in cell:
                metric_col_map[i] = "policies_issued"
            elif "PREMIUM" in cell:
                metric_col_map[i] = "premium_collected"
            elif "SUM" in cell:
                metric_col_map[i] = "sum_assured"
        if metric_col_map:
            return lob_col, particular_col, metric_col_map
    return _DEFAULT_LOB_COL, _DEFAULT_PARTICULAR_COL, _DEFAULT_METRIC_COL_MAP


def _resolve_lob(raw_label: str, company_key: str):
    """Normalise a PDF row label → canonical LOB key. Returns None for headers/blanks."""
    normalised = normalise_text(str(raw_label or ""))
    if not normalised:
        return None
    company_aliases = COMPANY_SPECIFIC_ALIASES.get(company_key, {})
    if normalised in company_aliases:
        return company_aliases[normalised]
    return LOB_ALIASES.get(normalised)


def _extract_table(table, company_key: str, period_data: PeriodData) -> int:
    """
    Parse one NL-43 table into period_data.

    Each LOB occupies two consecutive rows:
      Row A: [Sl.No, LOB_name, "Rural", policies, premium, sum_assured]
      Row B: [None,  None,     "Social", policies, premium, sum_assured]

    Column positions are auto-detected from the header row to handle PDFs
    where an extra leading column shifts everything right (e.g. Aditya Birla).

    Returns number of LOB keys extracted.
    """
    lob_col, particular_col, metric_col_map = _detect_column_layout(table)
    current_lob = None
    lobs_found = 0

    for row in table:
        if not row or len(row) <= particular_col:
            continue

        lob_label = row[lob_col] if lob_col < len(row) else None
        particular = normalise_text(str(row[particular_col] or ""))

        # Update current LOB when the LOB column has a LOB name
        if lob_label and str(lob_label).strip():
            candidate = _resolve_lob(lob_label, company_key)
            if candidate is not None:
                current_lob = candidate
                lobs_found += 1

        if current_lob is None:
            continue

        # Only process Rural / Social rows
        if particular not in ("rural", "social"):
            continue

        if current_lob not in period_data.data:
            period_data.data[current_lob] = {}

        for col_idx, metric_key in metric_col_map.items():
            if col_idx >= len(row):
                continue
            val = clean_number(row[col_idx])
            if metric_key not in period_data.data[current_lob]:
                period_data.data[current_lob][metric_key] = {"rural": None, "social": None}
            if val is not None:
                period_data.data[current_lob][metric_key][particular] = val

    return lobs_found


def parse_pdf(pdf_path: str, company_key: str, quarter: str = "", year: str = "") -> CompanyExtract:
    """
    Main entry point — parses one NL-43 PDF.

    NL-43 is a single-page, single-table form. Both Rural and Social rows
    are extracted into a single PeriodData stored in extract.current_year.
    """
    logger.info(f"Parsing NL-43 PDF: {pdf_path} for company: {company_key}")

    company_name = COMPANY_DISPLAY_NAMES.get(company_key, str(company_key).title())

    # Route to dedicated parser if registered
    dedicated_func_name = DEDICATED_PARSER.get(company_key)
    if dedicated_func_name:
        from extractor.companies import PARSER_REGISTRY
        dedicated_func = PARSER_REGISTRY.get(dedicated_func_name)
        if dedicated_func:
            logger.info(f"Routing to dedicated parser: {dedicated_func_name}")
            return dedicated_func(pdf_path, company_key, quarter, year)

    extract = CompanyExtract(
        source_file=Path(pdf_path).name,
        company_key=company_key,
        company_name=company_name,
        form_type="NL43",
        quarter=quarter,
        year=year,
    )

    period_data = PeriodData(period_label="current")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # NL-43 is always a single page/table — scan all pages, take first table found
            for pg_idx, page in enumerate(pdf.pages):
                table = page.extract_table()
                if not table:
                    continue
                n = _extract_table(table, company_key, period_data)
                logger.debug(f"Page {pg_idx}: {n} LOBs extracted")
                if n > 0:
                    break   # found data — don't process further pages

    except Exception as e:
        logger.error(f"Failed to parse {pdf_path}: {e}", exc_info=True)
        extract.extraction_errors.append(str(e))
        return extract

    lobs_with_data = len(period_data.data)
    if lobs_with_data == 0:
        logger.warning(f"No LOBs extracted from {pdf_path}")
        extract.extraction_warnings.append("No LOBs extracted")
    else:
        logger.info(f"Extraction complete: {lobs_with_data} LOBs")

    extract.current_year = period_data
    return extract
