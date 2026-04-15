"""
path_scanner.py

Walks the folder structure defined in extraction_config.yaml and returns
a list of ScanResult objects — one per PDF file to be processed.

Does NOT perform extraction. Only discovers files and derives metadata
from folder paths and filenames.
"""

import os
import hashlib
import logging
import re
from dataclasses import dataclass
from typing import List, Optional, Dict, Any

from config.company_registry import COMPANY_MAP

logger = logging.getLogger(__name__)


@dataclass
class ScanResult:
    pdf_path: str           # absolute path to the PDF
    company_key: str        # e.g. "bajaj_allianz"
    company_raw: str        # raw name from filename e.g. "BajajGeneral"
    quarter: str            # "Q1", "Q2", "Q3", "Q4"
    fiscal_year: str        # "FY2025"
    year_code: str          # "202425"
    source_type: str        # "direct" or "consolidated"
    file_hash: str          # MD5 hash of file contents


def _fy_to_year_code(fiscal_year: str) -> str:
    """
    Convert fiscal year string to year code.
    FY2025 -> "202425"
    FY2026 -> "202526"
    """
    try:
        y = int(fiscal_year.replace("FY", ""))
        return f"20{str(y-1)[-2:]}20{str(y)[-2:]}"
    except (ValueError, IndexError):
        logger.warning(f"Could not parse fiscal year: {fiscal_year}")
        return ""


def _extract_company_key(filename: str) -> Optional[tuple]:
    """
    Extract company key from any PDF filename by splitting on '_' and
    trying progressively longer suffixes against COMPANY_MAP.

    E.g. 'NL4_2024_25_Q1_BajajGeneral.pdf' -> splits to
    ['NL4', '2024', '25', 'Q1', 'BajajGeneral'] -> tries 'BajajGeneral',
    then 'Q1BajajGeneral', etc. until COMPANY_MAP match.

    Returns (company_key, company_raw) or None if not found.
    """
    # Strip .pdf extension
    name = filename
    if name.lower().endswith(".pdf"):
        name = name[:-4]

    parts = re.split(r'[_\-]', name)

    # Try progressively longer suffixes (shortest first = rightmost parts)
    for length in range(1, len(parts) + 1):
        suffix_parts = parts[len(parts) - length:]
        candidate = "".join(suffix_parts).lower().replace(" ", "")
        company_raw = "_".join(suffix_parts)

        # Check against COMPANY_MAP (longest key first for best match)
        for key in sorted(COMPANY_MAP.keys(), key=len, reverse=True):
            normalised_key = key.lower().replace("_", "").replace("-", "").replace(" ", "")
            if normalised_key == candidate or normalised_key in candidate:
                return (COMPANY_MAP[key], company_raw)

    logger.warning(f"Could not match company from filename: {filename}")
    return None


def _file_hash(path: str) -> str:
    """Compute MD5 hash of file contents for change detection."""
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _resolve_quarters(quarters_config) -> List[str]:
    """Resolve quarters config to list of quarter strings."""
    if quarters_config == "all" or quarters_config == ["all"]:
        return ["Q1", "Q2", "Q3", "Q4"]
    if isinstance(quarters_config, list):
        return [str(q).strip() for q in quarters_config]
    return ["Q1", "Q2", "Q3", "Q4"]


def scan(config: Dict[str, Any]) -> List[ScanResult]:
    """
    Walk the folder structure and return all PDFs to be processed.

    Priority rule: if a company has a direct NL43/ file, it takes priority
    over a consolidated file in the same quarter. Consolidated is ignored
    in that case.
    """
    base_path = config.get("base_path", "").strip()
    fiscal_years = config.get("fiscal_years", [])
    quarters = _resolve_quarters(config.get("quarters", "all"))

    if not base_path:
        raise ValueError("base_path is not set in extraction_config.yaml")
    if not os.path.exists(base_path):
        raise FileNotFoundError(f"base_path does not exist: {base_path}")

    results: List[ScanResult] = []

    for fy in fiscal_years:
        fy_path = os.path.join(base_path, str(fy))
        if not os.path.isdir(fy_path):
            logger.warning(f"Fiscal year folder not found, skipping: {fy_path}")
            continue

        year_code = _fy_to_year_code(str(fy))

        for quarter in quarters:
            q_path = os.path.join(fy_path, quarter)
            if not os.path.isdir(q_path):
                logger.debug(f"Quarter folder not found, skipping: {q_path}")
                continue

            nl39_path = os.path.join(q_path, "NL43")
            consolidated_path = os.path.join(q_path, "Consolidated")
            direct_companies = set()

            # --- Scan NL43/ subfolder: any .pdf here is a direct NL-39 form ---
            if os.path.isdir(nl39_path):
                for fname in os.listdir(nl39_path):
                    if not fname.lower().endswith(".pdf"):
                        continue

                    result = _extract_company_key(fname)
                    if result is None:
                        continue

                    company_key, company_raw = result
                    pdf_path = os.path.join(nl39_path, fname)

                    results.append(ScanResult(
                        pdf_path=os.path.abspath(pdf_path),
                        company_key=company_key,
                        company_raw=company_raw,
                        quarter=quarter,
                        fiscal_year=str(fy),
                        year_code=year_code,
                        source_type="direct",
                        file_hash=_file_hash(pdf_path),
                    ))
                    direct_companies.add(company_key)

            # --- Scan Consolidated/ subfolder ---
            # Only pick up consolidated PDFs for companies that don't have a direct NL43 file
            if os.path.isdir(consolidated_path):
                for fname in os.listdir(consolidated_path):
                    if not fname.lower().endswith(".pdf"):
                        continue

                    result = _extract_company_key(fname)
                    if result is None:
                        continue

                    company_key, company_raw = result

                    # Skip if we already have a direct file for this company
                    if company_key in direct_companies:
                        logger.debug(
                            f"Skipping consolidated {fname} — direct NL43 exists for {company_key}"
                        )
                        continue

                    pdf_path = os.path.join(consolidated_path, fname)

                    results.append(ScanResult(
                        pdf_path=os.path.abspath(pdf_path),
                        company_key=company_key,
                        company_raw=company_raw,
                        quarter=quarter,
                        fiscal_year=str(fy),
                        year_code=year_code,
                        source_type="consolidated",
                        file_hash=_file_hash(pdf_path),
                    ))
                    direct_companies.add(company_key)

    logger.info(
        f"Scan complete: {len(results)} PDFs found "
        f"({sum(1 for r in results if r.source_type == 'direct')} direct, "
        f"{sum(1 for r in results if r.source_type == 'consolidated')} consolidated)"
    )
    return results
