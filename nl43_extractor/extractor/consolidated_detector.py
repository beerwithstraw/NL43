"""
consolidated_detector.py

Finds the page range of the NL-43 form within a consolidated PDF.

Detection strategy:
  START: First page where >= min_matches NL-43 keywords appear
  END:   Page before the next form header, or last page of PDF

NL-43 is a single-page form (Rural & Social Obligations).
"""

import re
import logging
import tempfile
import os
from typing import Optional, Tuple, List

logger = logging.getLogger(__name__)

DEFAULT_KEYWORDS = [
    "NL-43",
    "RURAL & SOCIAL",
    "SUM ASSURED",     # always in the data table header; split column text prevents
                       # matching "POLICIES ISSUED" / "PREMIUM COLLECTED" across companies
]

# TOC pages match at most 2 keywords; the real form page matches 3+.
DEFAULT_MIN_MATCHES = 3

FORM_HEADER_PATTERN = re.compile(r"FORM\s+NL[-\s]?(\d+)", re.IGNORECASE)
TOC_SKIP_PATTERN = re.compile(
    r"TABLE\s+OF\s+CONTENTS|FORM\s+INDEX|INDEX\s+OF\s+FORMS",
    re.IGNORECASE,
)


def _page_keyword_count(text: str, keywords: List[str]) -> int:
    text_upper = text.upper()
    return sum(1 for kw in keywords if kw.upper() in text_upper)


def find_nl43_pages(
    pdf_path: str,
    keywords: Optional[List[str]] = None,
    min_matches: int = DEFAULT_MIN_MATCHES,
) -> Optional[Tuple[int, int]]:
    """
    Scan the consolidated PDF and return (start_page, end_page) 0-indexed
    for the NL-43 section. Returns None if not found.
    """
    try:
        import pdfplumber
    except ImportError:
        logger.error("pdfplumber not available")
        return None

    if keywords is None:
        keywords = DEFAULT_KEYWORDS

    try:
        with pdfplumber.open(pdf_path) as pdf:
            n_pages = len(pdf.pages)
            page_texts = [page.extract_text() or "" for page in pdf.pages]

        # Find start page
        start_page = None
        for i, text in enumerate(page_texts):
            if TOC_SKIP_PATTERN.search(text):
                logger.debug(f"  page {i + 1}: TOC page, skipping")
                continue
            if _page_keyword_count(text, keywords) >= min_matches:
                start_page = i
                break

        if start_page is None:
            logger.warning(f"NL-43 section not found in: {pdf_path}")
            return None

        # Find end page — stop when a DIFFERENT form number appears
        end_page = n_pages - 1
        for i in range(start_page + 1, n_pages):
            matches = FORM_HEADER_PATTERN.findall(page_texts[i])
            non_nl43 = [m for m in matches if m != "43"]
            if non_nl43:
                end_page = i - 1
                logger.debug(f"NL-43 ends at page {end_page} (NL-{non_nl43[0]} at page {i})")
                break

        logger.info(
            f"NL-43 found at pages {start_page}-{end_page} "
            f"(0-indexed) in {os.path.basename(pdf_path)}"
        )
        return (start_page, end_page)

    except Exception as e:
        logger.error(f"Error scanning consolidated PDF {pdf_path}: {e}")
        return None


def extract_nl43_to_temp(
    pdf_path: str,
    start_page: int,
    end_page: int,
) -> Optional[str]:
    """
    Extract pages start_page..end_page into a temporary PDF file.
    Caller must delete the temp file after use.
    """
    try:
        import pypdf
    except ImportError:
        try:
            import PyPDF2 as pypdf
        except ImportError:
            logger.error("pypdf or PyPDF2 not available")
            return None

    try:
        reader = pypdf.PdfReader(pdf_path)
        writer = pypdf.PdfWriter()
        for page_num in range(start_page, end_page + 1):
            if page_num < len(reader.pages):
                writer.add_page(reader.pages[page_num])

        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False, prefix="nl43_extract_")
        with open(tmp.name, "wb") as f:
            writer.write(f)

        logger.debug(f"Extracted pages {start_page}-{end_page} to {tmp.name}")
        return tmp.name

    except Exception as e:
        logger.error(f"Error extracting pages from {pdf_path}: {e}")
        return None
