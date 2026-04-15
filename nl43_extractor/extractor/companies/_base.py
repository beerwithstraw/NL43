"""
Shared base utilities for NL-39 dedicated parsers.

The generic parser in extractor/parser.py handles all companies by default.
This module provides helpers for dedicated parsers when company-specific
quirks are encountered.
"""

import logging

logger = logging.getLogger(__name__)


def get_nl39_pages(pdf) -> list:
    """
    Return only the pages that belong to FORM NL-39.

    For standalone NL-39 PDFs (always 2 pages) this returns all pages.
    For consolidated PDFs, filters pages containing NL-39 keywords.
    """
    keywords = {"nl-39", "nl 39", "ageing of claims", "ageing of claim"}
    all_pages = pdf.pages
    if len(all_pages) <= 2:
        return list(all_pages)

    nl39_pages = []
    for page in all_pages:
        text = (page.extract_text() or "").lower()
        if any(kw in text for kw in keywords):
            nl39_pages.append(page)

    return nl39_pages if nl39_pages else list(all_pages)
