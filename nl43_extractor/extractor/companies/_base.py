"""
Shared base utilities for NL-43 dedicated parsers.

The generic parser in extractor/parser.py handles all companies by default.
This module provides helpers for dedicated parsers when company-specific
quirks are encountered.
"""

import logging

logger = logging.getLogger(__name__)


def get_nl43_pages(pdf) -> list:
    """
    Return only the pages that belong to FORM NL-43.

    For standalone NL-43 PDFs this returns all pages.
    For consolidated PDFs, filters pages containing NL-43 keywords.
    """
    keywords = {"nl-43", "nl 43", "rural & social", "sum assured"}
    all_pages = pdf.pages
    if len(all_pages) <= 2:
        return list(all_pages)

    nl43_pages = []
    for page in all_pages:
        text = (page.extract_text() or "").lower()
        if any(kw in text for kw in keywords):
            nl43_pages.append(page)

    return nl43_pages if nl43_pages else list(all_pages)
