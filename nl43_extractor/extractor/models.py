"""
Data models for the extracted data returned by generic parser.

Source: approach document Section 5.3
"""

from dataclasses import dataclass, field
from typing import Optional

@dataclass
class PeriodData:
    """Holds one table's worth of data — either Current Year or Prior Year."""
    period_label: str               # "current" or "prior"
    # data[lob_key][row_key] = {"qtr": float|None, "ytd": float|None}
    data: dict = field(default_factory=dict)
    # auto_negated_ri set to True if constitutional heuristic triggered mutation
    auto_negated_ri: bool = False
    # low_confidence_cells: set of (lob_key, row_key) tuples
    low_confidence_cells: set = field(default_factory=set)

@dataclass
class CompanyExtract:
    """Top-level container for one extracted PDF."""
    source_file: str
    company_key: str                # e.g. "bajaj_allianz"
    company_name: str               # e.g. "Bajaj Allianz General Insurance"
    form_type: str                  # "NL43"
    quarter: str                    # e.g. "Q1"
    year: str                       # e.g. "202526"
    current_year: Optional[PeriodData] = None
    prior_year: Optional[PeriodData] = None
    extraction_warnings: list = field(default_factory=list)
    extraction_errors: list = field(default_factory=list)
