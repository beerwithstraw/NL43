"""
Generic table collection logic.

Source: approach document Section 9 (Collector)
"""

import pdfplumber
import logging
from typing import List

from config.settings import COLLECTOR_SNAP_TOLERANCE_LINES

logger = logging.getLogger(__name__)


def collect_tables(pdf_path: str, extraction_strategy: str = "lines") -> list[dict]:
    """
    Returns list of table objects, each with page metadata preserved.
    
    Each item:
    {
        "page": int,           # 1-based page number
        "table_index": int,    # index within page (0, 1, 2...)
        "rows": list[list[str]]
    }
    
    The parser is responsible for identifying CY vs PY blocks
    and merging across pages. The collector only collects.
    """
    # For Phase 1, we handle the standard 'lines' strategy.
    # Note: 'positional' strategy also feeds off the standard lines extraction
    # initially to get the raw text grid, then positional logic maps the columns.
    
    settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": COLLECTOR_SNAP_TOLERANCE_LINES,
        "snap_x_tolerance": COLLECTOR_SNAP_TOLERANCE_LINES,
        "snap_y_tolerance": COLLECTOR_SNAP_TOLERANCE_LINES,
        "intersection_tolerance": COLLECTOR_SNAP_TOLERANCE_LINES,
        "join_tolerance": COLLECTOR_SNAP_TOLERANCE_LINES,
        "join_x_tolerance": COLLECTOR_SNAP_TOLERANCE_LINES,
        "join_y_tolerance": COLLECTOR_SNAP_TOLERANCE_LINES,
    }

    if extraction_strategy == "text":
        settings["vertical_strategy"] = "text"
        settings["horizontal_strategy"] = "lines" # Usually text + lines works

    try:
        table_data = []
        with pdfplumber.open(pdf_path) as pdf:
            from extractor.companies._base import get_nl4_pages
            for i, page in enumerate(get_nl4_pages(pdf)):
                logger.debug(f"Extracting table from page {i+1} using {extraction_strategy} strategy")
                
                # Copy settings so we can modify them per page if needed
                page_settings = settings.copy()
                

                # Check if it's HDFC Ergo where _split_on_header is explicitly NOT needed,
                # but we need to merge multiple table objects on the SAME page.
                # pdfplumber page.extract_tables() returns a list of tables for a single page.
                tables = page.extract_tables(table_settings=page_settings)
                
                if not tables:
                    logger.debug(f"No tables found on page {i+1}")
                    continue
                    
                for t_idx, table in enumerate(tables):
                    # Clean up the row strings
                    cleaned_table = []
                    for row in table:
                        # Replace None with empty string and clean newlines
                        cleaned_row = [
                            str(cell).strip() if cell is not None else "" 
                            for cell in row
                        ]
                        # Only add rows that aren't entirely empty
                        if any(cell for cell in cleaned_row):
                            cleaned_table.append(cleaned_row)
                            
                    if cleaned_table and max(len(r) for r in cleaned_table) >= 3:
                        table_data.append({
                            "page": i + 1,
                            "table_index": t_idx,
                            "rows": cleaned_table
                        })
                        
        if not table_data:
            logger.warning(f"Extracted table is empty for {pdf_path}")
            
        return table_data

    except Exception as e:
        logger.error(f"Table extraction failed for {pdf_path}: {e}")
        return []
