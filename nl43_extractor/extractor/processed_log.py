"""
processed_log.py

Tracks which PDF files have been successfully extracted.
Enables incremental extraction — only new or changed files are processed.

A file is considered "already processed" if:
  - Its absolute path exists in the log AND
  - Its MD5 hash matches the stored hash (file has not changed)

If the PDF changes (new version uploaded), the hash changes and the file
is re-extracted on the next run.
"""

import json
import logging
import os
from datetime import datetime
from typing import List, Dict, Any, Optional

from extractor.path_scanner import ScanResult

logger = logging.getLogger(__name__)

LOG_VERSION = 1


def load(log_path: str) -> Dict[str, Any]:
    """Load the processed log from disk. Returns empty log if file doesn't exist."""
    if not os.path.exists(log_path):
        return {"version": LOG_VERSION, "processed": {}}
    try:
        with open(log_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data
    except (json.JSONDecodeError, IOError) as e:
        logger.warning(f"Could not read processed log ({e}), starting fresh")
        return {"version": LOG_VERSION, "processed": {}}


def save(log_path: str, log_data: Dict[str, Any]) -> None:
    """Save the processed log to disk."""
    os.makedirs(os.path.dirname(os.path.abspath(log_path)), exist_ok=True)
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(log_data, f, indent=2)


def filter_unprocessed(
    scan_results: List[ScanResult],
    log_data: Dict[str, Any],
    force: bool = False,
    force_company: Optional[str] = None,
) -> List[ScanResult]:
    """
    Return only the ScanResults that need to be (re-)extracted.

    force=True: return all results regardless of log
    force_company='bajaj_allianz': return all results for that company
    Otherwise: return results whose path is not in the log OR whose hash changed
    """
    if force:
        logger.info("--force flag set: re-extracting all files")
        return scan_results

    processed = log_data.get("processed", {})
    to_process = []

    for result in scan_results:
        path = result.pdf_path

        if force_company and result.company_key == force_company:
            logger.info(f"--force-company: re-extracting {path}")
            to_process.append(result)
            continue

        if path not in processed:
            to_process.append(result)
            continue

        stored_hash = processed[path].get("file_hash", "")
        if stored_hash != result.file_hash:
            logger.info(f"File changed (hash mismatch), re-extracting: {path}")
            to_process.append(result)

    skipped = len(scan_results) - len(to_process)
    logger.info(
        f"Incremental filter: {len(to_process)} to process, {skipped} already up-to-date"
    )
    return to_process


def mark_processed(
    log_data: Dict[str, Any],
    result: ScanResult,
    rows_written: int,
) -> None:
    """Record a successfully processed file in the log."""
    log_data["processed"][result.pdf_path] = {
        "file_hash": result.file_hash,
        "processed_at": datetime.now().isoformat(),
        "company_key": result.company_key,
        "quarter": result.quarter,
        "fiscal_year": result.fiscal_year,
        "year_code": result.year_code,
        "source_type": result.source_type,
        "rows_written": rows_written,
    }
