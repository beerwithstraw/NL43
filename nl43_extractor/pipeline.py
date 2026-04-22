"""
pipeline.py — Config-driven NL-43 extraction pipeline.

Usage:
  python3 pipeline.py                          # incremental (default)
  python3 pipeline.py --force                  # re-extract everything
  python3 pipeline.py --force-company bajaj_allianz
  python3 pipeline.py --quarter Q1 Q3          # override config quarters
  python3 pipeline.py --dry-run                # show what would run, don't extract
  python3 pipeline.py --skip-consolidated      # ignore consolidated PDFs this run
"""

import argparse
import logging
import os
import sys
import yaml

# Ensure repo root is on path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from extractor.path_scanner import scan
from extractor.consolidated_detector import find_nl43_pages, extract_nl43_to_temp
from extractor.processed_log import load as load_log, save as save_log
from extractor.processed_log import filter_unprocessed, mark_processed
from extractor.parser import parse_pdf
from validation.checks import run_validations, write_validation_report
from output.excel_writer import (
    save_workbook,
    write_validation_summary_sheet,
    write_validation_detail_sheet,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-5s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "extraction_config.yaml")


def load_config(config_path: str) -> dict:
    if not os.path.exists(config_path):
        raise FileNotFoundError(
            f"Config file not found: {config_path}\n"
            f"Expected at: {config_path}"
        )
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def validate_config(config: dict) -> None:
    """Raise ValueError with helpful message if config is incomplete."""
    if not config.get("base_path", "").strip():
        raise ValueError(
            "base_path is not set in extraction_config.yaml\n"
            "Set it to the folder containing your FY* subfolders."
        )
    if not config.get("master_sheet_path", "").strip():
        raise ValueError(
            "master_sheet_path is not set in extraction_config.yaml\n"
            "Set it to the full path of the output Excel file."
        )
    if not config.get("processed_log_path", "").strip():
        raise ValueError(
            "processed_log_path is not set in extraction_config.yaml\n"
            "Set it to the full path of the processed log JSON file."
        )


def main():
    parser = argparse.ArgumentParser(description="NL-43 Config-Driven Extraction Pipeline")
    parser.add_argument("--force", action="store_true",
                        help="Re-extract all files regardless of processed log")
    parser.add_argument("--force-company", type=str, default=None,
                        help="Re-extract only this company key across all quarters")
    parser.add_argument("--quarter", nargs="+", default=None,
                        help="Override config quarters e.g. --quarter Q1 Q3")
    parser.add_argument("--dry-run", action="store_true",
                        help="Show what would be extracted without running")
    parser.add_argument("--skip-consolidated", action="store_true",
                        help="Ignore consolidated PDFs this run")
    parser.add_argument("--config", type=str, default=CONFIG_PATH,
                        help="Path to config file (default: extraction_config.yaml)")
    args = parser.parse_args()

    # --- Load and validate config ---
    try:
        config = load_config(args.config)
        validate_config(config)
    except (FileNotFoundError, ValueError) as e:
        logger.error(str(e))
        sys.exit(1)

    # Apply CLI overrides
    if args.quarter:
        config["quarters"] = args.quarter
    if args.skip_consolidated:
        config["consolidated_mode"] = "skip"

    # --- Scan for PDFs ---
    logger.info("Scanning folder structure...")
    try:
        scan_results, unrecognized = scan(config)
    except (FileNotFoundError, ValueError) as e:
        logger.error(str(e))
        sys.exit(1)

    if not scan_results:
        logger.info("No PDFs found matching the config. Check base_path and fiscal_years.")
        sys.exit(0)

    # --- Filter to unprocessed ---
    log_path = config["processed_log_path"]
    log_data = load_log(log_path)
    to_process = filter_unprocessed(
        scan_results, log_data,
        force=args.force,
        force_company=args.force_company,
    )

    if not to_process:
        logger.info("All files are up-to-date. Nothing to extract.")
        logger.info("Use --force to re-extract everything.")
        sys.exit(0)

    # --- Dry run ---
    if args.dry_run:
        from collections import defaultdict
        ind_map = defaultdict(set)
        con_map = defaultdict(set)
        
        for r in to_process:
            display_name = r.company_key.replace("_", " ").title()
            if r.source_type == "direct":
                ind_map[display_name].add(r.quarter)
            else:
                con_map[display_name].add(r.quarter)
        
        print("\n" + "="*60)
        print("DRY RUN — Summary of files to be processed:")
        print("="*60)
        
        print(f"\nINDIVIDUAL ({len(ind_map)} companies):")
        for name in sorted(ind_map.keys()):
            qs = ", ".join(sorted(list(ind_map[name])))
            print(f"  - {name:25s} ({qs})")
            
        print(f"\nCONSOLIDATED ({len(con_map)} companies):")
        for name in sorted(con_map.keys()):
            qs = ", ".join(sorted(list(con_map[name])))
            print(f"  - {name:25s} ({qs})")
            
        if unrecognized:
            print(f"\nUNDETECTED FILES ({len(unrecognized)}):")
            for path in sorted(unrecognized):
                print(f"  [!] {os.path.basename(path)}  ({path})")
        
        print("\n" + "="*60)
        sys.exit(0)

    # --- Extract ---
    master_path = config["master_sheet_path"]
    consolidated_mode = config.get("consolidated_mode", "dynamic")
    succeeded = 0
    failed = 0
    failed_files = []
    all_extractions = []
    all_validation_results = []

    for result in to_process:
        logger.info(
            f"Extracting [{result.source_type}] {result.company_key} "
            f"{result.fiscal_year} {result.quarter}"
        )

        pdf_to_parse = result.pdf_path
        temp_file = None

        try:
            # Handle consolidated PDFs
            if result.source_type == "consolidated" and consolidated_mode != "skip":
                page_overrides = config.get("nl43_page_overrides", {})
                override = page_overrides.get(result.company_key) or {}
                keywords = config.get("nl43_keywords", None)

                if "start" in override:
                    # Hard-pinned page range — bypass auto-detection entirely
                    start_0 = int(override["start"]) - 1
                    end_0 = int(override.get("end", override["start"])) - 1
                    pages = (start_0, end_0)
                    logger.info(
                        f"  Using hard page override for {result.company_key}: "
                        f"pages {start_0}-{end_0} (0-indexed)"
                    )
                else:
                    min_matches = int(override.get(
                        "min_matches",
                        config.get("nl43_keyword_min_matches", 2),
                    ))
                    pages = find_nl43_pages(result.pdf_path, keywords, min_matches)

                if pages is None:
                    logger.warning(
                        f"  Could not find NL-43 section in consolidated PDF, skipping: "
                        f"{result.pdf_path}"
                    )
                    failed += 1
                    failed_files.append((result.pdf_path, "NL-43 section not found"))
                    continue

                temp_file = extract_nl43_to_temp(result.pdf_path, pages[0], pages[1])
                if temp_file is None:
                    logger.warning(f"  Page extraction failed, skipping: {result.pdf_path}")
                    failed += 1
                    failed_files.append((result.pdf_path, "Page extraction failed"))
                    continue

                pdf_to_parse = temp_file

            # Parse
            extract = parse_pdf(
                pdf_to_parse,
                result.company_key,
                result.quarter,
                result.year_code,
            )

            # Validate
            validation_results = run_validations([extract])

            all_extractions.append(extract)
            all_validation_results.extend(validation_results)

            # Mark as processed
            mark_processed(log_data, result, 0)
            save_log(log_path, log_data)

            succeeded += 1
            logger.info(f"  Done: {result.company_key}")

        except Exception as e:
            logger.error(f"  Failed: {result.company_key} — {e}", exc_info=True)
            failed += 1
            failed_files.append((result.pdf_path, str(e)))

        finally:
            if temp_file and os.path.exists(temp_file):
                os.unlink(temp_file)

    # --- Write output workbook ---
    if all_extractions:
        stats = {
            "files_processed": succeeded + failed,
            "files_succeeded": succeeded,
            "files_failed": failed,
        }

        # 1. Master_Data + per-company verification sheets + _meta
        logger.info(f"Writing master workbook to {master_path}")
        save_workbook(all_extractions, master_path, stats=stats)

        # 2. Validation report CSV
        report_path = os.path.join(
            os.path.dirname(master_path), "NL43_validation_report.csv"
        )
        write_validation_report(all_validation_results, report_path)

        # 3. Validation_Summary tab
        write_validation_summary_sheet(report_path, master_path, force_company=args.force_company)

        # 4. Validation_Detail tab (FAILs + WARNs)
        write_validation_detail_sheet(report_path, master_path, force_company=args.force_company)

        logger.info(f"Validation report: {report_path}")

    # --- Summary ---
    logger.info("=" * 60)
    logger.info(f"Pipeline complete.")
    logger.info(f"  Succeeded: {succeeded}")
    logger.info(f"  Failed:    {failed}")
    if failed_files:
        logger.info("  Failed files:")
        for path, reason in failed_files:
            logger.info(f"    {path} — {reason}")


if __name__ == "__main__":
    main()
