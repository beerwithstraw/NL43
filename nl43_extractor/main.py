"""
CLI entry point for the NL43 Batch Premium Extractor.

Source: approach document Section 6
"""

import sys
import logging
from pathlib import Path

import click
from rich.console import Console
from rich.table import Table

from output.manifest import generate_manifest
from output.organiser import organise_all
from config.settings import DEFAULT_INPUT_DIR, DEFAULT_OUTPUT_DIR

# NOTE: For config-driven path-based extraction, use pipeline.py instead:
#   python3 pipeline.py
# main.py continues to work for the existing inputs/ folder workflow.

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-5s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
# We will add file logging appropriately once logs directory exists
# For CLI we mainly want rich output, but we keep basic logging for now.

console = Console()

@click.group()
def cli():
    """NL43 Batch Premium Extractor CLI."""
    pass


# ---------------------------------------------------------------------------
# `scan` command: Phase 1 (Dry Run / Manifest Generation)
# ---------------------------------------------------------------------------

@cli.command()
@click.option("--input", "-i", "input_dir", default=DEFAULT_INPUT_DIR,
              help="Directory containing source PDFs", type=click.Path(exists=True))
@click.option("--manifest", "-m", "manifest_csv", default="outputs/manifest.csv",
              help="Path to output manifest CSV", type=click.Path())
def scan(input_dir, manifest_csv):
    """
    Scan PDFs and generate a manifest CSV for human review.
    """
    console.print(f"[bold blue]Scanning PDFs in:[/bold blue] {input_dir}")
    
    try:
        count = generate_manifest(input_dir, manifest_csv)
        
        console.print(f"\n[bold green]Scan complete![/bold green]")
        console.print(f"Processed {count} files.")
        console.print(f"Manifest written to: [bold]{manifest_csv}[/bold]")
        console.print("\n[yellow]Next step: Review and edit the manifest in Excel.[/yellow]")
        console.print("Change 'action' to 'skip' for files you want to exclude.")
        console.print("Correct any wrongly detected metadata, then run the 'extract' command.\n")
        
        # Print a preview table
        try:
            import csv
            with open(manifest_csv, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                
                if rows:
                    table = Table(title="Manifest Preview")
                    # We just show a few key columns to keep it readable
                    table.add_column("Filename", style="cyan", no_wrap=True)
                    table.add_column("Company", style="magenta")
                    table.add_column("Qtr", style="green")
                    table.add_column("Year", style="green")
                    table.add_column("Confidence", style="yellow")
                    table.add_column("Action", style="white")
                    
                    # Show max 10 rows
                    for row in rows[:10]:
                        table.add_row(
                            row["filename"][:25] + ("..." if len(row["filename"]) > 25 else ""),
                            row["detected_company"],
                            row["detected_quarter"],
                            row["detected_year"],
                            row["confidence"],
                            row["action"]
                        )
                    console.print(table)
                    if len(rows) > 10:
                        console.print(f"... and {len(rows) - 10} more rows.")
        except Exception as e:
            logging.debug(f"Preview table build failed: {e}")
            
    except Exception as e:
        console.print(f"[bold red]Error generating manifest:[/bold red] {e}")
        sys.exit(1)

# ---------------------------------------------------------------------------
# `extract` command: Phase 2 (Actual Extraction)
# ---------------------------------------------------------------------------

@cli.command()
@click.option("--input", "-i", "input_dir", default=DEFAULT_INPUT_DIR,
              help="Directory containing source PDFs", type=click.Path(exists=True))
@click.option("--manifest", "-m", "manifest_csv", default="outputs/manifest.csv",
              help="Path to input manifest CSV", type=click.Path())
@click.option("--output", "-o", "output_dir", default=DEFAULT_OUTPUT_DIR,
              help="Directory for output Excel and organised PDFs", type=click.Path())
@click.option("--auto", is_flag=True, help="Skip manifest, run fully automated")
@click.option("--year-selection", "year_selection", default="both",
              type=click.Choice(["current", "previous", "both"]),
              help="Which year to extract: current, previous, or both (default: both)")
def extract(input_dir, manifest_csv, output_dir, auto, year_selection):
    """
    Run extraction using manifest CSV (Phase 2).
    """
    console.print(f"[bold blue]Starting extraction run...[/bold blue]")
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    try:
        from output.manifest import read_manifest
        from extractor.parser import parse_pdf
        from output.excel_writer import save_workbook, write_validation_summary_sheet, write_validation_detail_sheet
        from validation.checks import run_validations, write_validation_report, build_validation_summary_table
        
        if auto:
            from output.manifest import generate_manifest
            manifest_csv = output_path / "manifest_auto.csv"
            generate_manifest(input_dir, manifest_csv)

        rows = read_manifest(manifest_csv)
        to_process = [r for r in rows if r["action"] == "proceed"]
        
        if not to_process:
            console.print("[yellow]No files marked for 'proceed' in manifest. Nothing to do.[/yellow]")
            return
            
        console.print(f"Found {len(to_process)} files to process.")
        
        # 2. Extract
        extractions = []
        stats = {"files_processed": 0, "files_succeeded": 0, "files_failed": 0, "files_uncategorised": 0}
        
        with console.status("[bold green]Extracting data from PDFs...") as status:
            for row in to_process:
                pdf_name = row["filename"]
                ckey = row["detected_company"]
                qtr = row["detected_quarter"]
                year = row["detected_year"]
                
                pdf_path = input_path / pdf_name
                if not pdf_path.exists():
                    console.print(f"[red]File not found: {pdf_name}[/red]")
                    stats["files_failed"] += 1
                    continue
                    
                try:
                    stats["files_processed"] += 1
                    extract = parse_pdf(str(pdf_path), ckey, qtr, year)
                    extractions.append(extract)
                    stats["files_succeeded"] += 1
                    console.print(f"  [green]\u2713[/green] {pdf_name}: [dim]{ckey}[/dim]")
                except Exception as e:
                    console.print(f"  [red]\u2717[/red] {pdf_name}: {e}")
                    stats["files_failed"] += 1
        
        if not extractions:
            console.print("[bold red]No data successfully extracted.[/bold red]")
            return
            
        # 3. Validate
        console.print("\n[bold blue]Running validation checks...[/bold blue]")
        val_results = run_validations(extractions)
        report_path = output_path / "validation_report.csv"
        write_validation_report(val_results, str(report_path))
        
        # 4. Save Excel
        excel_path = output_path / "NL43_Master.xlsx"
        save_workbook(extractions, str(excel_path), stats=stats, year_selection=year_selection)
        
        # 5. Append Validation Sheets
        write_validation_summary_sheet(str(report_path), str(excel_path))
        write_validation_detail_sheet(str(report_path), str(excel_path))
        
        # 6. Organise PDFs
        console.print("[bold blue]Organising PDFs into folder structure...[/bold blue]")
        # Note: organise_all re-detects, but satisfies the requirement
        organise_all(input_dir, output_dir)
        
        # Final Summary
        console.print(f"\n[bold green]Extraction complete![/bold green]")
        console.print(f"Processed: {stats['files_processed']}")
        console.print(f"Succeeded: {stats['files_succeeded']}")
        console.print(f"Failed:    {stats['files_failed']}")
        console.print(f"\nExcel Output: [bold]{excel_path}[/bold]")
        console.print(f"Validation Report: [bold]{report_path}[/bold]\n")
        
        console.print(build_validation_summary_table(val_results))

    except Exception as e:
        console.print(f"[bold red]Extraction failed:[/bold red] {e}")
        import traceback
        logging.error(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    cli()
