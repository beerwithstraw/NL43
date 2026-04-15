"""
Validation Checks for NL-43 (Rural & Social Obligations).

Check families:
  1. COMPLETENESS — mandatory LOBs (fire, health, motor_od, motor_tp) have data
  2. TOTAL_SUM    — sum of component LOBs ≈ grand_total (Rural and Social separately)
"""

import csv
import logging
from typing import Dict, List, Optional
from dataclasses import dataclass, asdict

from extractor.models import CompanyExtract, PeriodData
from config.lob_registry import LOB_ORDER
from config.row_registry import ROW_ORDER
from config.company_registry import COMPLETENESS_IGNORE

logger = logging.getLogger(__name__)

TOLERANCE = 2.0
MANDATORY_LOBS = {"fire", "health", "motor_od", "motor_tp"}

# LOBs that sum to grand_total
_COMPONENT_LOBS = [
    "fire", "marine_cargo", "marine_hull", "motor_od", "motor_tp",
    "health", "personal_accident", "travel_insurance", "wc_el",
    "public_product_liability", "engineering", "aviation",
    "crop_insurance", "other_segments", "total_miscellaneous",
]


@dataclass
class ValidationResult:
    company: str
    quarter: str
    year: str
    lob: str
    period: str   # "rural" or "social"
    check_name: str
    status: str   # PASS, WARN, FAIL
    expected: Optional[float]
    actual: Optional[float]
    delta: Optional[float]
    note: str


def run_validations(extractions: List[CompanyExtract]) -> List[ValidationResult]:
    """Run all NL-43 validation checks on the provided extractions."""
    results: List[ValidationResult] = []

    for exc in extractions:
        period_data = exc.current_year
        if not period_data:
            continue

        # Completeness
        results.extend(_check_completeness(exc, period_data))

        # Total sum for each segment and each metric
        for segment in ("rural", "social"):
            for metric in ROW_ORDER:
                r = _check_total_sum(exc, period_data, metric, segment)
                if r:
                    results.append(r)

    return results


def _get_val(lob_data: Dict, metric: str, segment: str) -> Optional[float]:
    v = lob_data.get(metric, {}).get(segment)
    return float(v) if v is not None else None


def _check_completeness(exc: CompanyExtract, period_data: PeriodData) -> List[ValidationResult]:
    """COMPLETENESS: mandatory LOBs must have at least one non-None value."""
    results = []
    ignore = COMPLETENESS_IGNORE.get(exc.company_key, set())
    for lob in LOB_ORDER:
        if lob in ignore or lob == "grand_total":
            continue
        lob_data = period_data.data.get(lob, {})
        has_data = any(
            any(v is not None for v in cell.values())
            for cell in lob_data.values()
        )
        if not has_data:
            status = "FAIL" if lob in MANDATORY_LOBS else "WARN"
            results.append(ValidationResult(
                exc.company_name, exc.quarter, exc.year, lob, "current",
                "COMPLETENESS", status,
                expected=None, actual=None, delta=None,
                note=f"LOB '{lob}' is missing",
            ))
    return results


def _check_total_sum(
    exc: CompanyExtract,
    period_data: PeriodData,
    metric: str,
    segment: str,
) -> Optional[ValidationResult]:
    """TOTAL_SUM: sum of component LOBs ≈ grand_total for the given metric+segment."""
    total_data = period_data.data.get("grand_total", {})
    total_val = _get_val(total_data, metric, segment)
    if total_val is None:
        return None

    component_sum = 0.0
    valid_count = 0
    for lob in _COMPONENT_LOBS:
        val = _get_val(period_data.data.get(lob, {}), metric, segment)
        if val is not None:
            component_sum += val
            valid_count += 1

    if valid_count == 0:
        return None

    delta = abs(total_val - component_sum)
    status = "PASS" if delta <= TOLERANCE else "FAIL"
    return ValidationResult(
        exc.company_name, exc.quarter, exc.year, "grand_total", segment,
        f"TOTAL_SUM_{metric.upper()}", status,
        expected=component_sum, actual=total_val, delta=delta, note="",
    )


def build_validation_summary_table(results: List[ValidationResult]):
    from rich.table import Table
    counts: Dict[str, int] = {"PASS": 0, "SKIP": 0, "WARN": 0, "FAIL": 0}
    for r in results:
        counts[r.status] = counts.get(r.status, 0) + 1
    t = Table(title="Validation Summary")
    t.add_column("Status", style="bold")
    t.add_column("Count", justify="right")
    t.add_row("[green]PASS[/green]", str(counts["PASS"]))
    t.add_row("[blue]SKIP[/blue]", str(counts["SKIP"]))
    t.add_row("[yellow]WARN[/yellow]", str(counts["WARN"]))
    t.add_row("[red]FAIL[/red]", str(counts["FAIL"]))
    return t


def write_validation_report(results: List[ValidationResult], output_path: str):
    with open(output_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "company", "quarter", "year", "lob", "period",
            "check_name", "status", "expected", "actual", "delta", "note",
        ])
        writer.writeheader()
        for r in results:
            writer.writerow(asdict(r))
    logger.info(f"Validation report saved to {output_path}")
