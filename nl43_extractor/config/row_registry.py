"""
Metric Registry for NL-43 (Rural & Social Obligations).

NL-43 has 3 metrics per LOB, each split by segment (Rural / Social):
  - No. of Policies Issued
  - Premium Collected
  - Sum Assured

The nested dict in PeriodData uses "rural" and "social" as the period keys
(analogous to "qtr"/"ytd" in other extractors).
"""

ROW_ORDER = [
    "policies_issued",
    "premium_collected",
    "sum_assured",
]

ROW_DISPLAY_NAMES = {
    "policies_issued":   "No. of Policies Issued",
    "premium_collected": "Premium Collected",
    "sum_assured":       "Sum Assured",
}

# The two segment keys used as the period-axis in data[lob][metric]
SEGMENT_KEYS = ["rural", "social"]
