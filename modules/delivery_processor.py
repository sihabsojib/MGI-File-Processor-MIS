"""Delivery file processor.

Pipeline (see the SOP, section 5):
1. Select the 8 required columns (raw data is never touched).
2. Parse dates / numbers; rows with invalid data are excluded and reported.
3. Remove rows whose Item Value is exactly zero (negatives are kept; the
   Billed Quantity is never used for filtering).
4. Consolidate on the six key columns (sum Billed Quantity, sum Item Value).
5. Sort by Billing Date, Customer Code, Material.
"""
from __future__ import annotations

import pandas as pd

from .validation import (
    ACTION_EXCLUDED, ACTION_KEPT, ProcessingResult, clean_key, combine_issues,
    make_issues, parse_dates, parse_numeric, prepare_source, sort_frame,
    unique_divisions,
)

DELIVERY_REQUIRED = [
    "Billing Date", "Customer Code", "Customer Name", "Division", "Material",
    "Material Description", "Billed Quantity", "Item Value",
]
GROUP_KEYS = ["Billing Date", "Customer Code", "Customer Name", "Division",
              "Material", "Material Description"]
QTY_OUT = "Billed Quantity"
VAL_OUT = "Item Value"
OUTPUT_COLUMNS = GROUP_KEYS + [QTY_OUT, VAL_OUT]
SORT_BY = ["Billing Date", "Customer Code", "Material",
           # deterministic tie-breakers only (do not change the SOP order)
           "Customer Name", "Division", "Material Description"]


def process_delivery(raw: pd.DataFrame) -> ProcessingResult:
    source = prepare_source(raw, DELIVERY_REQUIRED)
    src, rows = source.src, source.rows
    issues: list = []
    warnings: list[str] = []
    infos: list[str] = []

    # ---- parse ---------------------------------------------------------- #
    dates, d_blank, d_invalid, date_notes = parse_dates(src["Billing Date"])
    qty, q_blank, q_inv = parse_numeric(src["Billed Quantity"])
    val, v_blank, v_inv = parse_numeric(src["Item Value"])
    infos.extend(date_notes)

    # ---- validation: exclude + report invalid rows ---------------------- #
    def _flag(mask, col, text):
        issues.append(make_issues(rows, src[col], mask, col, text, ACTION_EXCLUDED))

    _flag(d_blank, "Billing Date", "Missing date")
    _flag(d_invalid, "Billing Date", "Invalid date")
    _flag(q_blank, "Billed Quantity", "Missing number")
    _flag(q_inv, "Billed Quantity", "Invalid number")
    _flag(v_blank, "Item Value", "Missing number")
    _flag(v_inv, "Item Value", "Invalid number")

    date_bad = d_blank | d_invalid
    exclude = date_bad | q_blank | q_inv | v_blank | v_inv
    ok = ~exclude

    # ---- date coverage and divisions (all loaded rows, before filtering) - #
    valid_dates = dates[~date_bad]
    date_min = valid_dates.min().date() if not valid_dates.empty else None
    date_max = valid_dates.max().date() if not valid_dates.empty else None
    division_all = clean_key(src["Division"])
    divisions = unique_divisions(division_all)
    if not divisions:
        warnings.append("Division is missing in the source data. The file name "
                        "will use 'Unknown-Division'.")
    elif len(divisions) > 1:
        warnings.append("Multiple Divisions found (" + ", ".join(divisions) +
                        "). The file name will use 'Multi-Division'.")

    # ---- working frame (only valid rows) -------------------------------- #
    work = pd.DataFrame({
        "_row": rows[ok],
        "Billing Date": dates[ok],
        "Customer Code": clean_key(src["Customer Code"][ok]),
        "Customer Name": clean_key(src["Customer Name"][ok]),
        "Division": division_all[ok],
        "Material": clean_key(src["Material"][ok]),
        "Material Description": clean_key(src["Material Description"][ok]),
        QTY_OUT: qty[ok],
        VAL_OUT: val[ok],
    })

    # ---- zero Item Value filter ------------------------------------------ #
    before_zero = len(work)
    work = work[work[VAL_OUT] != 0]
    removed_zero = before_zero - len(work)

    # ---- data-quality notes (rows that are kept) ------------------------ #
    for col in ("Customer Code", "Material", "Division"):
        issues.append(make_issues(work["_row"], work[col], work[col].isna(), col,
                                  f"Missing {col} (row kept)", ACTION_KEPT))
    neg_count = int((work[VAL_OUT] < 0).sum())
    if neg_count:
        infos.append(f"{neg_count:,} row(s) with negative Item Value were retained.")
    zero_qty = int((work[QTY_OUT] == 0).sum())
    if zero_qty:
        infos.append(f"{zero_qty:,} row(s) with zero Billed Quantity (non-zero Item "
                     "Value) were retained.")
    dup = int(src.astype(str).duplicated().sum())
    if dup:
        infos.append(f"{dup:,} duplicate source row(s) detected. They are kept and "
                     "added together during consolidation.")

    # ---- consolidation --------------------------------------------------- #
    rows_before_consol = len(work)
    if work.empty:
        grouped = pd.DataFrame(columns=OUTPUT_COLUMNS)
        warnings.append("No rows remain after filtering, so there is no processed "
                        "output to export.")
    else:
        grouped = (work.groupby(GROUP_KEYS, dropna=False, sort=False, as_index=False)
                   [[QTY_OUT, VAL_OUT]].sum())
        grouped = sort_frame(grouped[OUTPUT_COLUMNS], SORT_BY)
    rows_after_consol = len(grouped)

    if int(exclude.sum()):
        warnings.append(f"{int(exclude.sum()):,} row(s) with invalid data were "
                        "excluded from the processed sheet. See the validation report.")

    total_qty = float(grouped[QTY_OUT].sum()) if rows_after_consol else 0.0
    total_val = float(grouped[VAL_OUT].sum()) if rows_after_consol else 0.0
    metrics = [
        ("Rows before processing", len(src)),
        ("Blank rows skipped", source.blank_rows),
        ("Rows excluded (invalid data)", int(exclude.sum())),
        ("Rows removed (zero Item Value)", removed_zero),
        ("Rows before consolidation", rows_before_consol),
        ("Rows after consolidation", rows_after_consol),
        ("Rows consolidated", rows_before_consol - rows_after_consol),
        ("Final row count", rows_after_consol),
        ("Total Billed Quantity", total_qty),
        ("Total Item Value", total_val),
        ("Earliest date", date_min),
        ("Latest date", date_max),
    ]
    specific = [
        ("Rows removed due to zero Item Value", removed_zero),
        ("Total Billed Quantity", total_qty),
        ("Total Item Value", total_val),
    ]

    return ProcessingResult(
        kind="delivery", processed=grouped, issues=combine_issues(issues),
        metrics=metrics, specific_metrics=specific, warnings=warnings,
        infos=infos, divisions=divisions, date_min=date_min, date_max=date_max,
        rows_excluded=int(exclude.sum()),
    )
