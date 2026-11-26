"""Primary file processor.

Pipeline (see the SOP, section 4):
1. Select the 21 required columns (raw data is never touched).
2. Parse dates / numbers; rows with invalid data are excluded and reported.
3. SHORT CLOSED correction (those rows only):
       Order quantity = Delivery qty
       Net Value      = Delivery qty x Net Price     (Net Price is per 1 unit)
4. Remove rows whose corrected Net Value is exactly zero (negatives are kept).
5. Consolidate on the six key columns (sum Order quantity, sum Net Value).
6. Sort by Document Date, Sold-to party, Material.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from .validation import (
    ACTION_EXCLUDED, ACTION_KEPT, ProcessingResult, clean_key, combine_issues,
    make_issues, parse_dates, parse_numeric, prepare_source, sort_frame,
    unique_divisions,
)

PRIMARY_REQUIRED = [
    "Document Date", "Sold-to party", "Customer Name", "Division", "Material",
    "Description", "Base quantity", "Base UoM", "Order quantity", "Sales unit",
    "Base Price", "Discount", "Charge", "Freight", "VAT", "Net Price",
    "Net Discount", "Net Value", "Delivery qty", "Billed qty", "Order Status",
]

GROUP_KEYS = ["Document Date", "Sold-to party", "Customer Name", "Division",
              "Material", "Description"]
QTY_OUT = "(Sum of) Order quantity"
VAL_OUT = "(Sum of) Net Value"
OUTPUT_COLUMNS = GROUP_KEYS + [QTY_OUT, VAL_OUT]
SORT_BY = ["Document Date", "Sold-to party", "Material",
           # deterministic tie-breakers only (do not change the SOP order)
           "Customer Name", "Division", "Description"]
SHORT_CLOSED = "SHORT CLOSED"


def normalize_status(s: pd.Series) -> pd.Series:
    """Trim, collapse inner whitespace and upper-case the Order Status."""
    txt = s.astype(object).where(s.notna(), "").astype(str)
    return txt.str.replace(r"\s+", " ", regex=True).str.strip().str.upper()


def process_primary(raw: pd.DataFrame) -> ProcessingResult:
    source = prepare_source(raw, PRIMARY_REQUIRED)
    src, rows = source.src, source.rows
    issues: list = []
    warnings: list[str] = []
    infos: list[str] = []

    # ---- parse ---------------------------------------------------------- #
    dates, d_blank, d_invalid, date_notes = parse_dates(src["Document Date"])
    oq, oq_blank, oq_inv = parse_numeric(src["Order quantity"])
    nv, nv_blank, nv_inv = parse_numeric(src["Net Value"])
    dq, dq_blank, dq_inv = parse_numeric(src["Delivery qty"])
    npz, np_blank, np_inv = parse_numeric(src["Net Price"])
    infos.extend(date_notes)

    status = normalize_status(src["Order Status"])
    is_sc = status.eq(SHORT_CLOSED)
    non_sc = ~is_sc

    # ---- validation: exclude + report invalid rows ---------------------- #
    def _flag(mask, col, text):
        issues.append(make_issues(rows, src[col], mask, col, text, ACTION_EXCLUDED))

    _flag(d_blank, "Document Date", "Missing date")
    _flag(d_invalid, "Document Date", "Invalid date")
    _flag(oq_blank & non_sc, "Order quantity", "Missing number")
    _flag(oq_inv & non_sc, "Order quantity", "Invalid number")
    _flag(nv_blank & non_sc, "Net Value", "Missing number")
    _flag(nv_inv & non_sc, "Net Value", "Invalid number")
    _flag(dq_blank & is_sc, "Delivery qty",
          "SHORT CLOSED calculation issue: Delivery qty is missing")
    _flag(dq_inv & is_sc, "Delivery qty",
          "SHORT CLOSED calculation issue: Delivery qty is not a valid number")
    _flag(np_blank & is_sc, "Net Price",
          "SHORT CLOSED calculation issue: Net Price is missing")
    _flag(np_inv & is_sc, "Net Price",
          "SHORT CLOSED calculation issue: Net Price is not a valid number")

    date_bad = d_blank | d_invalid
    sc_calc_bad = is_sc & (dq_blank | dq_inv | np_blank | np_inv)
    exclude = (date_bad
               | (non_sc & (oq_blank | oq_inv | nv_blank | nv_inv))
               | sc_calc_bad)
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
        "Document Date": dates[ok],
        "Sold-to party": clean_key(src["Sold-to party"][ok]),
        "Customer Name": clean_key(src["Customer Name"][ok]),
        "Division": division_all[ok],
        "Material": clean_key(src["Material"][ok]),
        "Description": clean_key(src["Description"][ok]),
        "Order quantity": oq[ok],
        "Net Value": nv[ok],
    })

    # ---- SHORT CLOSED correction (SHORT CLOSED rows only) --------------- #
    sc_ok = is_sc & ok
    sc_orig_qty = float(oq[sc_ok].sum())
    sc_orig_val = float(nv[sc_ok].sum())
    sc_new_qty = float(dq[sc_ok].sum())
    sc_new_val = float((dq * npz)[sc_ok].sum())
    sc_mask_work = is_sc[ok]
    work.loc[sc_mask_work, "Order quantity"] = dq[ok][sc_mask_work]
    work.loc[sc_mask_work, "Net Value"] = (dq[ok] * npz[ok])[sc_mask_work]

    # Informational check of the "Net Price is per 1 unit" assumption on the
    # rows that are NOT recalculated. Never changes any number.
    chk = non_sc & ok & ~(np_blank | np_inv)
    if chk.any():
        diff = ((oq[chk] * npz[chk]) - nv[chk]).abs()
        tol = np.maximum(0.01, 1e-6 * nv[chk].abs())
        mismatch = int((diff > tol).sum())
        total = int(chk.sum())
        infos.append(
            f"Net Price check (information only): {total - mismatch:,} of {total:,} "
            "non-SHORT CLOSED rows have Net Value = Order quantity x Net Price"
            + (f"; {mismatch:,} rows differ (e.g. discounts/charges included in "
               "Net Value)." if mismatch else "."))

    # ---- zero Net Value filter ------------------------------------------ #
    before_zero = len(work)
    work = work[work["Net Value"] != 0]
    removed_zero = before_zero - len(work)

    # ---- data-quality notes (rows that are kept) ------------------------ #
    for col in ("Sold-to party", "Material", "Division"):
        issues.append(make_issues(work["_row"], work[col], work[col].isna(), col,
                                  f"Missing {col} (row kept)", ACTION_KEPT))
    neg_count = int((work["Net Value"] < 0).sum())
    if neg_count:
        infos.append(f"{neg_count:,} row(s) with negative Net Value were retained.")
    zero_qty = int((work["Order quantity"] == 0).sum())
    if zero_qty:
        infos.append(f"{zero_qty:,} row(s) with zero Order quantity (non-zero Net "
                     "Value) were retained.")
    dup = int(src.astype(str).duplicated().sum())
    if dup:
        infos.append(f"{dup:,} duplicate source row(s) detected. They are kept and "
                     "added together during consolidation.")
    counts = status.where(status != "", "(blank)").value_counts()
    infos.append("Order Status values: " +
                 ", ".join(f"{k} ({v:,})" for k, v in counts.items()))

    # ---- consolidation --------------------------------------------------- #
    rows_before_consol = len(work)
    if work.empty:
        grouped = pd.DataFrame(columns=OUTPUT_COLUMNS)
        warnings.append("No rows remain after filtering, so there is no processed "
                        "output to export.")
    else:
        grouped = (work.groupby(GROUP_KEYS, dropna=False, sort=False, as_index=False)
                   [["Order quantity", "Net Value"]].sum())
        grouped = grouped.rename(columns={"Order quantity": QTY_OUT,
                                          "Net Value": VAL_OUT})
        grouped = sort_frame(grouped[OUTPUT_COLUMNS], SORT_BY)
    rows_after_consol = len(grouped)

    if int(exclude.sum()):
        warnings.append(f"{int(exclude.sum()):,} row(s) with invalid data were "
                        "excluded from the processed sheet. See the validation report.")

    metrics = [
        ("Rows before processing", len(src)),
        ("Blank rows skipped", source.blank_rows),
        ("Rows excluded (invalid data)", int(exclude.sum())),
        ("Rows removed (zero Net Value)", removed_zero),
        ("Rows before consolidation", rows_before_consol),
        ("Rows after consolidation", rows_after_consol),
        ("Rows consolidated", rows_before_consol - rows_after_consol),
        ("Final row count", rows_after_consol),
        ("Total Order quantity", float(grouped[QTY_OUT].sum()) if rows_after_consol else 0.0),
        ("Total Net Value", float(grouped[VAL_OUT].sum()) if rows_after_consol else 0.0),
        ("Earliest date", date_min),
        ("Latest date", date_max),
    ]
    specific = [
        ("SHORT CLOSED rows detected", int(is_sc.sum())),
        ("SHORT CLOSED rows recalculated", int(sc_ok.sum())),
        ("SHORT CLOSED Order qty - original", sc_orig_qty),
        ("SHORT CLOSED Order qty - corrected", sc_new_qty),
        ("SHORT CLOSED Net Value - original", sc_orig_val),
        ("SHORT CLOSED Net Value - recalculated", sc_new_val),
        ("Rows removed due to zero Net Value", removed_zero),
        ("Invalid SHORT CLOSED calculations", int(sc_calc_bad.sum())),
    ]

    return ProcessingResult(
        kind="primary", processed=grouped, issues=combine_issues(issues),
        metrics=metrics, specific_metrics=specific, warnings=warnings,
        infos=infos, divisions=divisions, date_min=date_min, date_max=date_max,
        rows_excluded=int(exclude.sum()),
    )
