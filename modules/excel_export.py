"""Excel export: two-sheet workbooks, date-only cells and file naming.

Each workbook has exactly two sheets:
    * ``Primary`` / ``Delivery`` - processed output (dates are genuine Excel
      dates, date only, number format ``mm/dd/yy``)
    * ``Raw Data``               - the original upload, written as-is

The workbook is written with XlsxWriter in ``constant_memory`` mode so large
raw files stay practical.
"""
from __future__ import annotations

import io
import re
from datetime import date, datetime
from typing import Optional

import numpy as np
import pandas as pd
import xlsxwriter

from .validation import ExportError, ProcessingResult

PROCESSED_SHEET = {"primary": "Primary", "delivery": "Delivery"}
RAW_SHEET = "Raw Data"
DATE_FORMAT = "mm/dd/yy"
NUMBER_FORMAT = "#,##0.00##"   # display only - stored values are never rounded
MAX_ROWS = 1_048_576           # Excel worksheet limit (including header)
MAX_COLS = 16_384

# Columns holding identifiers/text (written as text, so no '.0' and no lost
# leading zeros) and the numeric measure columns, per module.
DATE_COLUMNS = {"primary": "Document Date", "delivery": "Billing Date"}
NUMERIC_COLUMNS = {
    "primary": ["(Sum of) Order quantity", "(Sum of) Net Value"],
    "delivery": ["Billed Quantity", "Item Value"],
}
_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# --------------------------------------------------------------------------- #
# File naming
# --------------------------------------------------------------------------- #


def sanitize_division(text: str) -> str:
    text = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "", str(text))
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:40]


def _short_date(d: date) -> str:
    return f"{_MONTHS[d.month - 1]} {d.day}"


def build_filename(kind: str, divisions: list[str], date_min: Optional[date],
                   date_max: Optional[date], suffix: str = "") -> str:
    """e.g. 'Primary PL-K (Nov 1 to Nov 25).xlsx' / 'Delivery PL-K (Nov 25).xlsx'."""
    label = PROCESSED_SHEET[kind]
    if len(divisions) == 1:
        division = sanitize_division(divisions[0]) or "Unknown-Division"
    elif len(divisions) > 1:
        division = "Multi-Division"
    else:
        division = "Unknown-Division"

    if date_min is None or date_max is None:
        span = "no valid dates"
    elif date_min == date_max:
        span = _short_date(date_min)
    else:
        span = f"{_short_date(date_min)} to {_short_date(date_max)}"
    return f"{label} {division} ({span}){suffix}.xlsx"


def result_filename(result: ProcessingResult) -> str:
    return build_filename(result.kind, result.divisions, result.date_min, result.date_max)


def report_filename(result: ProcessingResult) -> str:
    return build_filename(result.kind, result.divisions, result.date_min,
                          result.date_max, suffix=" - Validation Report")


# --------------------------------------------------------------------------- #
# Limits
# --------------------------------------------------------------------------- #


def check_excel_limits(raw: pd.DataFrame, processed: pd.DataFrame) -> None:
    for name, df in ((RAW_SHEET, raw), ("processed", processed)):
        if len(df) + 1 > MAX_ROWS:
            raise ExportError(
                f"The {name} data has {len(df):,} rows. Excel allows at most "
                f"{MAX_ROWS - 1:,} data rows per worksheet (plus the header). "
                "Please split the source file and process it in parts.")
        if df.shape[1] > MAX_COLS:
            raise ExportError(f"The {name} data has more than {MAX_COLS:,} columns, "
                              "which Excel cannot hold.")


# --------------------------------------------------------------------------- #
# Cell writers
# --------------------------------------------------------------------------- #


def _is_blank_value(v) -> bool:
    return v is None or v is pd.NaT or (isinstance(v, float) and v != v)


def _write_raw_cell(ws, r, c, v, fmts):
    """Write one raw value as faithfully as possible (no conversions)."""
    if _is_blank_value(v):
        return
    if isinstance(v, str):
        if v != "":
            ws.write_string(r, c, v[:32767])          # never interpreted as a formula
    elif isinstance(v, (bool, np.bool_)):
        ws.write_boolean(r, c, bool(v))
    elif isinstance(v, (int, np.integer)):
        if abs(int(v)) >= 10 ** 15:                   # beyond Excel's 15-digit precision
            ws.write_string(r, c, str(int(v)))
        else:
            ws.write_number(r, c, int(v))
    elif isinstance(v, (float, np.floating)):
        if np.isinf(v):
            ws.write_string(r, c, str(v))
        else:
            ws.write_number(r, c, float(v))
    elif isinstance(v, (datetime, date)):
        try:
            midnight = (not isinstance(v, datetime)) or (
                v.hour == v.minute == v.second == v.microsecond == 0)
            ws.write_datetime(r, c, v, fmts["raw_date"] if midnight else fmts["raw_datetime"])
        except (OverflowError, ValueError, TypeError):
            ws.write_string(r, c, str(v))
    else:
        ws.write_string(r, c, str(v)[:32767])


def _col_width(header: str, sample: pd.Series, minimum=9, maximum=50) -> float:
    longest = max([len(str(header))] + [len(str(x)) for x in sample.head(200).tolist()
                                         if not _is_blank_value(x)] + [0])
    return float(min(max(longest + 2, minimum), maximum))


def _write_raw_sheet(wb, fmts, raw: pd.DataFrame) -> None:
    ws = wb.add_worksheet(RAW_SHEET)
    ncols = raw.shape[1]
    for c, name in enumerate(raw.columns):
        ws.write_string(0, c, str(name), fmts["header"])
        ws.set_column(c, c, _col_width(name, raw.iloc[:, c]))
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(raw), 1), max(ncols - 1, 0))

    chunk = 5000
    for start in range(0, len(raw), chunk):
        part = raw.iloc[start:start + chunk]
        cols = [part.iloc[:, i].tolist() for i in range(ncols)]
        for offset in range(len(part)):
            r = start + offset + 1
            for c in range(ncols):
                _write_raw_cell(ws, r, c, cols[c][offset], fmts)


def _write_processed_sheet(wb, fmts, kind: str, df: pd.DataFrame) -> None:
    ws = wb.add_worksheet(PROCESSED_SHEET[kind])
    date_col = DATE_COLUMNS[kind]
    numeric_cols = NUMERIC_COLUMNS[kind]
    columns = list(df.columns)

    for c, name in enumerate(columns):
        ws.write_string(0, c, name, fmts["header"])
        if name == date_col:
            width = 12
        elif name in numeric_cols:
            width = 18
        else:
            width = _col_width(name, df[name], minimum=12, maximum=45)
        ws.set_column(c, c, width)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(df), 1), max(len(columns) - 1, 0))

    data = [df[name].tolist() for name in columns]
    for i in range(len(df)):
        r = i + 1
        for c, name in enumerate(columns):
            v = data[c][i]
            if _is_blank_value(v):
                continue
            if name == date_col:
                # date() has no time part -> stored as a whole-number Excel date
                ws.write_datetime(r, c, pd.Timestamp(v).date(), fmts["date"])
            elif name in numeric_cols:
                ws.write_number(r, c, float(v), fmts["number"])
            else:
                ws.write_string(r, c, str(v))


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #


def build_workbook(kind: str, processed: pd.DataFrame, raw: pd.DataFrame) -> bytes:
    """Create the two-sheet workbook and return it as bytes."""
    if kind not in PROCESSED_SHEET:
        raise ExportError(f"Unknown module '{kind}'.")
    check_excel_limits(raw, processed)

    buffer = io.BytesIO()
    wb = xlsxwriter.Workbook(buffer, {"constant_memory": True})
    fmts = {
        "header": wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1,
                                 "valign": "vcenter"}),
        "date": wb.add_format({"num_format": DATE_FORMAT}),
        "number": wb.add_format({"num_format": NUMBER_FORMAT}),
        "raw_date": wb.add_format({"num_format": "yyyy-mm-dd"}),
        "raw_datetime": wb.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"}),
    }
    _write_processed_sheet(wb, fmts, kind, processed)
    _write_raw_sheet(wb, fmts, raw)
    wb.close()
    return buffer.getvalue()


def build_validation_report(result: ProcessingResult) -> bytes:
    """Excel report listing every flagged row with its original row number."""
    issues = result.issues
    limit = MAX_ROWS - 1
    truncated = len(issues) > limit
    issues = issues.head(limit)

    summary = (issues.groupby(["Action", "Issue", "Column"]).size()
               .reset_index(name="Rows") if not issues.empty
               else pd.DataFrame(columns=["Action", "Issue", "Column", "Rows"]))

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        issues.to_excel(writer, sheet_name="Issues", index=False)
        header = writer.book.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        for name, df, widths in (("Summary", summary, [12, 60, 22, 10]),
                                 ("Issues", issues, [12, 22, 28, 60, 12])):
            ws = writer.sheets[name]
            for c, col in enumerate(df.columns):
                ws.write(0, c, col, header)
                ws.set_column(c, c, widths[c])
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
        if truncated:
            writer.sheets["Summary"].write(len(summary) + 2, 0,
                                           f"Only the first {limit:,} issues are listed.")
    return buffer.getvalue()
