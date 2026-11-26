"""Validation, parsing and shared helpers for the MIS File Processor.

Everything that is common to the Primary and Delivery processors lives here:
required-column matching, date / numeric parsing, identifier cleaning, the
issue (validation report) builder and the ``ProcessingResult`` container.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------- #
# Exceptions (all user-facing: the UI shows the message, never a traceback)
# --------------------------------------------------------------------------- #


class MISError(Exception):
    """Base class for errors that are safe to show to the user."""


class MissingColumnsError(MISError):
    def __init__(self, missing: list[str]):
        self.missing = list(missing)
        super().__init__("Missing required column(s): " + ", ".join(self.missing))


class EmptyDataError(MISError):
    """The file / worksheet has no usable data."""


class ExportError(MISError):
    """The processed data cannot be exported to Excel."""


# --------------------------------------------------------------------------- #
# Result container
# --------------------------------------------------------------------------- #

ISSUE_COLUMNS = ["Source Row", "Column", "Value", "Issue", "Action"]
ACTION_EXCLUDED = "Excluded"
ACTION_KEPT = "Kept"


@dataclass
class ProcessingResult:
    kind: str                                   # "primary" | "delivery"
    processed: pd.DataFrame                     # final output sheet
    issues: pd.DataFrame                        # validation report rows
    metrics: list = field(default_factory=list)           # [(label, value)]
    specific_metrics: list = field(default_factory=list)  # [(label, value)]
    warnings: list = field(default_factory=list)
    infos: list = field(default_factory=list)
    divisions: list = field(default_factory=list)
    date_min: Optional[date] = None
    date_max: Optional[date] = None
    rows_excluded: int = 0


# --------------------------------------------------------------------------- #
# Column matching
# --------------------------------------------------------------------------- #


def clean_header(name) -> str:
    """Header text used for matching: leading/trailing spaces ignored."""
    return str(name).strip()


@dataclass
class ColumnValidation:
    positions: dict          # required column -> position in the raw frame
    missing: list            # required columns that were not found

    @property
    def ok(self) -> bool:
        return not self.missing


def validate_columns(raw: pd.DataFrame, required: list[str]) -> ColumnValidation:
    """Match required columns (exact text, ignoring surrounding spaces)."""
    first_pos: dict[str, int] = {}
    for pos, col in enumerate(raw.columns):
        first_pos.setdefault(clean_header(col), pos)
    positions = {c: first_pos[c] for c in required if c in first_pos}
    missing = [c for c in required if c not in first_pos]
    return ColumnValidation(positions=positions, missing=missing)


# --------------------------------------------------------------------------- #
# Blank handling
# --------------------------------------------------------------------------- #


def blank_mask(s: pd.Series) -> pd.Series:
    """True where a value is missing or only whitespace."""
    isna = s.isna()
    if s.dtype == object:
        return isna | s.map(lambda v: isinstance(v, str) and v.strip() == "")
    return isna


def blank_rows_mask(raw: pd.DataFrame) -> pd.Series:
    """True for rows in which every cell is blank."""
    result = pd.Series(True, index=raw.index)
    for i in range(raw.shape[1]):
        result &= blank_mask(raw.iloc[:, i])
        if not result.any():
            break
    return result


# --------------------------------------------------------------------------- #
# Numeric parsing
# --------------------------------------------------------------------------- #

_THOUSANDS = r"^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$"


def _parse_numeric_text(strs: pd.Series) -> pd.Series:
    t = strs.str.replace(r"[\s\u00a0]", "", regex=True)
    t = t.str.replace(r"^\((.+)\)$", r"-\1", regex=True)      # (100) -> -100
    t = t.str.replace(r"^(\d[\d.,]*)-$", r"-\1", regex=True)  # 100-  -> -100 (SAP style)
    thousands = t.str.match(_THOUSANDS)
    t = t.mask(thousands, t.str.replace(",", "", regex=False))
    return pd.to_numeric(t, errors="coerce")


def parse_numeric(s: pd.Series):
    """Parse a column to float64 without rounding.

    Returns ``(values, blank, invalid)``. ``blank`` = empty cell, ``invalid`` =
    something is there but it is not a valid number. Blanks are never turned
    into zero.
    """
    if not pd.api.types.is_numeric_dtype(s) or pd.api.types.is_bool_dtype(s):
        s = s.astype(object)
    blank = blank_mask(s)
    out = pd.Series(np.nan, index=s.index, dtype="float64")

    if s.dtype == object:
        is_str = s.map(lambda v: isinstance(v, str))
        is_bool = s.map(lambda v: isinstance(v, (bool, np.bool_)))
        other = ~is_str & ~is_bool & ~blank
        if other.any():
            try:
                out[other] = pd.to_numeric(s[other], errors="coerce").astype("float64")
            except (TypeError, ValueError):
                pass  # stays NaN -> reported as invalid
        text = is_str & ~blank
        if text.any():
            out[text] = _parse_numeric_text(s[text]).astype("float64")
    else:
        out = pd.to_numeric(s, errors="coerce").astype("float64")

    out[np.isinf(out)] = np.nan
    invalid = out.isna() & ~blank
    return out, blank, invalid


# --------------------------------------------------------------------------- #
# Date parsing (date only, no time component)
# --------------------------------------------------------------------------- #

_TIME_SUFFIX = re.compile(r"[ T]\d{1,2}:\d{2}(:\d{2})?(\.\d+)?\s*([AaPp]\.?[Mm]\.?)?$")
# Day-first formats are listed before month-first ones, so an ambiguous text
# column (all values <= 12) is read day/month/year. Month-first is only used
# when day/month cannot parse the column (e.g. 11/25/2026).
_DATE_FORMATS = [
    "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y",
    "%m/%d/%Y", "%m-%d-%Y", "%d-%b-%Y", "%d-%B-%Y", "%d %b %Y", "%d %B %Y",
    "%b %d, %Y", "%B %d, %Y", "%d/%m/%y", "%m/%d/%y", "%d-%m-%y",
    "%d.%m.%y", "%d-%b-%y", "%Y%m%d",
]


def _parse_date_text(strs: pd.Series):
    cleaned = strs.str.strip().str.replace(_TIME_SUFFIX, "", regex=True).str.strip()
    best, best_fmt, best_n = None, None, -1
    for fmt in _DATE_FORMATS:
        parsed = pd.to_datetime(cleaned, format=fmt, errors="coerce")
        n = int(parsed.notna().sum())
        if n > best_n:
            best, best_fmt, best_n = parsed, fmt, n
        if n == len(cleaned):
            break
    return best, best_fmt


def _strip_tz(parsed: pd.Series) -> pd.Series:
    if getattr(parsed.dt, "tz", None) is not None:
        parsed = parsed.dt.tz_localize(None)
    return parsed


def parse_dates(s: pd.Series):
    """Parse a column to date-only datetime64 values.

    Returns ``(dates, blank, invalid, notes)``. ``dates`` are normalised to
    midnight so no time component survives.
    """
    notes: list[str] = []
    blank = blank_mask(s)

    if pd.api.types.is_datetime64_any_dtype(s):
        dates = _strip_tz(pd.to_datetime(s)).dt.normalize()
        return dates, blank, pd.Series(False, index=s.index), notes

    s = s.astype(object)
    parts = []
    kind = s.map(lambda v: "dt" if isinstance(v, (datetime, date, pd.Timestamp))
                 else "str" if isinstance(v, str)
                 else "num" if isinstance(v, (int, float, np.integer, np.floating))
                 and not isinstance(v, (bool, np.bool_))
                 else "other")

    m = (kind == "dt") & ~blank
    if m.any():
        parts.append(_strip_tz(pd.to_datetime(s[m], errors="coerce")))

    m = (kind == "num") & ~blank
    if m.any():
        nums = pd.to_numeric(s[m], errors="coerce").astype("float64")
        nums = nums.where((nums >= 1) & (nums < 2958466))  # valid Excel serial range
        parts.append(pd.to_datetime(nums, unit="D", origin="1899-12-30", errors="coerce"))

    m = (kind == "str") & ~blank
    if m.any():
        parsed, fmt = _parse_date_text(s[m])
        parts.append(parsed)
        if fmt and parsed.notna().any():
            notes.append(f"Text dates were read using the format {fmt} "
                         "(%d = day, %m = month, %Y = year).")

    if parts:
        dates = pd.concat(parts).reindex(s.index)
        dates = pd.to_datetime(dates).dt.normalize()
    else:
        dates = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    invalid = dates.isna() & ~blank
    return dates, blank, invalid, notes


# --------------------------------------------------------------------------- #
# Identifier / text key cleaning
# --------------------------------------------------------------------------- #

_TRAILING_ZERO = re.compile(r"^(\d+)\.0+$")


def _clean_one(v):
    if isinstance(v, str):
        v = v.strip()
        m = _TRAILING_ZERO.match(v)
        return m.group(1) if m else v
    if isinstance(v, (bool, np.bool_)):
        return str(bool(v))
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        if v != v:
            return np.nan
        return str(int(v)) if float(v).is_integer() else repr(float(v))
    return str(v).strip()


def clean_key(s: pd.Series) -> pd.Series:
    """Clean a grouping-key column: text, trimmed, no '.0', blanks -> NaN."""
    blank = blank_mask(s)
    out = s.astype(object).map(_clean_one).astype(object)
    out[blank] = np.nan
    return out


def unique_divisions(division: pd.Series) -> list[str]:
    return sorted(str(v) for v in division.dropna().unique())


# --------------------------------------------------------------------------- #
# Source preparation
# --------------------------------------------------------------------------- #


@dataclass
class SourceFrame:
    src: pd.DataFrame     # required columns only, original values, blank rows removed
    rows: pd.Series       # original worksheet row number (header = row 1)
    blank_rows: int       # fully blank rows that were skipped


def prepare_source(raw: pd.DataFrame, required: list[str]) -> SourceFrame:
    """Select the required columns from the raw data (raw is never modified)."""
    if raw is None or raw.shape[0] == 0:
        raise EmptyDataError("The file has no data rows to process.")
    validation = validate_columns(raw, required)
    if validation.missing:
        raise MissingColumnsError(validation.missing)

    blank = blank_rows_mask(raw)
    keep = ~blank
    positions = [validation.positions[c] for c in required]
    src = raw.iloc[:, positions].copy()
    src.columns = list(required)
    rows = pd.Series(np.arange(len(raw)) + 2, index=raw.index)
    src, rows = src[keep], rows[keep]
    if src.empty:
        raise EmptyDataError("Every row in the file is blank.")
    return SourceFrame(src=src, rows=rows, blank_rows=int(blank.sum()))


# --------------------------------------------------------------------------- #
# Issue (validation report) helpers
# --------------------------------------------------------------------------- #


def display_value(v) -> str:
    if v is None or (isinstance(v, float) and v != v) or v is pd.NaT:
        return "(blank)"
    if isinstance(v, str):
        return v if v.strip() else "(blank)"
    return str(v)


def make_issues(rows: pd.Series, values: pd.Series, mask: pd.Series,
                column: str, issue: str, action: str):
    """Build issue rows for ``mask``; returns None when there is nothing."""
    if not mask.any():
        return None
    return pd.DataFrame({
        "Source Row": rows[mask].to_numpy(),
        "Column": column,
        "Value": values[mask].map(display_value).to_numpy(),
        "Issue": issue,
        "Action": action,
    })


def combine_issues(frames: list) -> pd.DataFrame:
    frames = [f for f in frames if f is not None and not f.empty]
    if not frames:
        return pd.DataFrame(columns=ISSUE_COLUMNS)
    out = pd.concat(frames, ignore_index=True)
    return out.sort_values(["Source Row", "Column"], kind="mergesort").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Sorting
# --------------------------------------------------------------------------- #


def _sort_key(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s) or pd.api.types.is_numeric_dtype(s):
        return s
    num = pd.to_numeric(s, errors="coerce")
    if s.notna().any() and num.notna().sum() == s.notna().sum():
        return num                      # numeric-looking codes sort numerically
    return s.astype(object).str.lower()


def sort_frame(df: pd.DataFrame, by: list[str]) -> pd.DataFrame:
    if df.empty:
        return df
    return df.sort_values(by=by, key=_sort_key, na_position="last",
                          kind="mergesort").reset_index(drop=True)
