"""Reading uploaded CSV / Excel files and keeping the raw data untouched.

* CSV files are read as text (every cell stays exactly as written).
* Excel files are read with their native cell types (dates stay dates,
  numbers stay numbers); nothing is converted or corrected here.
* Header row = first row of the file / worksheet.
"""
from __future__ import annotations

import csv
import io
import os
from dataclasses import dataclass, field

import pandas as pd

from .validation import EmptyDataError, MISError, blank_mask

SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")
_ENCODINGS = ("utf-8-sig", "cp1252", "latin-1")


@dataclass
class FileInfo:
    name: str
    size_bytes: int
    extension: str
    sheet_names: list = field(default_factory=list)
    sheet: str | None = None
    rows: int = 0
    columns: int = 0


def get_extension(filename: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise MISError(f"Unsupported file type '{ext or filename}'. "
                       "Please upload an .xlsx, .xls or .csv file.")
    return ext


def format_size(num_bytes: int) -> str:
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:,.0f} {unit}" if unit == "B" else f"{size:,.1f} {unit}"
        size /= 1024
    return f"{num_bytes} B"


def list_sheets(data: bytes, filename: str) -> list[str]:
    """Worksheet names for Excel files; empty list for CSV."""
    ext = get_extension(filename)
    if ext == ".csv":
        return []
    try:
        return list(pd.ExcelFile(io.BytesIO(data)).sheet_names)
    except Exception as exc:  # noqa: BLE001 - converted to a friendly message
        raise MISError(_excel_hint(ext, exc)) from exc


def _excel_hint(ext: str, exc: Exception) -> str:
    msg = str(exc)
    if "xlrd" in msg.lower():
        return ("Reading .xls files needs the 'xlrd' package. "
                "Run: pip install xlrd (it is listed in requirements.txt).")
    return ("The Excel file could not be read. It may be corrupted, password "
            f"protected or not a real {ext} file.")


def _read_csv(data: bytes) -> pd.DataFrame:
    last_exc: Exception | None = None
    for enc in _ENCODINGS:
        try:
            text_sample = data[:20000].decode(enc)
            try:
                delimiter = csv.Sniffer().sniff(text_sample, delimiters=",;\t|").delimiter
            except csv.Error:
                delimiter = ","
            return pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False,
                               encoding=enc, sep=delimiter)
        except UnicodeDecodeError as exc:
            last_exc = exc
        except pd.errors.EmptyDataError as exc:
            raise EmptyDataError("The CSV file is empty.") from exc
        except pd.errors.ParserError as exc:
            raise MISError("The CSV file could not be parsed. Check that every row "
                           "has the same number of columns.") from exc
    raise MISError("The CSV file's text encoding could not be read.") from last_exc


def _read_excel(data: bytes, ext: str, sheet: str | None) -> pd.DataFrame:
    engines = ["openpyxl", "xlrd"] if ext == ".xlsx" else ["xlrd", "openpyxl"]
    last_exc: Exception | None = None
    for engine in engines:                      # tolerate a wrong extension
        try:
            return pd.read_excel(io.BytesIO(data), sheet_name=sheet or 0,
                                 dtype=object, header=0, engine=engine)
        except ImportError as exc:
            last_exc = exc
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
    raise MISError(_excel_hint(ext, last_exc)) from last_exc


def load_raw(data: bytes, filename: str, sheet: str | None = None) -> pd.DataFrame:
    """Read the uploaded file exactly as it is. This is the 'Raw Data' copy."""
    ext = get_extension(filename)
    if not data:
        raise EmptyDataError("The uploaded file is empty.")
    df = _read_csv(data) if ext == ".csv" else _read_excel(data, ext, sheet)

    df.columns = [str(c) for c in df.columns]
    # Drop phantom columns created by stray formatting (blank 'Unnamed: n').
    phantom = [c for c in df.columns
               if c.startswith("Unnamed:") and blank_mask(df[c]).all()]
    if phantom:
        df = df.drop(columns=phantom)

    if df.shape[1] == 0:
        raise EmptyDataError("The file or worksheet is empty.")
    if df.shape[0] == 0:
        raise EmptyDataError("The file has column headers but no data rows.")
    return df.reset_index(drop=True)


def describe(data: bytes, filename: str, sheet: str | None, sheets: list[str],
             raw: pd.DataFrame) -> FileInfo:
    return FileInfo(name=filename, size_bytes=len(data), extension=get_extension(filename),
                    sheet_names=sheets, sheet=sheet, rows=int(raw.shape[0]),
                    columns=int(raw.shape[1]))
