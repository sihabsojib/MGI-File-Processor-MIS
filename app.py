"""MIS File Processor - Streamlit entry point.

Run locally:  streamlit run app.py
"""
from __future__ import annotations

import logging
from datetime import date

import pandas as pd
import streamlit as st

from modules import delivery_processor, excel_export, file_handler, primary_processor
from modules.validation import MISError, validate_columns

logger = logging.getLogger("mis_file_processor")

st.set_page_config(page_title="MIS File Processor", page_icon=":bar_chart:",
                   layout="wide")

MODULES = {
    "primary": {
        "title": "Primary File Processor",
        "required": primary_processor.PRIMARY_REQUIRED,
        "process": primary_processor.process_primary,
        "specific_title": "Primary-specific checks (SHORT CLOSED correction)",
        "date_column": "Document Date",
    },
    "delivery": {
        "title": "Delivery File Processor",
        "required": delivery_processor.DELIVERY_REQUIRED,
        "process": delivery_processor.process_delivery,
        "specific_title": "Delivery-specific totals",
        "date_column": "Billing Date",
    },
}
STATE_KEYS = ("sig", "raw", "info", "result", "xlsx", "report")


# --------------------------------------------------------------------------- #
# Small UI helpers
# --------------------------------------------------------------------------- #


def _key(kind: str, name: str) -> str:
    return f"{kind}_{name}"


def _clear(kind: str) -> None:
    for name in STATE_KEYS:
        st.session_state.pop(_key(kind, name), None)


def _fmt(value) -> str:
    if isinstance(value, date):
        return value.strftime("%m/%d/%y")
    if value is None:
        return "-"
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return f"{value:,}"
    if isinstance(value, float):
        if value == int(value):
            return f"{value:,.2f}"
        return f"{value:,.4f}".rstrip("0").rstrip(".")   # display only, not rounding data
    return str(value)


def _metric_grid(items: list, per_row: int = 4) -> None:
    for start in range(0, len(items), per_row):
        cols = st.columns(per_row)
        for col, (label, value) in zip(cols, items[start:start + per_row]):
            col.metric(label, _fmt(value))


def _show_df(df: pd.DataFrame) -> None:
    try:
        st.dataframe(df, hide_index=True, width="stretch")
    except Exception:  # older Streamlit versions
        st.dataframe(df, hide_index=True, use_container_width=True)


def _preview_frame(df: pd.DataFrame, date_column: str, rows: int = 500) -> pd.DataFrame:
    preview = df.head(rows).copy()
    if date_column in preview.columns:
        preview[date_column] = pd.to_datetime(preview[date_column]).dt.strftime("%m/%d/%y")
    return preview


# --------------------------------------------------------------------------- #
# Pages
# --------------------------------------------------------------------------- #


def render_home() -> None:
    st.title("MIS File Processor")
    st.write("Process **Primary** and **Delivery** files, check the results, and "
             "download a ready-to-use Excel workbook. The two processors work "
             "independently - processing one never affects the other.")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Primary file")
        st.markdown(
            "- Uses the 21 Primary columns\n"
            "- **SHORT CLOSED** rows: Order quantity = Delivery qty, "
            "Net Value = Delivery qty x Net Price\n"
            "- Removes rows with zero Net Value (negatives kept)\n"
            "- Consolidates by Date, Sold-to party, Customer, Division, "
            "Material, Description")
        st.caption("Status this session: " +
                   ("processed" if _key("primary", "result") in st.session_state
                    else "not processed"))
    with c2:
        st.subheader("Delivery file")
        st.markdown(
            "- Uses the 8 Delivery columns\n"
            "- Removes rows with zero Item Value (negatives kept)\n"
            "- Consolidates by Billing Date, Customer Code, Customer Name, "
            "Division, Material, Material Description")
        st.caption("Status this session: " +
                   ("processed" if _key("delivery", "result") in st.session_state
                    else "not processed"))

    st.subheader("What you get")
    st.markdown(
        "- A workbook with exactly two sheets: the processed sheet and **Raw Data** "
        "(your original upload, untouched)\n"
        "- Dates are real Excel dates, date only, shown as `mm/dd/yy`\n"
        "- A validation report if any rows had invalid dates or numbers "
        "(those rows are excluded from the processed sheet, never silently dropped)\n"
        "- File names are generated automatically from the Division and date range")
    st.info("Use the sidebar to open the Primary or Delivery processor.")


def render_processor(kind: str) -> None:
    cfg = MODULES[kind]
    st.title(cfg["title"])

    uploaded = st.file_uploader(
        "1. Upload file (.xlsx, .xls or .csv)", type=["xlsx", "xls", "csv"],
        key=_key(kind, "uploader"))
    if uploaded is None:
        _clear(kind)
        st.info("Upload a file to begin.")
        return

    data = uploaded.getvalue()
    try:
        sheets = file_handler.list_sheets(data, uploaded.name)
    except MISError as exc:
        _clear(kind)
        st.error(str(exc))
        return

    sheet = None
    if sheets:
        sheet = st.selectbox("2. Select worksheet", sheets, key=_key(kind, "sheet"))

    file_id = getattr(uploaded, "file_id", None) or uploaded.name
    signature = (file_id, uploaded.name, len(data), sheet)
    if st.session_state.get(_key(kind, "sig")) != signature:
        _clear(kind)
        try:
            with st.spinner("Reading file..."):
                raw = file_handler.load_raw(data, uploaded.name, sheet)
        except MISError as exc:
            st.error(str(exc))
            return
        except Exception:  # noqa: BLE001
            logger.exception("Unexpected error while reading the file")
            st.error("The file could not be read. Please check that it is a valid "
                     "Excel or CSV file and try again.")
            return
        st.session_state[_key(kind, "sig")] = signature
        st.session_state[_key(kind, "raw")] = raw
        st.session_state[_key(kind, "info")] = file_handler.describe(
            data, uploaded.name, sheet, sheets, raw)

    raw: pd.DataFrame = st.session_state[_key(kind, "raw")]
    info = st.session_state[_key(kind, "info")]

    # ---- file overview and column validation ---------------------------- #
    st.subheader("File overview")
    _metric_grid([
        ("File", info.name),
        ("File size", file_handler.format_size(info.size_bytes)),
        ("Worksheet", info.sheet or "CSV (no worksheets)"),
        ("Original rows x columns", f"{info.rows:,} x {info.columns:,}"),
    ])

    validation = validate_columns(raw, cfg["required"])
    if validation.missing:
        st.error("Required column(s) missing: " + ", ".join(validation.missing) +
                 ". Processing is stopped. Please check the file or the selected worksheet.")
        with st.expander("Columns found in the file"):
            st.write([str(c) for c in raw.columns])
        return
    st.success(f"Required-column validation passed: all {len(cfg['required'])} "
               "required columns found.")

    # ---- processing ------------------------------------------------------ #
    st.subheader("Processing")
    if st.button("Process file", type="primary", key=_key(kind, "run")):
        progress = st.progress(0, text="Validating and processing data...")
        try:
            result = cfg["process"](raw)
            progress.progress(45, text="Building the Excel workbook...")
            xlsx = None
            if not result.processed.empty:
                xlsx = excel_export.build_workbook(kind, result.processed, raw)
            progress.progress(85, text="Preparing the validation report...")
            report = (excel_export.build_validation_report(result)
                      if not result.issues.empty else None)
            progress.progress(100, text="Done.")
        except MISError as exc:
            progress.empty()
            st.error(str(exc))
            return
        except Exception:  # noqa: BLE001
            progress.empty()
            logger.exception("Unexpected error while processing")
            st.error("Something went wrong while processing the file. Please check "
                     "that the data matches the expected layout and try again.")
            return
        progress.empty()
        st.session_state[_key(kind, "result")] = result
        st.session_state[_key(kind, "xlsx")] = xlsx
        st.session_state[_key(kind, "report")] = report

    result = st.session_state.get(_key(kind, "result"))
    if result is None:
        st.caption("Click **Process file** to validate, process and preview the data.")
        return

    # ---- summary --------------------------------------------------------- #
    st.subheader("Processing summary")
    for message in result.warnings:
        st.warning(message)
    if result.processed.empty:
        st.error("The processed output is empty, so no workbook was created.")
    else:
        st.success("Processing complete.")

    _metric_grid(result.metrics)
    st.markdown(f"**{cfg['specific_title']}**")
    _metric_grid(result.specific_metrics)

    if result.infos:
        with st.expander("Data-quality notes"):
            for note in result.infos:
                st.write("- " + note)

    # ---- preview ---------------------------------------------------------- #
    st.subheader("Processed data preview")
    if result.processed.empty:
        st.write("No rows to preview.")
    else:
        _show_df(_preview_frame(result.processed, cfg["date_column"]))
        if len(result.processed) > 500:
            st.caption(f"Showing the first 500 of {len(result.processed):,} rows. "
                       "The download contains all rows.")

    # ---- downloads --------------------------------------------------------- #
    st.subheader("Download")
    xlsx = st.session_state.get(_key(kind, "xlsx"))
    report = st.session_state.get(_key(kind, "report"))
    if xlsx:
        st.download_button(
            f"Download {excel_export.result_filename(result)}", data=xlsx,
            file_name=excel_export.result_filename(result),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", key=_key(kind, "dl"))
    if report:
        st.download_button(
            "Download validation report", data=report,
            file_name=excel_export.report_filename(result),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=_key(kind, "dl_report"))
        with st.expander(f"Validation issues ({len(result.issues):,})"):
            _show_df(result.issues.head(1000))
    elif not result.issues.empty:
        st.info("Validation issues were found; see the notes above.")


# --------------------------------------------------------------------------- #
# Navigation
# --------------------------------------------------------------------------- #

PAGES = {
    "Home / Dashboard": render_home,
    "Primary File Processor": lambda: render_processor("primary"),
    "Delivery File Processor": lambda: render_processor("delivery"),
}

st.sidebar.title("MIS File Processor")
page = st.sidebar.radio("Navigation", list(PAGES), label_visibility="collapsed")
PAGES[page]()
