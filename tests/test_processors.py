"""Tests for the MIS File Processor business logic (synthetic data only).

Run with:  pytest
"""
import io
from datetime import date, datetime

import openpyxl
import pandas as pd
import pytest

from modules import excel_export, file_handler
from modules.delivery_processor import DELIVERY_REQUIRED, process_delivery
from modules.excel_export import build_filename, build_validation_report, build_workbook
from modules.primary_processor import (OUTPUT_COLUMNS as P_COLS, PRIMARY_REQUIRED,
                                       process_primary)
from modules.validation import EmptyDataError, MissingColumnsError, parse_dates, parse_numeric


# --------------------------------------------------------------------------- #
# Synthetic data builders
# --------------------------------------------------------------------------- #

def prim(date_=datetime(2026, 11, 1), sold="1001", name="Cust A", div="PL-K", mat="M1",
         desc="Item 1", oq=10, price=5, nv=50, dq=10, status="Open"):
    row = {c: None for c in PRIMARY_REQUIRED}
    row.update({"Document Date": date_, "Sold-to party": sold, "Customer Name": name,
                "Division": div, "Material": mat, "Description": desc,
                "Order quantity": oq, "Net Price": price, "Net Value": nv,
                "Delivery qty": dq, "Order Status": status})
    return row


def deli(date_=datetime(2026, 11, 1), code="2001", name="Cust A", div="PL-K", mat="M1",
         desc="Item 1", qty=10, val=100):
    return {"Billing Date": date_, "Customer Code": code, "Customer Name": name,
            "Division": div, "Material": mat, "Material Description": desc,
            "Billed Quantity": qty, "Item Value": val}


def frame(rows, extra=True):
    df = pd.DataFrame(rows).astype(object)
    if extra:
        df["Extra col"] = "x"          # a non-required column that must stay in Raw Data
    return df


def run_primary(rows):
    return process_primary(frame(rows))


def only(result):
    assert len(result.processed) == 1
    return result.processed.iloc[0]


# --------------------------------------------------------------------------- #
# Primary
# --------------------------------------------------------------------------- #

def test_p1_open_row_unchanged():
    r = only(run_primary([prim(oq=10, nv=77, status="Open")]))
    assert r["(Sum of) Order quantity"] == 10 and r["(Sum of) Net Value"] == 77


def test_p2_p3_short_closed_qty_and_value():
    r = only(run_primary([prim(oq=10, dq=2, price=10, nv=100, status="SHORT CLOSED")]))
    assert r["(Sum of) Order quantity"] == 2
    assert r["(Sum of) Net Value"] == 20          # 2 x 10, per-unit Net Price


@pytest.mark.parametrize("status", ["SHORT CLOSED", "short closed", "Short Closed",
                                    "  SHORT CLOSED  ", "short   closed", "SHORT\u00a0CLOSED"])
def test_p4_short_closed_variants(status):
    res = run_primary([prim(oq=10, dq=3, price=4, nv=999, status=status)])
    assert only(res)["(Sum of) Net Value"] == 12
    assert dict(res.specific_metrics)["SHORT CLOSED rows detected"] == 1


def test_p5_short_closed_zero_delivery_is_zero_then_filtered():
    res = run_primary([prim(dq=0, price=10, nv=100, status="SHORT CLOSED"), prim(mat="M2")])
    assert len(res.processed) == 1 and res.processed.iloc[0]["Material"] == "M2"
    assert dict(res.metrics)["Rows removed (zero Net Value)"] == 1


def test_p6_negative_net_value_retained():
    r = only(run_primary([prim(oq=-5, nv=-25)]))
    assert r["(Sum of) Net Value"] == -25 and r["(Sum of) Order quantity"] == -5


def test_p7_zero_net_value_removed_for_open_row():
    res = run_primary([prim(nv=0, status="Open"), prim(mat="M2", nv=10)])
    assert list(res.processed["Material"]) == ["M2"]


def test_p8_duplicate_keys_consolidated():
    res = run_primary([prim(oq=10, nv=50), prim(oq=4, nv=20), prim(mat="M2", oq=1, nv=1)])
    assert len(res.processed) == 2
    m1 = res.processed[res.processed["Material"] == "M1"].iloc[0]
    assert m1["(Sum of) Order quantity"] == 14 and m1["(Sum of) Net Value"] == 70
    assert dict(res.metrics)["Rows consolidated"] == 1


def test_p8b_not_grouped_by_status_or_price():
    res = run_primary([prim(price=5, status="Open", nv=50, oq=10),
                       prim(price=6, status="Closed", nv=60, oq=10)])
    assert len(res.processed) == 1


def test_p9_dates_with_time_export_as_date_only():
    rows = [prim(datetime(2026, 11, 25, 14, 35), nv=1), prim(datetime(2026, 11, 25, 9, 5), nv=2)]
    res = run_primary(rows)
    assert len(res.processed) == 1                    # same day -> one group
    assert res.processed["Document Date"].iloc[0] == pd.Timestamp("2026-11-25")
    wb = openpyxl.load_workbook(io.BytesIO(build_workbook("primary", res.processed, frame(rows))))
    cell = wb["Primary"]["A2"]
    assert cell.number_format == "mm/dd/yy" and cell.is_date
    assert cell.value == datetime(2026, 11, 25, 0, 0)  # no time component


def test_p10_missing_required_columns():
    df = frame([prim()]).drop(columns=["Net Price", "Order Status"])
    with pytest.raises(MissingColumnsError) as exc:
        process_primary(df)
    assert exc.value.missing == ["Net Price", "Order Status"]


def test_p10b_headers_with_spaces_are_matched():
    df = frame([prim()])
    df.columns = [f"  {c} " if c in PRIMARY_REQUIRED else c for c in df.columns]
    assert len(process_primary(df).processed) == 1


def test_p11_multiple_divisions_warning_and_filename():
    res = run_primary([prim(div="PL-K"), prim(div="PL-X", mat="M2")])
    assert any("Multiple Divisions" in w for w in res.warnings)
    assert "Multi-Division" in excel_export.result_filename(res)


def test_p12_invalid_dates_and_inputs_reported_and_excluded():
    res = run_primary([
        prim(date_="not a date"),
        prim(date_=None, mat="M2"),
        prim(dq="abc", status="SHORT CLOSED", mat="M3"),
        prim(price=None, status="SHORT CLOSED", mat="M4"),
        prim(oq="xx", mat="M5"),
        prim(mat="OK"),
    ])
    assert list(res.processed["Material"]) == ["OK"]
    assert res.rows_excluded == 5
    issues = res.issues
    assert set(issues["Action"]) == {"Excluded"}
    assert sorted(issues["Source Row"]) == [2, 3, 4, 5, 6]
    assert dict(res.specific_metrics)["Invalid SHORT CLOSED calculations"] == 2
    assert any("SHORT CLOSED calculation issue" in i for i in issues["Issue"])


def test_p13_non_short_closed_never_recalculated():
    # Delivery qty / Net Price disagree with Net Value on purpose
    r = only(run_primary([prim(oq=10, dq=1, price=1, nv=999, status="Open")]))
    assert r["(Sum of) Net Value"] == 999 and r["(Sum of) Order quantity"] == 10


def test_p14_missing_keys_kept_and_not_dropped():
    res = run_primary([prim(sold=None, mat=None, nv=5), prim(sold=None, mat=None, nv=7)])
    r = only(res)
    assert r["(Sum of) Net Value"] == 12 and pd.isna(r["Sold-to party"])
    assert set(res.issues["Action"]) == {"Kept"}


def test_p15_no_rounding_and_identifiers_stay_text():
    res = run_primary([prim(sold=1001.0, mat="00045", oq=0.123456789, nv=1234.56789012)])
    r = only(res)
    assert r["Sold-to party"] == "1001" and r["Material"] == "00045"
    assert r["(Sum of) Order quantity"] == 0.123456789
    assert r["(Sum of) Net Value"] == 1234.56789012


def test_p16_sorting_and_output_columns():
    res = run_primary([
        prim(datetime(2026, 11, 3), sold="200", mat="B"),
        prim(datetime(2026, 11, 2), sold="300", mat="A"),
        prim(datetime(2026, 11, 2), sold="30", mat="A"),
        prim(datetime(2026, 11, 2), sold="30", mat="0"),
    ])
    assert list(res.processed.columns) == P_COLS
    assert list(res.processed["Sold-to party"]) == ["30", "30", "300", "200"]
    assert list(res.processed["Material"])[:2] == ["0", "A"]


def test_p17_empty_after_filtering():
    res = run_primary([prim(nv=0)])
    assert res.processed.empty and list(res.processed.columns) == P_COLS
    assert any("No rows remain" in w for w in res.warnings)


def test_p18_blank_rows_skipped_and_empty_file():
    df = frame([prim(), {c: None for c in PRIMARY_REQUIRED}], extra=False)
    res = process_primary(df)
    assert dict(res.metrics)["Blank rows skipped"] == 1
    with pytest.raises(EmptyDataError):
        process_primary(pd.DataFrame(columns=PRIMARY_REQUIRED))


def test_p19_raw_data_not_modified_by_processing():
    raw = frame([prim(status=" short closed ", dq=1, price=2, nv=100), prim(nv=0)])
    before = raw.copy(deep=True)
    process_primary(raw)
    pd.testing.assert_frame_equal(raw, before)


def test_p20_sap_style_numbers_and_text_dates():
    rows = [prim("25.11.2026", oq="1,234.50", nv="100-"), prim("26.11.2026", oq="(5)", nv="2,000")]
    res = run_primary(rows)
    assert list(res.processed["Document Date"]) == [pd.Timestamp("2026-11-25"),
                                                    pd.Timestamp("2026-11-26")]
    assert list(res.processed["(Sum of) Order quantity"]) == [1234.5, -5]
    assert list(res.processed["(Sum of) Net Value"]) == [-100, 2000]


# --------------------------------------------------------------------------- #
# Delivery
# --------------------------------------------------------------------------- #

def run_delivery(rows):
    return process_delivery(frame(rows))


def test_d1_zero_item_value_removed():
    res = run_delivery([deli(val=0), deli(mat="M2", val=5)])
    assert list(res.processed["Material"]) == ["M2"]
    assert dict(res.metrics)["Rows removed (zero Item Value)"] == 1


def test_d2_negative_item_value_retained():
    assert only(run_delivery([deli(qty=-2, val=-30)]))["Item Value"] == -30


def test_d3_duplicate_keys_consolidated():
    res = run_delivery([deli(qty=1, val=10), deli(qty=2, val=20), deli(mat="M2")])
    m1 = res.processed[res.processed["Material"] == "M1"].iloc[0]
    assert m1["Billed Quantity"] == 3 and m1["Item Value"] == 30


def test_d4_dates_with_time_export_as_date_only():
    rows = [deli(datetime(2026, 11, 25, 23, 59, 59)), deli(datetime(2026, 1, 5, 9, 30), mat="M2")]
    res = run_delivery(rows)
    wb = openpyxl.load_workbook(io.BytesIO(build_workbook("delivery", res.processed, frame(rows))))
    ws = wb["Delivery"]
    assert [ws["A2"].value, ws["A3"].value] == [datetime(2026, 1, 5), datetime(2026, 11, 25)]
    assert ws["A2"].number_format == ws["A3"].number_format == "mm/dd/yy"


def test_d5_missing_required_columns():
    with pytest.raises(MissingColumnsError):
        process_delivery(frame([deli()]).drop(columns=["Item Value"]))


def test_d6_multiple_divisions():
    res = run_delivery([deli(div="A"), deli(div="B", mat="M2")])
    assert any("Multiple Divisions" in w for w in res.warnings)
    assert excel_export.result_filename(res).startswith("Delivery Multi-Division")


def test_d7_invalid_numbers_reported():
    res = run_delivery([deli(qty="abc"), deli(val="1,5,x", mat="M2"), deli(mat="M3", qty=None),
                        deli(mat="OK")])
    assert list(res.processed["Material"]) == ["OK"]
    assert res.rows_excluded == 3 and len(res.issues) == 3


def test_d8_zero_qty_with_value_retained():
    r = only(run_delivery([deli(qty=0, val=50)]))
    assert r["Billed Quantity"] == 0 and r["Item Value"] == 50


def test_d9_empty_output_handled():
    res = run_delivery([deli(val=0)])
    assert res.processed.empty and list(res.processed.columns) == DELIVERY_REQUIRED
    assert any("No rows remain" in w for w in res.warnings)


def test_d10_missing_customer_code_kept():
    res = run_delivery([deli(code=None, val=3), deli(code=None, val=4)])
    assert only(res)["Item Value"] == 7 and pd.isna(only(res)["Customer Code"])


# --------------------------------------------------------------------------- #
# Export / filenames / file handling
# --------------------------------------------------------------------------- #

def test_workbook_structure_and_raw_preservation():
    raw = frame([prim(datetime(2026, 11, 1, 14, 35), nv=50), prim(nv=0, mat="Z"),
                 prim("bad date", mat="Y")])
    res = process_primary(raw)
    wb = openpyxl.load_workbook(io.BytesIO(build_workbook("primary", res.processed, raw)))
    assert wb.sheetnames == ["Primary", "Raw Data"]
    p, r = wb["Primary"], wb["Raw Data"]
    assert [c.value for c in p[1]] == P_COLS                    # exactly 8 columns, no index
    assert p.max_column == 8 and p.freeze_panes == "A2" and p.auto_filter.ref
    assert r.max_row == len(raw) + 1 and r.max_column == raw.shape[1]
    assert [c.value for c in r[1]] == list(raw.columns)
    assert r["A2"].value == datetime(2026, 11, 1, 14, 35)       # raw keeps its time
    assert r["A4"].value == "bad date"                          # raw keeps invalid values
    assert r.freeze_panes == "A2" and r.auto_filter.ref


def test_identifiers_written_as_text_without_dot_zero():
    raw = frame([prim(sold=100234.0, mat="00045")])
    res = process_primary(raw)
    ws = openpyxl.load_workbook(io.BytesIO(build_workbook("primary", res.processed, raw)))["Primary"]
    assert ws["B2"].value == "100234" and ws["E2"].value == "00045"
    assert ws["G2"].data_type == "n" and ws["H2"].data_type == "n"   # measures stay numeric


def test_filename_formats():
    d1, d2 = date(2026, 11, 1), date(2026, 11, 25)
    assert build_filename("primary", ["PL-K"], d1, d2) == "Primary PL-K (Nov 1 to Nov 25).xlsx"
    assert build_filename("primary", ["PL-K"], d2, d2) == "Primary PL-K (Nov 25).xlsx"
    assert build_filename("delivery", ["PL-K"], d1, d2) == "Delivery PL-K (Nov 1 to Nov 25).xlsx"
    assert build_filename("delivery", ["PL-K"], d2, d2) == "Delivery PL-K (Nov 25).xlsx"
    assert build_filename("delivery", [], d1, d2) == "Delivery Unknown-Division (Nov 1 to Nov 25).xlsx"
    assert build_filename("primary", ['A/B:C*?'], d2, d2) == "Primary ABC (Nov 25).xlsx"


def test_filename_uses_correct_date_field_and_pre_filter_range():
    # The earliest date row has zero value and is filtered out; the name must still
    # cover it because the range comes from the loaded source rows.
    res = run_primary([prim(datetime(2026, 11, 1), nv=0), prim(datetime(2026, 11, 25), nv=5)])
    assert excel_export.result_filename(res) == "Primary PL-K (Nov 1 to Nov 25).xlsx"
    res = run_delivery([deli(datetime(2026, 11, 3), val=1)])
    assert excel_export.result_filename(res) == "Delivery PL-K (Nov 3).xlsx"


def test_excel_row_limit_error():
    big = pd.DataFrame({"a": range(excel_export.MAX_ROWS)})
    with pytest.raises(excel_export.ExportError):
        excel_export.check_excel_limits(big, pd.DataFrame())


def test_validation_report_has_row_numbers():
    res = run_primary([prim(date_="zzz"), prim()])
    wb = openpyxl.load_workbook(io.BytesIO(build_validation_report(res)))
    assert wb.sheetnames == ["Summary", "Issues"]
    assert wb["Issues"]["A2"].value == 2 and wb["Issues"]["E2"].value == "Excluded"


def test_csv_round_trip_keeps_text_and_leading_zeros():
    csv = ("Billing Date,Customer Code,Customer Name,Division,Material,Material Description,"
           "Billed Quantity,Item Value\n"
           "2026-11-25 14:35:00,00123,A,PL-K,0045,Thing,\"1,000\",\"2,500.50\"\n").encode()
    raw = file_handler.load_raw(csv, "x.csv")
    assert raw["Customer Code"].iloc[0] == "00123"
    res = process_delivery(raw)
    r = only(res)
    assert r["Material"] == "0045" and r["Billed Quantity"] == 1000 and r["Item Value"] == 2500.5


def test_excel_file_round_trip_and_sheet_selection():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        pd.DataFrame({"junk": [1]}).to_excel(w, sheet_name="Notes", index=False)
        frame([deli()], extra=False).to_excel(w, sheet_name="Data", index=False)
    data = buf.getvalue()
    assert file_handler.list_sheets(data, "f.xlsx") == ["Notes", "Data"]
    raw = file_handler.load_raw(data, "f.xlsx", "Data")
    assert len(process_delivery(raw).processed) == 1


def test_file_handler_errors():
    with pytest.raises(EmptyDataError):
        file_handler.load_raw(b"", "x.csv")
    with pytest.raises(EmptyDataError):
        file_handler.load_raw(b"a,b,c\n", "x.csv")
    with pytest.raises(file_handler.MISError):
        file_handler.get_extension("report.pdf")
    with pytest.raises(file_handler.MISError):
        file_handler.list_sheets(b"not an excel file", "x.xlsx")


# --------------------------------------------------------------------------- #
# Parsers
# --------------------------------------------------------------------------- #

def test_parse_dates_formats_and_invalid():
    s = pd.Series(["2026-11-25 14:35:00", "2026-11-26", "oops", None, ""], dtype=object)
    d, blank, invalid, _ = parse_dates(s)
    assert d.iloc[0] == pd.Timestamp("2026-11-25") and d.iloc[1] == pd.Timestamp("2026-11-26")
    assert list(invalid) == [False, False, True, False, False]
    assert list(blank) == [False, False, False, True, True]


def test_parse_dates_month_first_when_day_first_cannot_parse():
    d, _, invalid, _ = parse_dates(pd.Series(["11/25/2026", "01/05/2026"], dtype=object))
    assert list(d) == [pd.Timestamp("2026-11-25"), pd.Timestamp("2026-01-05")]
    assert not invalid.any()


def test_parse_dates_excel_serial_and_native():
    serial = (date(2026, 11, 25) - date(1899, 12, 30)).days      # Excel serial number
    s = pd.Series([serial, datetime(2026, 11, 25, 8, 0), date(2026, 1, 5)], dtype=object)
    d, _, invalid, _ = parse_dates(s)
    assert d.iloc[0] == pd.Timestamp("2026-11-25")
    assert d.iloc[1] == pd.Timestamp("2026-11-25") and d.iloc[2] == pd.Timestamp("2026-01-05")
    assert not invalid.any()


def test_parse_numeric_blank_is_not_zero():
    v, blank, invalid = parse_numeric(pd.Series([1, "2.5", None, "", "x", "1e3", True], dtype=object))
    assert v.iloc[0] == 1 and v.iloc[1] == 2.5 and v.iloc[5] == 1000
    assert list(blank) == [False, False, True, True, False, False, False]
    assert list(invalid) == [False, False, False, False, True, False, True]
