# MIS File Processor

A Streamlit app that processes **Primary** and **Delivery** MIS files independently,
validates them, shows a preview and exports an Excel workbook.

Each exported workbook has exactly two sheets:

| Sheet | Content |
|---|---|
| `Primary` or `Delivery` | Processed output (date-only Excel dates shown as `mm/dd/yy`) |
| `Raw Data` | The original upload, never filtered, corrected or consolidated |

## Features

- Sidebar navigation: Home / Dashboard, Primary File Processor, Delivery File Processor
- `.xlsx`, `.xls` and `.csv` uploads, with worksheet selection for Excel files
- Required-column validation (leading/trailing spaces in headers are ignored)
- Processing summary cards, data-quality notes and a preview before download
- Automatic file names, e.g. `Primary PL-K (Nov 1 to Nov 25).xlsx`
- Optional validation report (Excel) listing every flagged row with its original row number

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Run the tests:

```bash
pip install -r requirements-dev.txt
pytest
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository and push this project (the `.gitignore` keeps Excel/CSV data out of Git).
2. Go to <https://share.streamlit.io>, sign in with GitHub and choose **New app**.
3. Select the repository and branch, and set the main file path to `app.py`.
4. Click **Deploy**. Streamlit installs everything in `requirements.txt` automatically.

Community Cloud has limited memory (about 1 GB). For very large files (hundreds of thousands
of rows) run the app locally instead. The upload limit is set in `.streamlit/config.toml`
(`maxUploadSize`, in MB).

## Processing rules

### Primary file (`Document Date` is the file-name date)

1. The 21 required columns must exist, otherwise processing stops and the missing names are listed.
2. **SHORT CLOSED** rows only (case-insensitive, extra spaces ignored):
   - `Order quantity` = `Delivery qty`
   - `Net Value` = `Delivery qty x Net Price` (Net Price is the price of **one** unit; no rounding)
   - All other statuses keep their original `Order quantity` and `Net Value`.
3. Rows whose corrected `Net Value` is exactly zero are removed. Negative values are kept.
4. Rows are consolidated on `Document Date`, `Sold-to party`, `Customer Name`, `Division`,
   `Material`, `Description` (sum of `Order quantity` and `Net Value`).
5. Output columns: the six keys + `(Sum of) Order quantity` + `(Sum of) Net Value`,
   sorted by Document Date, Sold-to party, Material.

### Delivery file (`Billing Date` is the file-name date)

1. The 8 required columns must exist.
2. Rows whose `Item Value` is exactly zero are removed. Negative values and zero quantities
   with a non-zero value are kept.
3. Rows are consolidated on `Billing Date`, `Customer Code`, `Customer Name`, `Division`,
   `Material`, `Material Description` (sum of `Billed Quantity` and `Item Value`).
4. Sorted by Billing Date, Customer Code, Material.

### Data-handling policy

- **Invalid rows are excluded and reported.** A row is invalid when its date is blank or
  unreadable, or a number the calculation needs is blank or not a valid number (blank is never
  treated as zero). Primary needs `Order quantity` and `Net Value` on normal rows, and
  `Delivery qty` and `Net Price` on SHORT CLOSED rows. Delivery needs `Billed Quantity` and
  `Item Value`. The excluded rows appear in the validation report with their original
  worksheet row number and the reason; they stay untouched in `Raw Data`.
- **Missing key values are kept.** Blank Sold-to party / Customer Code, Material, Division and
  so on stay in the output as their own group (and are listed in the report as "Kept").
- **Duplicates are kept** and added together during consolidation (the count is shown as a note).
- **Identifiers** (customer codes, materials) are treated as text: no `.0`, leading zeros are
  preserved, and they are written to Excel as text.
- **Numbers** are never rounded when stored. Numbers such as `1,234.50`, `(100)` and `100-`
  are understood.
- **Dates** are converted to date only (time removed). Text dates are read in the format
  that fits the whole column (day-first is preferred when a column is ambiguous); the
  format used is shown in the data-quality notes.
- **File name**: the date range comes from all valid dates in the uploaded rows *before*
  zero-value filtering. One Division gives its name (for example `PL-K`); several Divisions
  give `Multi-Division` plus a warning; none gives `Unknown-Division` plus a warning.
- The Primary summary includes an information-only check comparing
  `Order quantity x Net Price` with `Net Value` on non-SHORT CLOSED rows. It never changes data.

## Project structure

```text
mis-file-processor/
├── app.py                      # Streamlit UI: navigation, upload, summary, preview, downloads
├── requirements.txt
├── requirements-dev.txt        # adds pytest
├── pytest.ini
├── README.md
├── .gitignore
├── .streamlit/config.toml      # upload size limit, theme
├── modules/
│   ├── __init__.py
│   ├── primary_processor.py    # SHORT CLOSED correction, zero filter, consolidation, sort
│   ├── delivery_processor.py   # zero filter, consolidation, sort
│   ├── file_handler.py         # CSV/Excel reading, raw data preservation
│   ├── validation.py           # column, date, number and data-quality helpers
│   └── excel_export.py         # two-sheet workbook, formats, file names, validation report
└── tests/
    └── test_processors.py      # synthetic-data tests for both processors and the export
```

## Notes

- The header row must be the first row of the file or worksheet.
- CSV files are kept as text in `Raw Data` (exactly as written in the file).
- Typical speed on a modern laptop for a 200,000-row Primary file: about 15 s to read the
  `.xlsx`, 4 s to process, 25 s to export. Normal daily files are much faster.
