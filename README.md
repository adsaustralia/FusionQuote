# Excel Formula Fusion V2.0 Safe Stock Rate Build

This build focuses on stability and clarity around stock/material rate entry.

## Key fixes

- Upload responds first before any heavy Excel processing.
- Stock/material names are scanned only from the selected item start column to end column.
- Selected stock/material rates have a clear **Stock/material rates** section.
- Rates are staged and only saved when **Refresh / Update Rates** is pressed.
- Workbook generation only runs when **Generate Excel Workbook** is pressed.
- Country scan range is auto-detected to keep formulas lighter.
- Data-only workbook loading is read-only to reduce Streamlit Cloud memory use.
- Error details are shown inside the app instead of silently failing.

## Default HOKA mapping

Working sheet defaults:

- Name row: 4
- Size row: 5
- Stock/material row: 6
- Original total qty row: 7
- Country column: I
- DS/SS output row: 168
- Clean qty output row: 169
- Multiplier output row: 170
- SQM output row: 171
- Price output row: 172

Reference sheet defaults:

- Reference name column: C
- Reference size column: E
- Reference DS/SS column: F
- Reference stock/material column: G

## How to use stock rates

1. Upload the Excel workbook.
2. Press **Apply Mapping / Refresh Stock List**.
3. Pick one or more stocks from **Pick stock/material to calculate SQM and rate**.
4. Enter each $/sqm in the **Stock/material rates** section.
5. Press **Refresh / Update Rates**.
6. Press **Generate Excel Workbook** only when ready.

## Deployment

Upload these files to the root of your GitHub repository:

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `README.md`

Streamlit Cloud main file path:

```text
app.py
```

## Note on saved rates

Streamlit Cloud does not permanently store server-side session memory. Download the stock rate memory JSON and upload it next time to reuse rates.
