# Excel Formula Fusion V2.2 Ultra Stable

This build is designed to stop Streamlit Cloud startup crashes.

## Stability changes

- No heavy workbook processing at app startup.
- `openpyxl` is imported only when needed.
- Upload responds immediately.
- Sheet names are read using lightweight XLSX zip/XML parsing first.
- Workbook generation runs only after pressing **Generate Excel Workbook**.
- Stock names scan only from selected Start Column to End Column.
- Stock rates are entered in a clear **Stock/material rates** section and saved only after **Refresh / Update Rates**.

## Logic

- Working name row default: 4
- Size row default: 5
- Stock row default: 6
- Total qty row default: 7
- Country column default: I
- Clean qty = total qty minus ignored-country qty
- Blank country rows are not counted as ignored country rows
- DS/SS lookup from reference sheet
- DS loading default: 20%
- Multiplier detection:
  - `set of 4` multiplies by 4 and highlights red
  - `1 PACK = 100` multiplies by 100 and highlights red
  - unclear set/pack wording highlights orange and does not multiply

## Deploy

Upload all files to the GitHub repository root and set Streamlit main file path to:

```text
app.py
```
