# Excel Formula Fusion – First Build

This Streamlit app uploads an Excel workbook and writes formula-based rows to a selected working sheet.

## Current logic

Default setup is based on `HOKA Clifton 11_MASTER DBDL.xlsx`:

- Working sheet: `DL ANZ ALLOCATION`
- Reference sheet: `PRINT DB`
- Working item name row: 4
- Working size row: 5
- Working original qty row: 7
- Store rows: 8:167
- Country column: I
- Item columns: AC:EU
- Reference artwork/name column: C
- Reference size column: E
- Reference DS/SS column: F

## Formula outputs

- Country-excluded quantity row, default row 169
- DS/SS lookup row, default row 168

The workbook keeps formulas in Excel, so small edits can still be made outside the Streamlit app.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```
