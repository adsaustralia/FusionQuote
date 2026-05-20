# Excel Formula Fusion V2.1 Recovery Stable

This is a recovery/stability build. It deliberately avoids heavy processing during upload.

## Important change

Upload only stores the workbook. You must press **Read workbook / detect sheets** before mapping appears. This prevents Streamlit Cloud from crashing while typing or uploading.

## Main features

- Upload workbook safely.
- Read sheet names only after button click.
- Select working and reference sheets.
- Stock/material list scans only from selected Start Column to End Column.
- Enter multiple stock rates in a clear table-style section.
- Save stock rates in session and download/upload JSON backup.
- Clean qty formula uses total qty row minus ignored-country rows.
- DS loading default is 20%.
- Detects confident multipliers like `set of 4` and `1 PACK = 100`; red flag and multiply.
- Orange flag for doubtful set/pack wording; does not multiply.
- Formula workbook is generated only after **Generate Excel Workbook**.

## Default HOKA mapping

- Working sheet: DL ANZ ALLOCATION
- Reference sheet: PRINT DB
- Name row: 4
- Size row: 5
- Stock row: 6
- Original total qty row: 7
- Country column: I
- Item columns: AC to IG
- DS/SS output row: 168
- Clean qty output row: 169
- Multiplier row: 170
- SQM row: 171
- Price row: 172

## Reference defaults

- Name/artwork column: C
- Size column: E
- DS/SS column: F
- Stock/material column: G
- Reference rows: 12 to 141

## GitHub / Streamlit Cloud

Upload these to your repository root:

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `README.md`

Main file path:

```text
app.py
```
