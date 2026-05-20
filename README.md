# Excel Formula Fusion V1.9.2 Stable Upload Build

This build is designed to stop the Streamlit app freezing or doing nothing immediately after upload.

## What changed

- Upload now confirms the file name and sheet names immediately.
- Workbook processing only runs after clicking **Generate Excel Workbook**.
- Stock rate changes are staged and saved only after clicking **Refresh / Update Rates**.
- Multi-stock rate entry is supported.
- DS loading defaults to 20%.
- Quantity logic uses total qty row minus ignored-country quantity.
- Name-based multipliers are detected:
  - `set of 4` multiplies by 4 and highlights red.
  - `1 PACK = 100` multiplies by 100 and highlights red.
  - doubtful set/pack wording highlights orange and does not multiply.
- The app reads display values from a `data_only=True` workbook where possible, so linked stock/material values show as values instead of formulas.

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

## Deployment

Upload these files to the root of your GitHub repository:

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `README.md`

In Streamlit Cloud, set the main file path to:

```text
app.py
```

## Important

Do not expect Streamlit Cloud to permanently store stock rates on the server. Use the JSON download/upload option to back up your rates.
