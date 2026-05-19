# Excel Formula Fusion — Dynamic V1.6

Streamlit app for formula-based Excel automation.

## What changed in V1.6

- Uses a form/apply button pattern so mapping changes do not trigger full processing on every click.
- Multi-select stock/material picker.
- Stock SQM and stock rate summary sheet.
- Clean quantity formula uses:
  - original total qty row
  - minus ignored country quantities
  - blank country rows are ignored automatically.
- Uses `data_only=True` workbook copy for UI values so stock/material row displays calculated values instead of formula text.
- Keeps a formula-preserving workbook copy for export.

## Deploy on Streamlit Cloud

Put these files in the GitHub repo root:

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `README.md`

Set Streamlit main file path to:

```text
app.py
```

## Recommended workflow

1. Upload workbook.
2. Review auto-detected mapping.
3. Change mappings if required.
4. Click **Apply Mapping** once.
5. Pick one or more stock/materials.
6. Enter rates.
7. Generate workbook.
8. Download Excel.

## Important formula logic

Clean quantity formula example:

```excel
=AC$7-SUMIF($I:$I,"NZ",AC:AC)
```

For multiple ignored countries:

```excel
=AC$7-SUM(SUMIF($I:$I,{"NZ","FIJI"},AC:AC))
```

This avoids needing store quantity start/end rows.
