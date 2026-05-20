# Excel Formula Fusion V1.9.1 Stable

Streamlit app for formula-based Excel processing.

## What this version fixes

- Updates the README to match the current build.
- Reduces Streamlit Cloud memory pressure by loading workbooks in low-memory `read_only` mode for UI previews.
- Keeps heavy workbook editing only inside **Generate Excel Workbook**.
- Keeps stock-rate editing inside **Refresh / Update Rates**, so typing a price does not regenerate the workbook.
- Adds visible error diagnostics inside the app instead of failing silently.

## Included logic

- Working sheet and reference sheet selection.
- Auto-detected mapping with manual override.
- Clean quantity formula:
  - starts from original total quantity row
  - subtracts ignored country quantities, such as NZ
- DS/SS lookup formula from reference sheet.
- SQM formula from size and clean quantity.
- Stock/material multi-select.
- Saved stock rate memory.
- DS loading percentage, default 20%.
- Controlled quantity multiplier detection:
  - `set of 4` multiplies by 4 and highlights red
  - `1 PACK = 100` multiplies by 100 and highlights red
  - uncertain set/pack wording highlights orange but does not multiply
- `Stock SQM Summary` sheet.
- `Qty Multiplier Audit` sheet.

## How to deploy to Streamlit Cloud

Upload these items to the root of your GitHub repository:

```text
app.py
requirements.txt
README.md
.streamlit/config.toml
```

In Streamlit Cloud, set:

```text
Main file path: app.py
```

## How to use

1. Open the Streamlit app.
2. Upload the Excel workbook.
3. Confirm the mapping.
4. Press **Apply Mapping** only if you changed rows/columns.
5. Pick one or more stock/material names.
6. Enter rates.
7. Press **Refresh / Update Rates**.
8. Press **Generate Excel Workbook**.
9. Download the generated workbook.
10. Check the red/orange multiplier audit before production.

## Important warning

Do not trust automatic multiplier detection blindly. Red means the app multiplied. Orange means the app is warning you but did not multiply.
