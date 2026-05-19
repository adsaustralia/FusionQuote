# Excel Formula Fusion — V1.4

Streamlit app for formula-based Excel automation.

## What this version fixes

1. **Quantity double-counting fixed**
   - Clean quantity formula now excludes ignored countries such as NZ.
   - It also excludes rows where the country cell is blank, preventing total/subtotal rows from being counted again.

   Example formula:

   ```excel
   =SUMPRODUCT((AC$8:AC$166)*(--(TRIM($I$8:$I$166)<>""))*(--(UPPER(TRIM($I$8:$I$166))<>"NZ")))
   ```

2. **Visible UI buttons and input text**
   - Fixed Streamlit theme and CSS so upload, export, and download controls are readable.

3. **Stock/material summary added**
   - Select stock/material names from the working sheet.
   - Enter rate per SQM.
   - App writes Clean SQM formulas and creates a `Formula Fusion Summary` sheet with stock-level SQM and price formulas.

## Files

Upload these files to your GitHub repository root:

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `README.md`

## Streamlit Cloud

Main file path:

```text
app.py
```

## Important logic

The original total quantity row is kept untouched.

Clean Qty is calculated from store rows only, based on the mapped country column:

- included: rows where country is not blank and not ignored
- excluded: NZ or other ignored countries
- excluded: blank country rows

This prevents summary rows or repeated totals from doubling the quantity.
