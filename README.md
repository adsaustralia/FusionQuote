# Excel Formula Fusion — Dynamic V1.7

Streamlit app for formula-based Excel automation.

## V1.7 fixes

- Fixed file uploader/button visibility by removing the black uploader styling.
- Faster workbook generation: no repeated workbook load inside every item column loop.
- Stock/material selector supports multiple stocks.
- Stock rates are remembered in session and saved to `data/stock_rates_memory.json` when possible.
- Added stock rate memory JSON download/upload for backup and Streamlit Cloud redeploys.
- Added DS loading percentage, default 20%.
- Item price formula applies DS loading only when DS/SS lookup returns DS / Double Sided / D/S.
- Stock summary total price is summed from item price formulas, so DS loading is included correctly.

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

## Important notes

Streamlit reruns the script when widgets change. This app uses forms for mapping changes, but stock/rate widgets still rerun lightly. Heavy Excel export only runs when you click **Generate Excel Workbook**.

Stock rates can be backed up by downloading `stock_rates_memory.json`.
