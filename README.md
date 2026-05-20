# Excel Formula Fusion — Stable V1.7.1

Streamlit app for formula-based Excel workbook automation.

## Main fixes in V1.7.1

- safer app startup: no cached openpyxl workbook objects
- better error handling if workbook loading or export fails
- upload button and download button visibility kept light/high-contrast
- multi-stock/material selection
- stock rate memory with JSON backup/restore
- DS loading default 20%
- clean quantity = original total qty row minus ignored-country quantity
- does not require store qty start/end rows

## Deploy

Upload these files to the root of your GitHub repo:

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `README.md`

In Streamlit Cloud, set main file path:

```text
app.py
```

## Important

Streamlit reruns when widgets change. This version keeps mapping edits inside an Apply Mapping form to avoid exporting/processing on every small change.
