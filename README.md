# Excel Formula Fusion — Smart Dynamic V1.2

Streamlit app for generating formula-based Excel outputs while preserving workbook structure.

## What changed in V1.2

- Fixes invisible text/selectbox issue caused by dark sidebar CSS.
- Auto-detects likely working-sheet mappings:
  - name row
  - size row
  - active item start/end columns
  - country column
  - quantity start/end rows
  - output rows
- Auto-detects likely reference-sheet mappings:
  - reference start/end rows
  - artwork/name column
  - size column
  - DS/SS column
- Still lets the user override every mapping.
- Shows the detected defaults as JSON inside the sidebar.
- Keeps formulas dynamic and editable in Excel.

## Important usage rule

Do not trust auto-detection blindly. It is a setup assistant, not a brain replacement.

Always check:

1. Formula preview
2. Quantity range validation
3. DS/SS match validation
4. Output workbook in Excel

## Deploy to Streamlit Cloud

1. Upload `app.py`, `requirements.txt`, and `README.md` to GitHub root.
2. In Streamlit Cloud, set main file path to:

```text
app.py
```

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```
