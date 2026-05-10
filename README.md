# Fusion Excel Reformat + Flat/Roll Packaging Calculator

Streamlit app for Fusion Excel allocation workbooks.

## What it does
- Upload one `.xlsx` workbook.
- Preview sheets in an Excel-like grid.
- Detect item columns, size/spec row, material row, side/print row, country column, and data rows.
- Calculates per item and per zone/store:
  - flat-pack item count
  - roll item count
  - total quantity
  - maximum roll width in mm
  - number of stores
- Treats Ferrous/Magnetic/Banner/Vinyl/Fabric as roll materials by default.
- Adds summary sheets without modifying original sheet layouts.
- Optional reformat export with frozen panes, widths, and filters applied to copied sheets.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Important
The app is rule-based. Fusion sheets are messy, so always review the detected rows/columns before exporting. Wrong row detection equals wrong packaging count.
