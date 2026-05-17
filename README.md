# Excel Formula Fusion — Dynamic V1

Streamlit app for generating formula-based Excel outputs while preserving the workbook structure.

## What this version does

- Upload an Excel workbook.
- Select a working sheet and a reference sheet.
- Dynamically configure rows and columns instead of relying on fixed formulas.
- Generate a clean quantity formula that excludes selected countries, such as NZ.
- Generate a DS/SS lookup formula from a reference sheet using name + size matching.
- Validate possible quantity range issues before export.
- Validate DS/SS exact, missing, and duplicate matches.
- Export a workbook with formulas, not static values.
- Download a mapping JSON so the same setup can be reused later.

## Default HOKA Mapping

The defaults are tuned for `HOKA Clifton 11_MASTER DBDL.xlsx`:

- Working sheet: `DL ANZ ALLOCATION`
- Reference sheet: `PRINT DB`
- Name row: `4`
- Size row: `5`
- Quantity rows: `8:166`
- Country column: `I`
- DS/SS output row: `168`
- Clean qty output row: `169`
- Reference name column: `C`
- Reference size column: `E`
- Reference DS/SS column: `F`

## Important warning

For the HOKA workbook, row `167` can repeat the total quantity. Do not include it in the store quantity range unless you confirm it is a real store row. Using `8:167` can double-count.

## Formula examples

Clean qty:

```excel
=SUMPRODUCT((AC$8:AC$166)*(--(UPPER($I$8:$I$166)<>"NZ")))
```

DS/SS lookup:

```excel
=IFERROR(INDEX('PRINT DB'!$F$12:$F$141,MATCH(1,INDEX(('PRINT DB'!$C$12:$C$141=AC$4)*('PRINT DB'!$E$12:$E$141=AC$5),0),0)),"")
```

## Deploy to Streamlit Cloud

1. Upload this repository to GitHub.
2. In Streamlit Cloud, create/update your app.
3. Set the main file path to:

```text
app.py
```

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Ruthless usage rule

Do not trust auto-detection blindly. Always check:

- Quantity range validation
- DS/SS match validation
- Formula preview

Then export.
