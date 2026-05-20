# Excel Formula Fusion V1.8

Streamlit app for formula-based Excel processing.

## New in V1.8

- Detects quantity multipliers in item/name text.
- Confident patterns are multiplied automatically and highlighted red.
- Suspicious patterns are highlighted orange and are not multiplied.
- Adds a `Qty Multiplier Audit` sheet for checking all flagged columns.

## Confident multiplier examples

- `set of 4` → multiplier 4
- `set x 4` → multiplier 4
- `1 PACK = 100` → multiplier 100
- `PACK OF 100` → multiplier 100

## Safety rule

If the app is not confident, it does not multiply. It only highlights the item orange for manual review.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud

Upload these files to your GitHub repository root:

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `README.md`

Set main file path to:

```text
app.py
```
