import io
import json
import re
from copy import copy

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

st.set_page_config(page_title="Excel Formula Fusion", layout="wide")

st.markdown("""
<style>
html, body, [class*="css"] { color:#111827 !important; }
.stApp { background:#f7f7fb !important; }
section[data-testid="stSidebar"] { background:#ffffff !important; border-right:1px solid #e5e7eb; }
label, p, span, div, h1, h2, h3, h4 { color:#111827 !important; }
input, textarea { color:#111827 !important; background:#ffffff !important; border:1px solid #9ca3af !important; }
button { color:#111827 !important; background:#ffffff !important; border:1px solid #9ca3af !important; }
.stButton > button, .stDownloadButton > button { color:#111827 !important; background:#f97316 !important; border:1px solid #ea580c !important; font-weight:700 !important; }
.stDownloadButton > button p, .stButton > button p { color:#111827 !important; font-weight:700 !important; }
div[data-baseweb="select"] > div { color:#111827 !important; background:#ffffff !important; border-color:#9ca3af !important; }
div[data-baseweb="select"] span, div[data-baseweb="select"] div { color:#111827 !important; }
div[role="listbox"], ul[role="listbox"] { background:#ffffff !important; color:#111827 !important; }
div[role="option"] { color:#111827 !important; background:#ffffff !important; }
[data-testid="stFileUploader"] section { background:#ffffff !important; color:#111827 !important; border:1px dashed #9ca3af !important; }
[data-testid="stFileUploader"] * { color:#111827 !important; }
.card { background:#ffffff; padding:16px; border-radius:14px; border:1px solid #e5e7eb; box-shadow:0 1px 2px rgba(0,0,0,.06); margin-bottom:14px; }
.good { color:#166534 !important; font-weight:700; }
.warn { color:#b45309 !important; font-weight:700; }
.bad { color:#b91c1c !important; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# -------------------- Helpers --------------------
def norm(v):
    if v is None:
        return ""
    return str(v).strip()

def qsheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"

def col_to_int(col):
    if isinstance(col, int):
        return col
    return column_index_from_string(str(col).strip().upper())

def int_to_col(idx):
    return get_column_letter(int(idx))

def safe_col(col, fallback="A"):
    try:
        return int_to_col(col_to_int(col))
    except Exception:
        return fallback

def detect_sheet(sheetnames, preferred):
    for p in preferred:
        for s in sheetnames:
            if p.lower() in s.lower():
                return s
    return sheetnames[0]

def header_col(ws_values, row, keywords, default_col):
    for c in range(1, ws_values.max_column + 1):
        val = norm(ws_values.cell(row, c).value).upper()
        if any(k.upper() in val for k in keywords):
            return int_to_col(c)
    return default_col

def detect_item_cols(ws_formula, total_qty_row=7):
    cols = []
    for c in range(1, ws_formula.max_column + 1):
        v = ws_formula.cell(total_qty_row, c).value
        if v not in (None, ""):
            # Campaign/item columns normally have a formula or number in total qty row
            if c > 10:
                cols.append(c)
    if not cols:
        return "AC", int_to_col(ws_formula.max_column)
    return int_to_col(min(cols)), int_to_col(max(cols))

def detect_last_store_row(ws_values, country_col="I", start_row=8):
    cc = col_to_int(country_col)
    last = start_row
    for r in range(start_row, ws_values.max_row + 1):
        if norm(ws_values.cell(r, cc).value):
            last = r
    return last

def country_list(ws_values, country_col="I", start_row=8, end_row=None):
    cc = col_to_int(country_col)
    end_row = end_row or ws_values.max_row
    vals = []
    for r in range(start_row, end_row + 1):
        v = norm(ws_values.cell(r, cc).value).upper()
        if v and v not in vals:
            vals.append(v)
    return vals

def excel_array_constant(values):
    clean = [str(v).strip().upper().replace('"', '') for v in values if str(v).strip()]
    if not clean:
        clean = ["NZ"]
    return "{" + ",".join(f'"{v}"' for v in clean) + "}"

def build_clean_qty_formula(col, total_qty_row, store_start_row, store_end_row, country_col, ignore_countries):
    # Corrected logic: use original total from row 7, subtract only ignored-country row qty.
    ignored = excel_array_constant(ignore_countries)
    return (
        f'={col}${total_qty_row}-SUMPRODUCT(({col}${store_start_row}:{col}${store_end_row})*'
        f'(--ISNUMBER(MATCH(UPPER(${country_col}${store_start_row}:${country_col}${store_end_row}),{ignored},0))))'
    )

def build_lookup_formula(working_col, name_row, size_row, ref_sheet, ref_start_row, ref_end_row, ref_name_col, ref_size_col, return_col):
    sh = qsheet(ref_sheet)
    # If user sets same column for name and size, match against the same reference field twice only when both are needed.
    # This keeps the formula valid while allowing the user's C/C setting.
    return (
        f'=IFERROR(INDEX({sh}!${return_col}${ref_start_row}:${return_col}${ref_end_row},'
        f'MATCH(1,INDEX(({sh}!${ref_name_col}${ref_start_row}:${ref_name_col}${ref_end_row}={working_col}${name_row})*'
        f'({sh}!${ref_size_col}${ref_start_row}:${ref_size_col}${ref_end_row}={working_col}${size_row}),0),0)),"")'
    )

def build_sqm_formula(col, clean_qty_row, size_row):
    s = f'LOWER(SUBSTITUTE({col}${size_row}," ",""))'
    return (
        f'=IFERROR({col}${clean_qty_row}*VALUE(LEFT({s},FIND("x",{s})-1))/1000*'
        f'VALUE(MID({s},FIND("x",{s})+1,99))/1000,0)'
    )

def copy_row_style(ws, src_row, dst_row, start_col, end_col):
    for c in range(start_col, end_col + 1):
        src = ws.cell(src_row, c)
        dst = ws.cell(dst_row, c)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)

def preview(ws_values, max_rows=20, max_cols=20):
    cols = [int_to_col(i) for i in range(1, min(ws_values.max_column, max_cols) + 1)]
    data = []
    for r in range(1, min(ws_values.max_row, max_rows) + 1):
        row = {"Row": r}
        for c in range(1, min(ws_values.max_column, max_cols) + 1):
            row[int_to_col(c)] = ws_values.cell(r, c).value
        data.append(row)
    return pd.DataFrame(data)

# -------------------- UI --------------------
st.title("Excel Formula Fusion — HOKA Dynamic V1.5")
st.caption("Formula-based export. Uses calculated cell values for UI detection, but preserves workbook formulas/styles when exporting.")

uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx"])
json_upload = st.file_uploader("Optional: upload saved mapping JSON", type=["json"])

if not uploaded:
    st.info("Upload your workbook to begin.")
    st.stop()

raw = uploaded.read()
wb_formula = load_workbook(io.BytesIO(raw), data_only=False)
wb_values = load_workbook(io.BytesIO(raw), data_only=True)
sheets = wb_formula.sheetnames

saved = {}
if json_upload is not None:
    try:
        saved = json.loads(json_upload.read().decode("utf-8"))
        st.success("Mapping JSON loaded.")
    except Exception as e:
        st.error(f"Could not read JSON: {e}")

working_default = saved.get("working_sheet") or detect_sheet(sheets, ["DL ANZ ALLOCATION", "ALLOCATION"])
ref_default = saved.get("reference_sheet") or detect_sheet(sheets, ["PRINT DB", "PRINT"])

st.markdown('<div class="card">', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1:
    working_sheet = st.selectbox("Working sheet", sheets, index=sheets.index(working_default))
with c2:
    reference_sheet = st.selectbox("Reference sheet", sheets, index=sheets.index(ref_default))
st.markdown('</div>', unsafe_allow_html=True)

ws_f = wb_formula[working_sheet]
ws_v = wb_values[working_sheet]
ref_v = wb_values[reference_sheet]

start_col_auto, end_col_auto = detect_item_cols(ws_f, saved.get("total_qty_row", 7))
country_col_auto = header_col(ws_v, 6, ["COUNTRY"], "I")
store_end_auto = detect_last_store_row(ws_v, country_col_auto, 8)
ref_name_auto = header_col(ref_v, 11, ["ARTWORK NAME", "NAME"], "C")
ref_size_auto = header_col(ref_v, 11, ["FINISH SIZE", "SIZE"], "E")
ref_ds_auto = header_col(ref_v, 11, ["DS/SS", "SS", "DS"], "F")
ref_stock_auto = header_col(ref_v, 11, ["MATERIAL", "STOCK"], "G")

st.subheader("Mapping")
st.caption("Defaults are auto-detected. You can override them before export.")

with st.expander("Working sheet mapping", expanded=True):
    a,b,c,d = st.columns(4)
    with a:
        name_row = st.number_input("Name row", 1, 500, int(saved.get("name_row", 4)))
        size_row = st.number_input("Size row", 1, 500, int(saved.get("size_row", 5)))
    with b:
        stock_row = st.number_input("Stock/material row", 1, 500, int(saved.get("stock_row", 6)))
        total_qty_row = st.number_input("Original total qty row", 1, 500, int(saved.get("total_qty_row", 7)))
    with c:
        country_col = safe_col(st.text_input("Country column", saved.get("country_col", country_col_auto)))
        store_start_row = 8  # hidden by design: first store line after header/total rows
        store_end_row = int(saved.get("store_end_row", store_end_auto))
        st.write(f"Detected store rows used internally: **{store_start_row}:{store_end_row}**")
    with d:
        item_start_col = safe_col(st.text_input("First item column", saved.get("item_start_col", start_col_auto)))
        item_end_col = safe_col(st.text_input("Last item column", saved.get("item_end_col", end_col_auto)))

with st.expander("Reference sheet mapping", expanded=True):
    r1,r2,r3,r4 = st.columns(4)
    with r1:
        ref_start_row = st.number_input("Reference start row", 1, 1000, int(saved.get("ref_start_row", 12)))
        ref_end_row = st.number_input("Reference end row", 1, 5000, int(saved.get("ref_end_row", ref_v.max_row)))
    with r2:
        ref_name_col = safe_col(st.text_input("Reference name/artwork column", saved.get("ref_name_col", ref_name_auto)))
        ref_size_col = safe_col(st.text_input("Reference size column", saved.get("ref_size_col", ref_size_auto)))
    with r3:
        ref_ds_col = safe_col(st.text_input("Reference DS/SS column", saved.get("ref_ds_col", ref_ds_auto)))
        ref_stock_col = safe_col(st.text_input("Reference stock/material column", saved.get("ref_stock_col", ref_stock_auto)))
    with r4:
        ds_output_row = st.number_input("DS/SS output row", 1, 500, int(saved.get("ds_output_row", 168)))
        clean_qty_row = st.number_input("Clean qty output row", 1, 500, int(saved.get("clean_qty_row", 169)))
        sqm_output_row = st.number_input("SQM output row", 1, 500, int(saved.get("sqm_output_row", 170)))

with st.expander("Country + stock calculation", expanded=True):
    countries = country_list(ws_v, country_col, 8, store_end_row)
    default_ignore = saved.get("ignore_countries", ["NZ"] if "NZ" in countries else [])
    ignore_countries = st.multiselect("Countries to exclude from row 7 total qty", options=countries, default=[c for c in default_ignore if c in countries])

    # Stock names must come from data_only values, not raw formulas.
    stock_values = []
    for c in range(col_to_int(item_start_col), col_to_int(item_end_col) + 1):
        v = norm(ws_v.cell(int(stock_row), c).value)
        if v and v not in stock_values:
            stock_values.append(v)
    selected_stock = st.selectbox("Pick stock/material to calculate SQM and rate", [""] + stock_values, index=0)
    sqm_rate = st.number_input("Square metre rate for selected stock", min_value=0.0, value=float(saved.get("sqm_rate", 0.0)), step=0.1, format="%.2f")
    st.caption("Stock list is read from calculated values, so it shows material names instead of formula text.")

# Preview of current defaults
mapping = {
    "working_sheet": working_sheet, "reference_sheet": reference_sheet,
    "name_row": int(name_row), "size_row": int(size_row), "stock_row": int(stock_row), "total_qty_row": int(total_qty_row),
    "country_col": country_col, "store_start_row": store_start_row, "store_end_row": int(store_end_row),
    "item_start_col": item_start_col, "item_end_col": item_end_col,
    "ref_start_row": int(ref_start_row), "ref_end_row": int(ref_end_row),
    "ref_name_col": ref_name_col, "ref_size_col": ref_size_col, "ref_ds_col": ref_ds_col, "ref_stock_col": ref_stock_col,
    "ds_output_row": int(ds_output_row), "clean_qty_row": int(clean_qty_row), "sqm_output_row": int(sqm_output_row),
    "ignore_countries": ignore_countries, "selected_stock": selected_stock, "sqm_rate": sqm_rate,
}

st.subheader("Current mapping summary")
st.dataframe(pd.DataFrame([mapping]).T.rename(columns={0:"Value"}), use_container_width=True)

with st.expander("Workbook preview", expanded=False):
    st.dataframe(preview(ws_v, 18, 35), use_container_width=True, height=450)

# Formula preview
first_col = item_start_col
st.subheader("Formula preview")
clean_preview = build_clean_qty_formula(first_col, int(total_qty_row), store_start_row, int(store_end_row), country_col, ignore_countries)
ds_preview = build_lookup_formula(first_col, int(name_row), int(size_row), reference_sheet, int(ref_start_row), int(ref_end_row), ref_name_col, ref_size_col, ref_ds_col)
stock_preview = build_lookup_formula(first_col, int(name_row), int(size_row), reference_sheet, int(ref_start_row), int(ref_end_row), ref_name_col, ref_size_col, ref_stock_col)
sqm_preview = build_sqm_formula(first_col, int(clean_qty_row), int(size_row))
st.code("Clean Qty: " + clean_preview)
st.code("DS/SS: " + ds_preview)
st.code("Stock: " + stock_preview)
st.code("SQM: " + sqm_preview)

# JSON download
json_bytes = json.dumps(mapping, indent=2).encode("utf-8")
st.download_button("Download mapping JSON", data=json_bytes, file_name="excel_formula_fusion_mapping.json", mime="application/json")

# Export
if st.button("Generate formula workbook"):
    out_wb = load_workbook(io.BytesIO(raw), data_only=False)
    out_ws = out_wb[working_sheet]

    start_idx = col_to_int(item_start_col)
    end_idx = col_to_int(item_end_col)

    # Labels
    out_ws.cell(int(ds_output_row), max(1, start_idx-1)).value = "DS/SS"
    out_ws.cell(int(clean_qty_row), max(1, start_idx-1)).value = "Clean Qty"
    out_ws.cell(int(sqm_output_row), max(1, start_idx-1)).value = "SQM"

    # Copy visible style from nearby total row where possible
    for r in [int(ds_output_row), int(clean_qty_row), int(sqm_output_row)]:
        copy_row_style(out_ws, int(total_qty_row), r, start_idx, end_idx)

    for c in range(start_idx, end_idx + 1):
        col = int_to_col(c)
        # Skip columns that are not real item columns by checking row 7 value/formula
        if out_ws.cell(int(total_qty_row), c).value in (None, ""):
            continue
        out_ws.cell(int(ds_output_row), c).value = build_lookup_formula(col, int(name_row), int(size_row), reference_sheet, int(ref_start_row), int(ref_end_row), ref_name_col, ref_size_col, ref_ds_col)
        out_ws.cell(int(clean_qty_row), c).value = build_clean_qty_formula(col, int(total_qty_row), store_start_row, int(store_end_row), country_col, ignore_countries)
        out_ws.cell(int(sqm_output_row), c).value = build_sqm_formula(col, int(clean_qty_row), int(size_row))
        out_ws.cell(int(sqm_output_row), c).number_format = '0.00'

    # Summary sheet for selected stock
    if "Stock SQM Summary" in out_wb.sheetnames:
        del out_wb["Stock SQM Summary"]
    sum_ws = out_wb.create_sheet("Stock SQM Summary")
    sum_ws["A1"] = "Selected Stock"
    sum_ws["B1"] = selected_stock
    sum_ws["A2"] = "Rate per SQM"
    sum_ws["B2"] = sqm_rate
    sum_ws["A3"] = "Total SQM"
    sum_ws["A4"] = "Total Value"
    if selected_stock:
        sh = qsheet(working_sheet)
        sum_ws["B3"] = f'=SUMPRODUCT(--({sh}!${item_start_col}${stock_row}:${item_end_col}${stock_row}=B1),{sh}!${item_start_col}${sqm_output_row}:${item_end_col}${sqm_output_row})'
    else:
        sum_ws["B3"] = 0
    sum_ws["B4"] = "=B2*B3"
    sum_ws["B2"].number_format = '$#,##0.00'
    sum_ws["B3"].number_format = '0.00'
    sum_ws["B4"].number_format = '$#,##0.00'

    bio = io.BytesIO()
    out_wb.save(bio)
    bio.seek(0)
    st.success("Workbook generated. Download below.")
    st.download_button(
        "Download Excel workbook",
        data=bio.getvalue(),
        file_name="excel_formula_fusion_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
