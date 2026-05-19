import io
import json
import re
from copy import copy
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

st.set_page_config(page_title="Excel Formula Fusion", layout="wide")

st.markdown("""
<style>
:root { color-scheme: light; }
html, body, [data-testid="stAppViewContainer"] { background: #f5f7fb !important; color: #111827 !important; }
[data-testid="stHeader"] { background: #f5f7fb !important; }
[data-testid="stSidebar"] { background: #ffffff !important; color: #111827 !important; border-right: 1px solid #d1d5db; }
[data-testid="stSidebar"] * { color: #111827 !important; }
.stMarkdown, .stText, label, p, span, div { color: #111827; }
input, textarea, select {
    color: #111827 !important;
    background-color: #ffffff !important;
    border: 1px solid #9ca3af !important;
}
div[data-baseweb="select"] > div {
    color: #111827 !important;
    background-color: #ffffff !important;
    border: 1px solid #9ca3af !important;
}
div[data-baseweb="select"] span, div[data-baseweb="select"] div {
    color: #111827 !important;
}
div[role="listbox"], ul[role="listbox"] {
    background-color: #ffffff !important;
    color: #111827 !important;
}
div[role="option"] { color: #111827 !important; background-color: #ffffff !important; }
div[role="option"]:hover { background-color: #e5e7eb !important; }
[data-testid="stFileUploader"] section {
    background-color: #ffffff !important;
    border: 1px dashed #6b7280 !important;
    color: #111827 !important;
}
[data-testid="stFileUploader"] * { color: #111827 !important; }
.stButton > button, .stDownloadButton > button {
    background-color: #ff7a00 !important;
    color: #111827 !important;
    border: 2px solid #111827 !important;
    font-weight: 800 !important;
    border-radius: 8px !important;
}
.stButton > button p, .stDownloadButton > button p,
.stButton > button span, .stDownloadButton > button span {
    color: #111827 !important;
    font-weight: 800 !important;
}
[data-testid="stMetricValue"], [data-testid="stMetricLabel"] { color: #111827 !important; }
.block-card {
    background: #ffffff;
    border: 1px solid #d1d5db;
    border-radius: 12px;
    padding: 14px;
    margin-bottom: 12px;
}
.good { color: #047857 !important; font-weight: 700; }
.warn { color: #b45309 !important; font-weight: 700; }
.bad { color: #b91c1c !important; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

COUNTRY_VALUES = {"AUS", "AU", "AUSTRALIA", "NZ", "NEW ZEALAND", "FIJI", "FJ", "SG", "SINGAPORE"}
SIZE_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:mm|cm)?\s*[xX×]\s*(\d+(?:\.\d+)?)", re.I)


def norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def safe_sheet_name(name: str) -> str:
    return name.replace("'", "''")


def quote_sheet(name: str) -> str:
    return f"'{safe_sheet_name(name)}'"


def parse_size_sqm(size_text: str) -> float:
    """Returns square metres for simple W x H sizes. Defaults to 0 if not parseable."""
    text = norm(size_text).replace(",", "")
    m = SIZE_RE.search(text)
    if not m:
        dia = re.search(r"(?:dia|diameter|ø|Ø)\s*(\d+(?:\.\d+)?)|(?:\d+(?:\.\d+)?)\s*(?:mm\s*)?(?:dia|diameter|ø|Ø)", text, re.I)
        if dia:
            nums = re.findall(r"\d+(?:\.\d+)?", text)
            if nums:
                d = float(nums[0]) / 1000.0
                return 3.141592653589793 * (d / 2) ** 2
        return 0.0
    w = float(m.group(1))
    h = float(m.group(2))
    # If text says cm, convert cm to m. Otherwise assume mm.
    if "cm" in text.lower() and "mm" not in text.lower():
        return (w / 100.0) * (h / 100.0)
    return (w / 1000.0) * (h / 1000.0)


def sheet_to_dataframe(ws, max_rows=40, max_cols=40) -> pd.DataFrame:
    max_r = min(ws.max_row, max_rows)
    max_c = min(ws.max_column, max_cols)
    rows = []
    for r in range(1, max_r + 1):
        row = {"Row": r}
        for c in range(1, max_c + 1):
            row[get_column_letter(c)] = ws.cell(r, c).value
        rows.append(row)
    return pd.DataFrame(rows)


def detect_country_col(ws) -> str:
    best_col = 1
    best_score = -1
    for c in range(1, ws.max_column + 1):
        score = 0
        for r in range(1, min(ws.max_row, 250) + 1):
            val = norm(ws.cell(r, c).value).upper()
            if val in COUNTRY_VALUES:
                score += 1
        if score > best_score:
            best_score = score
            best_col = c
    return get_column_letter(best_col)


def detect_item_columns(ws, name_row: int, size_row: int) -> Tuple[str, str]:
    cols = []
    for c in range(1, ws.max_column + 1):
        name = norm(ws.cell(name_row, c).value)
        size = norm(ws.cell(size_row, c).value)
        if name and (SIZE_RE.search(size) or "dia" in size.lower() or "ø" in size.lower()):
            cols.append(c)
    if not cols:
        # fallback: populated cells after column J
        cols = [c for c in range(10, ws.max_column + 1) if norm(ws.cell(name_row, c).value) or norm(ws.cell(size_row, c).value)]
    if not cols:
        return "A", get_column_letter(ws.max_column)
    return get_column_letter(min(cols)), get_column_letter(max(cols))


def detect_name_size_qty_rows(ws) -> Tuple[int, int, int, int]:
    # Name row: text-dense row across the right side
    best_name, best_score = 1, -1
    best_size, size_score = 1, -1
    best_qty, qty_score = 1, -1
    for r in range(1, min(ws.max_row, 80) + 1):
        text_count = 0
        size_count = 0
        num_count = 0
        for c in range(10, ws.max_column + 1):
            val = ws.cell(r, c).value
            s = norm(val)
            if not s:
                continue
            if isinstance(val, (int, float)) or re.fullmatch(r"\d+(?:\.\d+)?", s):
                num_count += 1
            if SIZE_RE.search(s) or "dia" in s.lower() or "ø" in s.lower():
                size_count += 1
            if len(s) > 3 and not re.fullmatch(r"\d+(?:\.\d+)?", s):
                text_count += 1
        if text_count > best_score:
            best_score, best_name = text_count, r
        if size_count > size_score:
            size_score, best_size = size_count, r
        if num_count > qty_score:
            qty_score, best_qty = num_count, r
    stock_row = best_size + 1 if best_size + 1 <= ws.max_row else best_size
    return best_name, best_size, stock_row, best_qty


def detect_qty_rows_by_country(ws, country_col: str) -> Tuple[int, int]:
    ci = column_index_from_string(country_col)
    rows = []
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, ci).value):
            rows.append(r)
    if not rows:
        return 8, min(166, ws.max_row)
    return min(rows), max(rows)


def detect_reference_columns(ws) -> Dict[str, str]:
    # Detect by header row and content patterns. Defaults aligned to PRINT DB: C/E/F.
    candidates = {"name_col": "C", "size_col": "E", "ds_col": "F"}
    size_col_scores = {}
    ds_col_scores = {}
    text_col_scores = {}
    for c in range(1, ws.max_column + 1):
        size_score = ds_score = text_score = 0
        for r in range(1, min(ws.max_row, 250) + 1):
            s = norm(ws.cell(r, c).value)
            su = s.upper()
            if SIZE_RE.search(s) or "DIA" in su or "Ø" in su:
                size_score += 1
            if su in {"DS", "SS", "DOUBLE SIDED", "SINGLE SIDED"}:
                ds_score += 1
            if len(s) > 6 and not SIZE_RE.search(s):
                text_score += 1
        size_col_scores[c] = size_score
        ds_col_scores[c] = ds_score
        text_col_scores[c] = text_score
    if size_col_scores:
        candidates["size_col"] = get_column_letter(max(size_col_scores, key=size_col_scores.get))
    if ds_col_scores:
        candidates["ds_col"] = get_column_letter(max(ds_col_scores, key=ds_col_scores.get))
    # name col should be a text-heavy col that is not size or ds
    exclude = {column_index_from_string(candidates["size_col"]), column_index_from_string(candidates["ds_col"])}
    filtered = {c: s for c, s in text_col_scores.items() if c not in exclude}
    if filtered:
        candidates["name_col"] = get_column_letter(max(filtered, key=filtered.get))
    return candidates


def detect_reference_rows(ws, name_col: str, size_col: str, ds_col: str) -> Tuple[int, int]:
    nci = column_index_from_string(name_col)
    sci = column_index_from_string(size_col)
    dci = column_index_from_string(ds_col)
    rows = []
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, nci).value) and (norm(ws.cell(r, sci).value) or norm(ws.cell(r, dci).value)):
            # skip obvious header rows
            row_vals = " ".join(norm(ws.cell(r, c).value).lower() for c in (nci, sci, dci))
            if "name" in row_vals and "size" in row_vals:
                continue
            rows.append(r)
    if not rows:
        return 1, ws.max_row
    return min(rows), max(rows)


def build_clean_qty_formula(col: str, qty_start: int, qty_end: int, country_col: str, ignore_countries: List[str]) -> str:
    # Critical: also excludes blank country rows to prevent subtotal/total rows being counted again.
    country_range = f"${country_col}${qty_start}:${country_col}${qty_end}"
    qty_range = f"{col}${qty_start}:{col}${qty_end}"
    conditions = [f'(--(TRIM({country_range})<>""))']
    for country in ignore_countries:
        c = country.strip().upper().replace('"', '')
        if c:
            conditions.append(f'(--(UPPER(TRIM({country_range}))<>"{c}"))')
    return f'=SUMPRODUCT(({qty_range})*' + "*".join(conditions) + ")"


def build_ds_formula(col: str, name_row: int, size_row: int, ref_sheet: str, ref_start: int, ref_end: int, ref_name_col: str, ref_size_col: str, ref_ds_col: str) -> str:
    sh = quote_sheet(ref_sheet)
    return (
        f'=IFERROR(INDEX({sh}!${ref_ds_col}${ref_start}:${ref_ds_col}${ref_end},'
        f'MATCH(1,INDEX(({sh}!${ref_name_col}${ref_start}:${ref_name_col}${ref_end}={col}${name_row})*'
        f'({sh}!${ref_size_col}${ref_start}:${ref_size_col}${ref_end}={col}${size_row}),0),0)),"")'
    )


def copy_style_from_above(ws, target_row: int, source_row: int, start_col: int, end_col: int):
    for c in range(start_col, end_col + 1):
        src = ws.cell(source_row, c)
        dst = ws.cell(target_row, c)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)


def apply_label_cell(ws, row: int, label: str):
    ws.cell(row, 1).value = label
    ws.cell(row, 1).font = Font(bold=True, color="FFFFFF")
    ws.cell(row, 1).fill = PatternFill("solid", fgColor="1F4E78")
    ws.cell(row, 1).alignment = Alignment(horizontal="left")


def build_workbook(uploaded_bytes: bytes, settings: Dict, selected_stocks: List[str], stock_rates: Dict[str, float]) -> bytes:
    wb = load_workbook(io.BytesIO(uploaded_bytes))
    ws = wb[settings["working_sheet"]]

    start_col = column_index_from_string(settings["item_start_col"])
    end_col = column_index_from_string(settings["item_end_col"])
    qty_start = int(settings["qty_start_row"])
    qty_end = int(settings["qty_end_row"])
    ignore = [c.strip().upper() for c in settings["ignore_countries"] if c.strip()]

    ds_row = int(settings["ds_output_row"])
    clean_qty_row = int(settings["clean_qty_output_row"])
    sqm_row = int(settings["sqm_output_row"])

    # Style output rows from qty row where possible.
    for row in [ds_row, clean_qty_row, sqm_row]:
        copy_style_from_above(ws, row, int(settings["qty_row"]), start_col, end_col)

    apply_label_cell(ws, ds_row, "DS/SS Lookup")
    apply_label_cell(ws, clean_qty_row, "Clean Qty excl. ignored countries")
    apply_label_cell(ws, sqm_row, "Clean SQM")

    # Write formulas across item columns.
    for c in range(start_col, end_col + 1):
        col = get_column_letter(c)
        ws.cell(ds_row, c).value = build_ds_formula(
            col=col,
            name_row=int(settings["name_row"]),
            size_row=int(settings["size_row"]),
            ref_sheet=settings["reference_sheet"],
            ref_start=int(settings["ref_start_row"]),
            ref_end=int(settings["ref_end_row"]),
            ref_name_col=settings["ref_name_col"],
            ref_size_col=settings["ref_size_col"],
            ref_ds_col=settings["ref_ds_col"],
        )
        ws.cell(clean_qty_row, c).value = build_clean_qty_formula(
            col=col,
            qty_start=qty_start,
            qty_end=qty_end,
            country_col=settings["country_col"],
            ignore_countries=ignore,
        )
        size_text = ws.cell(int(settings["size_row"]), c).value
        sqm_each = parse_size_sqm(size_text)
        if sqm_each > 0:
            ws.cell(sqm_row, c).value = f"={col}${clean_qty_row}*{sqm_each:.8f}"
        else:
            ws.cell(sqm_row, c).value = f"=0"
        ws.cell(sqm_row, c).number_format = "0.00"

    # Stock summary sheet
    summary_name = "Formula Fusion Summary"
    if summary_name in wb.sheetnames:
        del wb[summary_name]
    sh = wb.create_sheet(summary_name)
    sh.append(["Stock Name", "Rate per SQM", "Total SQM", "Total Price", "Formula Notes"])
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sh[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    work_q = quote_sheet(settings["working_sheet"])
    stock_row = int(settings["stock_row"])
    start = settings["item_start_col"]
    end = settings["item_end_col"]
    for idx, stock in enumerate(selected_stocks, start=2):
        clean_stock = stock.replace('"', '""')
        sh.cell(idx, 1).value = stock
        sh.cell(idx, 2).value = float(stock_rates.get(stock, 0.0))
        sh.cell(idx, 2).number_format = "$#,##0.00"
        # Sum item SQM row where stock row equals selected stock.
        sh.cell(idx, 3).value = f'=SUMPRODUCT(({work_q}!${start}${sqm_row}:${end}${sqm_row})*(--({work_q}!${start}${stock_row}:${end}${stock_row}="{clean_stock}")))'
        sh.cell(idx, 3).number_format = "0.00"
        sh.cell(idx, 4).value = f"=B{idx}*C{idx}"
        sh.cell(idx, 4).number_format = "$#,##0.00"
        sh.cell(idx, 5).value = "Uses clean SQM row and stock row match"
    total_row = len(selected_stocks) + 3
    sh.cell(total_row, 1).value = "TOTAL"
    sh.cell(total_row, 1).font = Font(bold=True)
    sh.cell(total_row, 3).value = f"=SUM(C2:C{len(selected_stocks)+1})" if selected_stocks else "=0"
    sh.cell(total_row, 4).value = f"=SUM(D2:D{len(selected_stocks)+1})" if selected_stocks else "=0"
    sh.cell(total_row, 3).number_format = "0.00"
    sh.cell(total_row, 4).number_format = "$#,##0.00"
    for col in range(1, 6):
        sh.column_dimensions[get_column_letter(col)].width = [32, 16, 16, 16, 42][col-1]

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def get_unique_stocks(ws, stock_row: int, start_col: str, end_col: str) -> List[str]:
    stocks = []
    for c in range(column_index_from_string(start_col), column_index_from_string(end_col) + 1):
        s = norm(ws.cell(stock_row, c).value)
        if s and s not in stocks:
            stocks.append(s)
    return stocks


st.title("Excel Formula Fusion — V1.4")
st.caption("Formula-based Excel export with smart mapping, country exclusion, DS/SS lookup, and stock SQM/rate summary.")

uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx"])
config_upload = st.file_uploader("Optional: upload saved mapping JSON", type=["json"], key="mapping_json")

if not uploaded:
    st.info("Upload the workbook first.")
    st.stop()

uploaded_bytes = uploaded.getvalue()
wb_preview = load_workbook(io.BytesIO(uploaded_bytes), data_only=False)
sheets = wb_preview.sheetnames

saved_config = {}
if config_upload:
    try:
        saved_config = json.loads(config_upload.getvalue().decode("utf-8"))
        st.success("Mapping JSON loaded.")
    except Exception as e:
        st.error(f"Could not read mapping JSON: {e}")

col1, col2 = st.columns(2)
with col1:
    default_working = saved_config.get("working_sheet") or ("DL ANZ ALLOCATION" if "DL ANZ ALLOCATION" in sheets else sheets[0])
    working_sheet = st.selectbox("Working sheet", sheets, index=sheets.index(default_working) if default_working in sheets else 0)
with col2:
    default_ref = saved_config.get("reference_sheet") or ("PRINT DB" if "PRINT DB" in sheets else sheets[0])
    reference_sheet = st.selectbox("Reference sheet", sheets, index=sheets.index(default_ref) if default_ref in sheets else 0)

ws = wb_preview[working_sheet]
ref_ws = wb_preview[reference_sheet]

auto_name_row, auto_size_row, auto_stock_row, auto_qty_row = detect_name_size_qty_rows(ws)
auto_country_col = detect_country_col(ws)
auto_qty_start, auto_qty_end = detect_qty_rows_by_country(ws, auto_country_col)
auto_item_start, auto_item_end = detect_item_columns(ws, auto_name_row, auto_size_row)
ref_cols = detect_reference_columns(ref_ws)
auto_ref_start, auto_ref_end = detect_reference_rows(ref_ws, ref_cols["name_col"], ref_cols["size_col"], ref_cols["ds_col"])

def getcfg(k, default):
    return saved_config.get(k, default)

st.subheader("1. Smart mapping defaults")
with st.expander("Auto-detected defaults — review before export", expanded=True):
    defaults_df = pd.DataFrame([
        ["Name row", auto_name_row],
        ["Size row", auto_size_row],
        ["Stock row", auto_stock_row],
        ["Original total qty row", auto_qty_row],
        ["Store qty start row", auto_qty_start],
        ["Store qty end row", auto_qty_end],
        ["Country column", auto_country_col],
        ["Item start column", auto_item_start],
        ["Item end column", auto_item_end],
        ["Reference name column", ref_cols["name_col"]],
        ["Reference size column", ref_cols["size_col"]],
        ["Reference DS/SS column", ref_cols["ds_col"]],
        ["Reference start row", auto_ref_start],
        ["Reference end row", auto_ref_end],
    ], columns=["Field", "Detected Value"])
    st.dataframe(defaults_df, use_container_width=True, hide_index=True)

with st.expander("Workbook preview", expanded=False):
    st.dataframe(sheet_to_dataframe(ws, max_rows=35, max_cols=45), use_container_width=True, hide_index=True)

st.subheader("2. Mapping setup")
left, mid, right = st.columns(3)
with left:
    name_row = st.number_input("Working name row", min_value=1, max_value=ws.max_row, value=int(getcfg("name_row", auto_name_row)))
    size_row = st.number_input("Working size row", min_value=1, max_value=ws.max_row, value=int(getcfg("size_row", auto_size_row)))
    stock_row = st.number_input("Working stock/material row", min_value=1, max_value=ws.max_row, value=int(getcfg("stock_row", auto_stock_row)))
    qty_row = st.number_input("Original total qty row (kept untouched)", min_value=1, max_value=ws.max_row, value=int(getcfg("qty_row", auto_qty_row)))
with mid:
    qty_start_row = st.number_input("Store qty start row", min_value=1, max_value=ws.max_row, value=int(getcfg("qty_start_row", auto_qty_start)))
    qty_end_row = st.number_input("Store qty end row", min_value=1, max_value=ws.max_row, value=int(getcfg("qty_end_row", auto_qty_end)))
    country_col = st.text_input("Country column", value=str(getcfg("country_col", auto_country_col))).upper().strip()
    ignore_str = st.text_input("Countries to ignore, comma separated", value=", ".join(getcfg("ignore_countries", ["NZ"])))
with right:
    item_start_col = st.text_input("Item start column", value=str(getcfg("item_start_col", auto_item_start))).upper().strip()
    item_end_col = st.text_input("Item end column", value=str(getcfg("item_end_col", auto_item_end))).upper().strip()
    ds_output_row = st.number_input("DS/SS output row", min_value=1, max_value=max(ws.max_row+20, 300), value=int(getcfg("ds_output_row", 168)))
    clean_qty_output_row = st.number_input("Clean qty output row", min_value=1, max_value=max(ws.max_row+20, 300), value=int(getcfg("clean_qty_output_row", 169)))
    sqm_output_row = st.number_input("Clean SQM output row", min_value=1, max_value=max(ws.max_row+20, 300), value=int(getcfg("sqm_output_row", 170)))

st.subheader("3. Reference lookup setup")
r1, r2, r3 = st.columns(3)
with r1:
    ref_name_col = st.text_input("Reference name column", value=str(getcfg("ref_name_col", ref_cols["name_col"]))).upper().strip()
    ref_size_col = st.text_input("Reference size column", value=str(getcfg("ref_size_col", ref_cols["size_col"]))).upper().strip()
with r2:
    ref_ds_col = st.text_input("Reference DS/SS column", value=str(getcfg("ref_ds_col", ref_cols["ds_col"]))).upper().strip()
    ref_start_row = st.number_input("Reference start row", min_value=1, max_value=ref_ws.max_row, value=int(getcfg("ref_start_row", auto_ref_start)))
with r3:
    ref_end_row = st.number_input("Reference end row", min_value=1, max_value=ref_ws.max_row, value=int(getcfg("ref_end_row", auto_ref_end)))

ignore_countries = [c.strip().upper() for c in ignore_str.split(",") if c.strip()]

st.subheader("4. Formula preview")
preview_col = item_start_col
try:
    clean_preview = build_clean_qty_formula(preview_col, int(qty_start_row), int(qty_end_row), country_col, ignore_countries)
    ds_preview = build_ds_formula(preview_col, int(name_row), int(size_row), reference_sheet, int(ref_start_row), int(ref_end_row), ref_name_col, ref_size_col, ref_ds_col)
    st.code(clean_preview, language="excel")
    st.code(ds_preview, language="excel")
    st.caption("Notice the clean qty formula excludes blank country rows as well as ignored countries. This prevents total/subtotal rows being counted again.")
except Exception as e:
    st.error(f"Formula preview error: {e}")

st.subheader("5. Stock SQM and rate summary")
try:
    stock_options = get_unique_stocks(ws, int(stock_row), item_start_col, item_end_col)
except Exception:
    stock_options = []
selected_stocks = st.multiselect("Pick stock/material names for summary", stock_options, default=stock_options[: min(5, len(stock_options))])
stock_rates = {}
if selected_stocks:
    cols = st.columns(3)
    for i, stock in enumerate(selected_stocks):
        with cols[i % 3]:
            stock_rates[stock] = st.number_input(f"Rate per SQM — {stock}", min_value=0.0, value=0.0, step=0.10, format="%.2f")
else:
    st.warning("No stocks selected. The workbook will still export DS/SS and clean qty formulas.")

settings = {
    "working_sheet": working_sheet,
    "reference_sheet": reference_sheet,
    "name_row": int(name_row),
    "size_row": int(size_row),
    "stock_row": int(stock_row),
    "qty_row": int(qty_row),
    "qty_start_row": int(qty_start_row),
    "qty_end_row": int(qty_end_row),
    "country_col": country_col,
    "ignore_countries": ignore_countries,
    "item_start_col": item_start_col,
    "item_end_col": item_end_col,
    "ds_output_row": int(ds_output_row),
    "clean_qty_output_row": int(clean_qty_output_row),
    "sqm_output_row": int(sqm_output_row),
    "ref_name_col": ref_name_col,
    "ref_size_col": ref_size_col,
    "ref_ds_col": ref_ds_col,
    "ref_start_row": int(ref_start_row),
    "ref_end_row": int(ref_end_row),
}

st.subheader("6. Export")
st.download_button(
    "Download mapping JSON",
    data=json.dumps(settings, indent=2).encode("utf-8"),
    file_name="excel_formula_fusion_mapping.json",
    mime="application/json",
)

st.markdown("<div class='block-card'><b>Current export behaviour:</b><br>Original qty row is untouched. Clean qty uses store rows only, excludes ignored countries, and excludes blank country rows. Stock summary uses Clean SQM row and selected stock names.</div>", unsafe_allow_html=True)

if st.button("Generate Excel Workbook"):
    try:
        output_bytes = build_workbook(uploaded_bytes, settings, selected_stocks, stock_rates)
        st.session_state["output_bytes"] = output_bytes
        st.success("Workbook generated. Use the download button below.")
    except Exception as e:
        st.error(f"Export failed: {e}")

if "output_bytes" in st.session_state:
    st.download_button(
        "Download Excel Workbook",
        data=st.session_state["output_bytes"],
        file_name="formula_fusion_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
