from __future__ import annotations

import io
import json
import re
import traceback
from copy import copy
from typing import Dict, List, Tuple, Any

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter, column_index_from_string

APP_VERSION = "V1.9.2 Stable Upload Build"

st.set_page_config(page_title="Excel Formula Fusion", layout="wide")

st.markdown(
    """
<style>
.stApp { background: #f7f8fb; color: #111827; }
section[data-testid="stSidebar"], div[data-testid="stFileUploader"] { background: #ffffff !important; color: #111827 !important; }
label, p, span, div, input, textarea { color: #111827 !important; }
input, textarea, select { background-color: #ffffff !important; color: #111827 !important; }
div[data-baseweb="select"] > div { background-color: #ffffff !important; color: #111827 !important; border-color: #9ca3af !important; }
div[role="listbox"], div[data-baseweb="popover"] { background-color: #ffffff !important; color: #111827 !important; }
button, div.stButton > button, div.stDownloadButton > button { background-color: #f36f21 !important; color: #ffffff !important; border: 1px solid #c95512 !important; font-weight: 700 !important; }
div.stDownloadButton > button * , div.stButton > button * { color: #ffffff !important; }
.uploadedFile, .uploadedFile * { color: #111827 !important; background: #ffffff !important; }
.small-note { font-size: 0.9rem; color: #4b5563 !important; }
.safe-card { background:#ffffff; padding:1rem; border-radius:0.75rem; border:1px solid #e5e7eb; margin-bottom:0.75rem; }
</style>
""",
    unsafe_allow_html=True,
)

# ---------- helpers ----------

def norm_text(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def clean_sheet_name(name: str) -> str:
    return name.replace("'", "''")


def find_sheet(sheets: List[str], preferred: str, fallback_idx: int = 0) -> str:
    for s in sheets:
        if s.strip().lower() == preferred.strip().lower():
            return s
    if sheets:
        return sheets[min(fallback_idx, len(sheets)-1)]
    return ""


def safe_int(value: Any, default: int) -> int:
    try:
        return int(value)
    except Exception:
        return default


def safe_col(value: str, default: str) -> str:
    v = (value or "").strip().upper()
    if re.fullmatch(r"[A-Z]{1,3}", v):
        return v
    return default


def detect_item_columns(ws, name_row: int, size_row: int, stock_row: int) -> Tuple[int, int]:
    cols = []
    for c in range(1, ws.max_column + 1):
        vals = [norm_text(ws.cell(name_row, c).value), norm_text(ws.cell(size_row, c).value), norm_text(ws.cell(stock_row, c).value)]
        if any(vals) and c > 1:
            # Avoid obvious left-side metadata columns with country/store lists.
            if c >= 10:
                cols.append(c)
    if not cols:
        return 1, ws.max_column
    return min(cols), max(cols)


def detect_country_col(ws) -> str:
    country_tokens = {"NZ", "AUS", "AU", "AUSTRALIA", "NEW ZEALAND"}
    best_col, best_count = 9, 0
    for c in range(1, min(ws.max_column, 40) + 1):
        count = 0
        for r in range(1, min(ws.max_row, 250) + 1):
            val = norm_text(ws.cell(r, c).value).upper()
            if val in country_tokens:
                count += 1
        if count > best_count:
            best_col, best_count = c, count
    return get_column_letter(best_col)


def detect_multiplier(name: str) -> Tuple[int, str, str]:
    """Return multiplier, status, reason. status: red/orange/none."""
    text = norm_text(name)
    t = text.upper()
    # Confident: set of N, set-of-N, set of 4.
    m = re.search(r"\bSET\s*(?:OF)?\s*(\d{1,3})\b", t)
    if m:
        n = int(m.group(1))
        if 1 < n <= 500:
            return n, "red", f"Detected set of {n}"
    # Confident: 1 pack = 100, 1 PACK=100, PACK = 100.
    m = re.search(r"\b(?:1\s*)?PACK\s*=\s*(\d{1,5})\b", t)
    if m:
        n = int(m.group(1))
        if 1 < n <= 10000:
            return n, "red", f"Detected pack = {n}"
    # Confident: pack of 100.
    m = re.search(r"\bPACK\s+OF\s+(\d{1,5})\b", t)
    if m:
        n = int(m.group(1))
        if 1 < n <= 10000:
            return n, "red", f"Detected pack of {n}"
    # Doubtful: wording contains set/pack but no clear multiplier.
    if " SET" in f" {t}" or "PACK" in t:
        return 1, "orange", "Contains set/pack wording but multiplier is unclear"
    return 1, "none", ""


def parse_size_to_sqm(size_text: str) -> float:
    """Best-effort size parser for UI summary only. Excel output uses formulas/values separately."""
    s = norm_text(size_text).lower().replace("×", "x")
    nums = re.findall(r"(\d+(?:\.\d+)?)", s.replace(",", ""))
    if len(nums) >= 2:
        w, h = float(nums[0]), float(nums[1])
        # Assume mm unless explicit cm/m.
        if "cm" in s and "mm" not in s:
            return (w / 100) * (h / 100)
        if re.search(r"\bm\b", s) and "mm" not in s and "cm" not in s:
            return w * h
        return (w / 1000) * (h / 1000)
    # round/dia/ø
    if len(nums) == 1 and ("dia" in s or "ø" in s or "round" in s):
        d = float(nums[0])
        if "cm" in s and "mm" not in s:
            d_m = d / 100
        elif re.search(r"\bm\b", s) and "mm" not in s and "cm" not in s:
            d_m = d
        else:
            d_m = d / 1000
        return 3.1415926535 * (d_m / 2) ** 2
    return 0.0


def array_constant(values: List[str]) -> str:
    escaped = [v.strip().upper().replace('"', '') for v in values if v.strip()]
    if not escaped:
        escaped = ["NZ"]
    return "{" + ",".join(f'"{v}"' for v in escaped) + "}"


def build_clean_qty_formula(col: str, total_qty_row: int, country_col: str, max_row: int, ignored: List[str], multiplier: int) -> str:
    arr = array_constant(ignored)
    # total qty minus all rows with ignored countries, then apply name multiplier.
    base = f'({col}${total_qty_row}-SUMPRODUCT(({col}$1:{col}${max_row})*(--ISNUMBER(MATCH(UPPER(${country_col}$1:${country_col}${max_row}),{arr},0)))))'
    if multiplier and multiplier != 1:
        return f"={base}*{multiplier}"
    return f"={base}"


def build_ds_formula(col: str, name_row: int, size_row: int, ref_sheet: str, ref_name_col: str, ref_size_col: str, ref_ds_col: str, ref_start_row: int, ref_end_row: int) -> str:
    rs = clean_sheet_name(ref_sheet)
    return (
        f'=IFERROR(INDEX(\'{rs}\'!${ref_ds_col}${ref_start_row}:${ref_ds_col}${ref_end_row},'
        f'MATCH(1,INDEX((\'{rs}\'!${ref_name_col}${ref_start_row}:${ref_name_col}${ref_end_row}={col}${name_row})*'
        f'(\'{rs}\'!${ref_size_col}${ref_start_row}:${ref_size_col}${ref_end_row}={col}${size_row}),0),0)),"")'
    )


def apply_header_style(cell):
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.font = Font(color="FFFFFF", bold=True)


def copy_row_style(ws, source_row: int, target_row: int, start_col: int, end_col: int):
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


def get_stock_options(ws_values, stock_row: int, start_col: int, end_col: int) -> List[str]:
    stocks = []
    for c in range(start_col, end_col + 1):
        s = norm_text(ws_values.cell(stock_row, c).value)
        if s and s not in stocks:
            stocks.append(s)
    return stocks


def update_rate_memory(selected_stocks: List[str], default_rate: float):
    if "rate_memory" not in st.session_state:
        st.session_state.rate_memory = {}
    with st.form("rate_update_form", clear_on_submit=False):
        st.markdown("### Stock rates")
        cols = st.columns(3)
        staged = {}
        for idx, stock in enumerate(selected_stocks):
            current = float(st.session_state.rate_memory.get(stock, default_rate))
            with cols[idx % 3]:
                staged[stock] = st.number_input(f"$ / sqm — {stock[:45]}", min_value=0.0, value=current, step=0.10, key=f"rate_{stock}")
        submitted = st.form_submit_button("Refresh / Update Rates")
        if submitted:
            st.session_state.rate_memory.update(staged)
            st.success("Rates updated in this session. Download JSON if you want a backup.")


def make_workbook(upload_bytes: bytes, cfg: Dict[str, Any], rate_memory: Dict[str, float]) -> bytes:
    wb = load_workbook(io.BytesIO(upload_bytes))
    wb_values = load_workbook(io.BytesIO(upload_bytes), data_only=True)
    ws = wb[cfg["working_sheet"]]
    ws_values = wb_values[cfg["working_sheet"]]

    name_row = cfg["name_row"]
    size_row = cfg["size_row"]
    stock_row = cfg["stock_row"]
    total_qty_row = cfg["total_qty_row"]
    country_col = cfg["country_col"]
    item_start_col = column_index_from_string(cfg["item_start_col"])
    item_end_col = column_index_from_string(cfg["item_end_col"])
    ds_row = cfg["ds_output_row"]
    clean_qty_row = cfg["clean_qty_output_row"]
    multiplier_row = cfg["multiplier_output_row"]
    sqm_row = cfg["sqm_output_row"]
    price_row = cfg["price_output_row"]
    ref_sheet = cfg["reference_sheet"]

    max_scan_row = max(ws.max_row, 250)
    red_fill = PatternFill("solid", fgColor="FFC7CE")
    orange_fill = PatternFill("solid", fgColor="FCE4D6")
    green_fill = PatternFill("solid", fgColor="C6EFCE")
    yellow_fill = PatternFill("solid", fgColor="FFF2CC")

    # Labels.
    ws.cell(ds_row, 1).value = "DS/SS Lookup"
    ws.cell(clean_qty_row, 1).value = "Clean Qty"
    ws.cell(multiplier_row, 1).value = "Qty Multiplier"
    ws.cell(sqm_row, 1).value = "SQM"
    ws.cell(price_row, 1).value = "Price"
    for r in [ds_row, clean_qty_row, multiplier_row, sqm_row, price_row]:
        apply_header_style(ws.cell(r, 1))

    copy_row_style(ws, total_qty_row, clean_qty_row, item_start_col, item_end_col)

    audit_rows = []
    selected_stocks = set(cfg.get("selected_stocks", []))
    ignored = cfg.get("ignored_countries", ["NZ"])
    ds_loading = float(cfg.get("ds_loading", 20.0)) / 100.0

    for c in range(item_start_col, item_end_col + 1):
        col = get_column_letter(c)
        name = norm_text(ws_values.cell(name_row, c).value or ws.cell(name_row, c).value)
        size = norm_text(ws_values.cell(size_row, c).value or ws.cell(size_row, c).value)
        stock = norm_text(ws_values.cell(stock_row, c).value or ws.cell(stock_row, c).value)
        multiplier, status, reason = detect_multiplier(name)

        ws.cell(ds_row, c).value = build_ds_formula(
            col, name_row, size_row, ref_sheet,
            cfg["ref_name_col"], cfg["ref_size_col"], cfg["ref_ds_col"],
            cfg["ref_start_row"], cfg["ref_end_row"]
        )
        ws.cell(clean_qty_row, c).value = build_clean_qty_formula(col, total_qty_row, country_col, max_scan_row, ignored, multiplier)
        ws.cell(multiplier_row, c).value = multiplier

        sqm_per = parse_size_to_sqm(size)
        if sqm_per:
            ws.cell(sqm_row, c).value = f"={col}${clean_qty_row}*{sqm_per:.6f}"
        else:
            ws.cell(sqm_row, c).value = ""

        rate = rate_memory.get(stock, 0.0)
        if stock in selected_stocks and rate:
            # price = sqm * rate * DS loading when DS
            ws.cell(price_row, c).value = f'=IF(UPPER({col}${ds_row})="DS",{col}${sqm_row}*{rate}*(1+{ds_loading}),{col}${sqm_row}*{rate})'
        else:
            ws.cell(price_row, c).value = ""

        if status == "red":
            for rr in [name_row, clean_qty_row, multiplier_row]:
                ws.cell(rr, c).fill = red_fill
        elif status == "orange":
            for rr in [name_row, multiplier_row]:
                ws.cell(rr, c).fill = orange_fill
        else:
            ws.cell(multiplier_row, c).fill = green_fill

        if stock in selected_stocks:
            ws.cell(stock_row, c).fill = yellow_fill

        if status != "none":
            audit_rows.append([col, name, size, stock, multiplier, status.upper(), reason])

    # Audit sheet.
    if "Qty Multiplier Audit" in wb.sheetnames:
        del wb["Qty Multiplier Audit"]
    aud = wb.create_sheet("Qty Multiplier Audit")
    headers = ["Column", "Name", "Size", "Stock", "Multiplier", "Flag", "Reason"]
    aud.append(headers)
    for cell in aud[1]:
        apply_header_style(cell)
    for row in audit_rows:
        aud.append(row)
        fill = red_fill if row[5] == "RED" else orange_fill
        for cell in aud[aud.max_row]:
            cell.fill = fill
    for col in range(1, len(headers) + 1):
        aud.column_dimensions[get_column_letter(col)].width = 24

    # Summary sheet.
    if "Stock SQM Summary" in wb.sheetnames:
        del wb["Stock SQM Summary"]
    sm = wb.create_sheet("Stock SQM Summary")
    sm.append(["Stock", "Rate / sqm", "Total SQM", "Estimated Price", "DS Loading %"])
    for cell in sm[1]:
        apply_header_style(cell)
    summary_row = 2
    wsname = clean_sheet_name(cfg["working_sheet"])
    for stock in cfg.get("selected_stocks", []):
        rate = float(rate_memory.get(stock, 0.0))
        sm.cell(summary_row, 1).value = stock
        sm.cell(summary_row, 2).value = rate
        sm.cell(summary_row, 5).value = cfg.get("ds_loading", 20.0)
        # Formula built using SUMIF over stock row.
        sm.cell(summary_row, 3).value = f'=SUMIF(\'{wsname}\'!${cfg["item_start_col"]}${stock_row}:${cfg["item_end_col"]}${stock_row},A{summary_row},\'{wsname}\'!${cfg["item_start_col"]}${sqm_row}:${cfg["item_end_col"]}${sqm_row})'
        sm.cell(summary_row, 4).value = f'=SUMIF(\'{wsname}\'!${cfg["item_start_col"]}${stock_row}:${cfg["item_end_col"]}${stock_row},A{summary_row},\'{wsname}\'!${cfg["item_start_col"]}${price_row}:${cfg["item_end_col"]}${price_row})'
        summary_row += 1
    for col in range(1, 6):
        sm.column_dimensions[get_column_letter(col)].width = 28

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ---------- UI ----------
st.title("Excel Formula Fusion")
st.caption(APP_VERSION)

if "rate_memory" not in st.session_state:
    st.session_state.rate_memory = {}
if "uploaded_bytes" not in st.session_state:
    st.session_state.uploaded_bytes = None
if "uploaded_name" not in st.session_state:
    st.session_state.uploaded_name = ""

with st.sidebar:
    st.header("1. Upload")
    uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx", "xlsm"], key="excel_upload")
    rate_json = st.file_uploader("Optional: upload stock rate memory JSON", type=["json"], key="rate_json")
    if rate_json is not None:
        try:
            st.session_state.rate_memory.update(json.loads(rate_json.getvalue().decode("utf-8")))
            st.success("Rate memory loaded.")
        except Exception as e:
            st.error(f"Could not load rate JSON: {e}")

if uploaded is not None:
    st.session_state.uploaded_bytes = uploaded.getvalue()
    st.session_state.uploaded_name = uploaded.name

if not st.session_state.uploaded_bytes:
    st.info("Upload an Excel workbook to start. The app will show sheet names before doing any heavy processing.")
    st.stop()

st.success(f"Uploaded: {st.session_state.uploaded_name}")
st.write(f"File size: {len(st.session_state.uploaded_bytes)/1024/1024:.2f} MB")

# Load only workbook metadata and values now. This is intentional and visible.
try:
    wb_meta = load_workbook(io.BytesIO(st.session_state.uploaded_bytes), read_only=True, data_only=False)
    sheet_names = wb_meta.sheetnames
    wb_meta.close()
except Exception:
    st.error("The workbook could not be opened. Check that it is a valid .xlsx/.xlsm file.")
    st.code(traceback.format_exc())
    st.stop()

st.markdown("### Workbook detected")
st.write("Sheets:", ", ".join(sheet_names))

working_default = find_sheet(sheet_names, "DL ANZ ALLOCATION", 0)
reference_default = find_sheet(sheet_names, "PRINT DB", 1 if len(sheet_names) > 1 else 0)

with st.form("mapping_form", clear_on_submit=False):
    st.markdown("### 2. Mapping")
    c1, c2 = st.columns(2)
    with c1:
        working_sheet = st.selectbox("Working sheet", sheet_names, index=sheet_names.index(working_default))
        name_row = st.number_input("Name row", min_value=1, value=4, step=1)
        size_row = st.number_input("Size row", min_value=1, value=5, step=1)
        stock_row = st.number_input("Stock/material row", min_value=1, value=6, step=1)
        total_qty_row = st.number_input("Original total quantity row", min_value=1, value=7, step=1)
        country_col = st.text_input("Country column", value="I").upper()
        ignored_text = st.text_input("Ignore countries, comma separated", value="NZ")
    with c2:
        reference_sheet = st.selectbox("Reference sheet", sheet_names, index=sheet_names.index(reference_default))
        ref_name_col = st.text_input("Reference name/artwork column", value="C").upper()
        ref_size_col = st.text_input("Reference size column", value="E").upper()
        ref_ds_col = st.text_input("Reference DS/SS column", value="F").upper()
        ref_stock_col = st.text_input("Reference stock/material column", value="G").upper()
        ref_start_row = st.number_input("Reference start row", min_value=1, value=12, step=1)
        ref_end_row = st.number_input("Reference end row", min_value=1, value=141, step=1)

    st.markdown("### 3. Output rows")
    o1, o2, o3, o4, o5 = st.columns(5)
    with o1:
        ds_output_row = st.number_input("DS/SS row", min_value=1, value=168, step=1)
    with o2:
        clean_qty_output_row = st.number_input("Clean qty row", min_value=1, value=169, step=1)
    with o3:
        multiplier_output_row = st.number_input("Multiplier row", min_value=1, value=170, step=1)
    with o4:
        sqm_output_row = st.number_input("SQM row", min_value=1, value=171, step=1)
    with o5:
        price_output_row = st.number_input("Price row", min_value=1, value=172, step=1)

    st.markdown("### 4. Item columns")
    auto_cols = st.checkbox("Auto-detect item start/end columns after Apply Mapping", value=True)
    a1, a2 = st.columns(2)
    with a1:
        manual_start_col = st.text_input("Manual item start column", value="AC").upper()
    with a2:
        manual_end_col = st.text_input("Manual item end column", value="IG").upper()

    ds_loading = st.number_input("DS loading %", min_value=0.0, max_value=500.0, value=20.0, step=1.0)
    default_rate = st.number_input("Default stock rate for new stocks", min_value=0.0, value=0.0, step=0.10)
    apply_clicked = st.form_submit_button("Apply Mapping / Refresh Stock List")

# Apply or initialise mapping.
if apply_clicked or "cfg" not in st.session_state:
    try:
        wb_values = load_workbook(io.BytesIO(st.session_state.uploaded_bytes), data_only=True, read_only=False)
        ws_values = wb_values[working_sheet]
        if auto_cols:
            start_idx, end_idx = detect_item_columns(ws_values, int(name_row), int(size_row), int(stock_row))
            item_start_col = get_column_letter(start_idx)
            item_end_col = get_column_letter(end_idx)
        else:
            item_start_col = safe_col(manual_start_col, "AC")
            item_end_col = safe_col(manual_end_col, "IG")
        stock_options = get_stock_options(ws_values, int(stock_row), column_index_from_string(item_start_col), column_index_from_string(item_end_col))
        wb_values.close()
        st.session_state.cfg = {
            "working_sheet": working_sheet,
            "reference_sheet": reference_sheet,
            "name_row": int(name_row),
            "size_row": int(size_row),
            "stock_row": int(stock_row),
            "total_qty_row": int(total_qty_row),
            "country_col": safe_col(country_col, "I"),
            "ignored_countries": [x.strip().upper() for x in ignored_text.split(",") if x.strip()],
            "ref_name_col": safe_col(ref_name_col, "C"),
            "ref_size_col": safe_col(ref_size_col, "E"),
            "ref_ds_col": safe_col(ref_ds_col, "F"),
            "ref_stock_col": safe_col(ref_stock_col, "G"),
            "ref_start_row": int(ref_start_row),
            "ref_end_row": int(ref_end_row),
            "ds_output_row": int(ds_output_row),
            "clean_qty_output_row": int(clean_qty_output_row),
            "multiplier_output_row": int(multiplier_output_row),
            "sqm_output_row": int(sqm_output_row),
            "price_output_row": int(price_output_row),
            "item_start_col": item_start_col,
            "item_end_col": item_end_col,
            "ds_loading": float(ds_loading),
            "stock_options": stock_options,
            "selected_stocks": st.session_state.get("selected_stocks", []),
        }
        st.success(f"Mapping applied. Item columns: {item_start_col} to {item_end_col}. Stocks found: {len(stock_options)}")
    except Exception:
        st.error("Mapping failed. Check row/column settings.")
        st.code(traceback.format_exc())

if "cfg" not in st.session_state:
    st.warning("Click Apply Mapping / Refresh Stock List after upload.")
    st.stop()

cfg = st.session_state.cfg

st.markdown("### Current mapping")
st.json({k: v for k, v in cfg.items() if k != "stock_options"}, expanded=False)

stock_options = cfg.get("stock_options", [])
selected_stocks = st.multiselect("Pick stock/material to calculate SQM and rate", stock_options, default=cfg.get("selected_stocks", []))
st.session_state.selected_stocks = selected_stocks
st.session_state.cfg["selected_stocks"] = selected_stocks

update_rate_memory(selected_stocks, float(default_rate))

rate_bytes = json.dumps(st.session_state.rate_memory, indent=2).encode("utf-8")
st.download_button("Download stock rate memory JSON", data=rate_bytes, file_name="stock_rate_memory.json", mime="application/json")

st.markdown("### Generate")
st.warning("Generating workbook only runs when you press this button. Typing rates or changing mapping will not generate the Excel file.")

if st.button("Generate Excel Workbook"):
    try:
        with st.spinner("Generating formula workbook..."):
            output_bytes = make_workbook(st.session_state.uploaded_bytes, st.session_state.cfg, st.session_state.rate_memory)
        st.success("Workbook generated.")
        st.download_button(
            "Download Excel Workbook",
            data=output_bytes,
            file_name=f"formula_fusion_{st.session_state.uploaded_name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception:
        st.error("Workbook generation failed. Technical details below.")
        st.code(traceback.format_exc())
