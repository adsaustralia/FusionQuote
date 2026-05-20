import io
import json
import re
from copy import copy
from dataclasses import dataclass, asdict
from typing import Dict, List, Tuple, Optional
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

st.set_page_config(page_title="Excel Formula Fusion", layout="wide")

st.markdown("""
<style>
html, body, [class*="css"] { color: #111827 !important; }
.stApp { background: #f6f8fb; }
section[data-testid="stSidebar"] { background: #ffffff !important; }
input, textarea, select { color: #111827 !important; background-color: #ffffff !important; }
div[data-baseweb="select"] > div { color: #111827 !important; background-color: #ffffff !important; border-color: #94a3b8 !important; }
div[data-baseweb="select"] span { color: #111827 !important; }
div[role="listbox"], div[role="option"] { color: #111827 !important; background-color: #ffffff !important; }
button, .stButton button, .stDownloadButton button { color: #ffffff !important; background-color: #0f172a !important; border: 1px solid #0f172a !important; font-weight: 700 !important; }
.stDownloadButton button p, .stButton button p { color: #ffffff !important; }
.stDownloadButton button:hover, .stButton button:hover { background-color: #f97316 !important; border-color: #f97316 !important; color: #111827 !important; }
.stDownloadButton button:hover p, .stButton button:hover p { color: #111827 !important; }
[data-testid="stFileUploader"] { background:#ffffff !important; border:1px solid #cbd5e1 !important; border-radius:12px !important; padding:8px !important; }
[data-testid="stFileUploader"] * { color: #111827 !important; }
[data-testid="stFileUploader"] button { background:#ffffff !important; color:#111827 !important; border:1px solid #94a3b8 !important; }
[data-testid="stFileUploader"] button p { color:#111827 !important; }
[data-testid="stFileUploaderFileName"], [data-testid="stFileUploaderFile"] * { color:#111827 !important; background:#ffffff !important; }
.block-container { padding-top: 1.5rem; }
.card { background: #ffffff; padding: 1rem; border-radius: 14px; border: 1px solid #dbe3ef; box-shadow: 0 1px 3px rgba(15,23,42,.08); }
.good { color: #047857; font-weight: 700; }
.warn { color: #b45309; font-weight: 700; }
.bad { color: #b91c1c; font-weight: 700; }
.small-note { color:#475569; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

COUNTRY_DEFAULTS = ["NZ"]
RATE_MEMORY_PATH = Path("data/stock_rates_memory.json")


def load_rate_memory() -> Dict[str, float]:
    if "rate_memory" in st.session_state:
        return dict(st.session_state.rate_memory)
    try:
        if RATE_MEMORY_PATH.exists():
            data = json.loads(RATE_MEMORY_PATH.read_text(encoding="utf-8"))
            st.session_state.rate_memory = {str(k): float(v) for k, v in data.items()}
            return dict(st.session_state.rate_memory)
    except Exception:
        pass
    st.session_state.rate_memory = {}
    return {}


def save_rate_memory(memory: Dict[str, float]) -> None:
    cleaned = {str(k): float(v) for k, v in memory.items() if str(k).strip()}
    st.session_state.rate_memory = cleaned
    try:
        RATE_MEMORY_PATH.parent.mkdir(parents=True, exist_ok=True)
        RATE_MEMORY_PATH.write_text(json.dumps(cleaned, indent=2, sort_keys=True), encoding="utf-8")
    except Exception:
        # Streamlit Cloud storage can be ephemeral/read-only in some deployments. Session memory still works.
        pass

@dataclass
class MappingConfig:
    working_sheet: str
    reference_sheet: str
    item_start_col: str = "AC"
    item_end_col: str = "IG"
    name_row: int = 4
    size_row: int = 5
    stock_row: int = 6
    total_qty_row: int = 7
    country_col: str = "I"
    clean_qty_output_row: int = 169
    ds_ss_output_row: int = 168
    sqm_output_row: int = 170
    price_output_row: int = 171
    ds_loading_percent: float = 20.0
    ref_name_col: str = "C"
    ref_size_col: str = "E"
    ref_ds_col: str = "F"
    ref_stock_col: str = "G"
    ref_start_row: int = 1
    ref_end_row: int = 300
    ignore_countries: Tuple[str, ...] = ("NZ",)


def norm_text(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def safe_sheet_quote(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def col_range(start_col: str, end_col: str) -> List[str]:
    start = column_index_from_string(start_col.upper())
    end = column_index_from_string(end_col.upper())
    if end < start:
        start, end = end, start
    return [get_column_letter(i) for i in range(start, end + 1)]


def sheet_preview_df(ws_values, max_rows=30, max_cols=30) -> pd.DataFrame:
    rows = []
    for r in range(1, min(ws_values.max_row, max_rows) + 1):
        row = []
        for c in range(1, min(ws_values.max_column, max_cols) + 1):
            row.append(ws_values.cell(r, c).value)
        rows.append(row)
    cols = [get_column_letter(c) for c in range(1, min(ws_values.max_column, max_cols) + 1)]
    df = pd.DataFrame(rows, columns=cols)
    df.insert(0, "Row", list(range(1, len(df) + 1)))
    return df


@st.cache_data(show_spinner=False)
def get_workbook_bytes(uploaded_bytes: bytes):
    return uploaded_bytes


@st.cache_resource(show_spinner=False)
def load_workbooks(uploaded_bytes: bytes):
    # Formula workbook is used for export so existing formulas/styles are preserved.
    wb_formula = load_workbook(io.BytesIO(uploaded_bytes))
    # Value workbook is used for UI display because many cells contain formulas.
    wb_values = load_workbook(io.BytesIO(uploaded_bytes), data_only=True)
    return wb_formula, wb_values


def detect_country_col(ws_values) -> str:
    best_col, best_count = "I", -1
    country_terms = {"AUS", "AU", "AUSTRALIA", "NZ", "NEW ZEALAND", "FIJI", "SG", "SINGAPORE"}
    for c in range(1, ws_values.max_column + 1):
        count = 0
        for r in range(1, min(ws_values.max_row, 250) + 1):
            v = norm_text(ws_values.cell(r, c).value).upper()
            if v in country_terms:
                count += 1
        if count > best_count:
            best_count = count
            best_col = get_column_letter(c)
    return best_col


def detect_item_cols(ws_values, name_row=4, size_row=5, qty_row=7) -> Tuple[str, str]:
    size_pat = re.compile(r"\d+\s*(mm|cm)?\s*[x×]\s*\d+", re.I)
    candidates = []
    for c in range(1, ws_values.max_column + 1):
        name = norm_text(ws_values.cell(name_row, c).value)
        size = norm_text(ws_values.cell(size_row, c).value)
        qty = ws_values.cell(qty_row, c).value
        score = 0
        if name:
            score += 1
        if size_pat.search(size):
            score += 3
        if isinstance(qty, (int, float)) and qty > 0:
            score += 2
        if score >= 3:
            candidates.append(c)
    if not candidates:
        return "AC", get_column_letter(ws_values.max_column)
    return get_column_letter(min(candidates)), get_column_letter(max(candidates))


def detect_reference_columns(ws_values) -> Dict[str, str]:
    header_rows = range(1, min(ws_values.max_row, 30) + 1)
    found = {"name": "C", "size": "E", "ds": "F", "stock": "G"}
    for r in header_rows:
        for c in range(1, ws_values.max_column + 1):
            text = norm_text(ws_values.cell(r, c).value).upper()
            col = get_column_letter(c)
            if text in {"ARTWORK", "ARTWORK NAME", "NAME", "DESCRIPTION", "ITEM"}:
                found["name"] = col
            if "FINISH SIZE" in text or text == "SIZE":
                found["size"] = col
            if "DS" in text and "SS" in text:
                found["ds"] = col
            if "MATERIAL" in text or "STOCK" in text:
                found["stock"] = col
    return found


def detect_reference_rows(ws_values, ref_name_col="C", ref_size_col="E") -> Tuple[int, int]:
    name_idx = column_index_from_string(ref_name_col)
    size_idx = column_index_from_string(ref_size_col)
    rows = []
    for r in range(1, ws_values.max_row + 1):
        if norm_text(ws_values.cell(r, name_idx).value) and norm_text(ws_values.cell(r, size_idx).value):
            rows.append(r)
    if not rows:
        return 1, ws_values.max_row
    return min(rows), max(rows)


def build_ignore_array(ignore_countries: List[str]) -> str:
    cleaned = [c.strip().upper().replace('"', '') for c in ignore_countries if c.strip()]
    if not cleaned:
        cleaned = ["NZ"]
    if len(cleaned) == 1:
        return f'"{cleaned[0]}"'
    return "{" + ",".join([f'"{c}"' for c in cleaned]) + "}"


def build_clean_qty_formula(col: str, cfg: MappingConfig) -> str:
    # Correct logic: clean qty = original total qty row - sum of quantities where country is ignored.
    # Blank country rows are automatically ignored by SUMIF.
    ignores = [c for c in cfg.ignore_countries if c.strip()]
    if len(ignores) == 1:
        return f'={col}${cfg.total_qty_row}-SUMIF(${cfg.country_col}:${cfg.country_col},"{ignores[0].upper()}",{col}:{col})'
    arr = build_ignore_array(list(ignores))
    return f'={col}${cfg.total_qty_row}-SUM(SUMIF(${cfg.country_col}:${cfg.country_col},{arr},{col}:{col}))'


def build_ds_ss_formula(col: str, cfg: MappingConfig) -> str:
    rs = safe_sheet_quote(cfg.reference_sheet)
    return (
        f'=IFERROR(INDEX({rs}!${cfg.ref_ds_col}${cfg.ref_start_row}:${cfg.ref_ds_col}${cfg.ref_end_row},'
        f'MATCH(1,INDEX(({rs}!${cfg.ref_name_col}${cfg.ref_start_row}:${cfg.ref_name_col}${cfg.ref_end_row}={col}${cfg.name_row})*'
        f'({rs}!${cfg.ref_size_col}${cfg.ref_start_row}:${cfg.ref_size_col}${cfg.ref_end_row}={col}${cfg.size_row}),0),0)),"")'
    )


def build_sqm_formula(col: str, cfg: MappingConfig) -> str:
    # Parses common W x H mm sizes, strips spaces and 'mm'. Returns m2 * clean qty.
    size_cell = f'{col}${cfg.size_row}'
    qty_cell = f'{col}${cfg.clean_qty_output_row}'
    cleaned = f'LOWER(SUBSTITUTE(SUBSTITUTE({size_cell}," ",""),"mm",""))'
    width = f'VALUE(LEFT({cleaned},FIND("x",SUBSTITUTE({cleaned},"×","x"))-1))'
    height = f'VALUE(MID(SUBSTITUTE({cleaned},"×","x"),FIND("x",SUBSTITUTE({cleaned},"×","x"))+1,99))'
    return f'=IFERROR(({width}*{height}/1000000)*{qty_cell},0)'


def build_price_formula(col: str, cfg: MappingConfig, rate_cell: str) -> str:
    # Increase price by DS loading when DS/SS lookup row returns DS / Double Sided variants.
    ds_cell = f'{col}${cfg.ds_ss_output_row}'
    loading = 1 + (float(cfg.ds_loading_percent) / 100.0)
    return f'=IFERROR({col}${cfg.sqm_output_row}*{rate_cell}*IF(OR(UPPER({ds_cell})="DS",UPPER({ds_cell})="DOUBLE SIDED",UPPER({ds_cell})="D/S"),{loading},1),0)'


def unique_stock_values(ws_values, cfg: MappingConfig) -> List[str]:
    stocks = []
    for col in col_range(cfg.item_start_col, cfg.item_end_col):
        v = norm_text(ws_values[f"{col}{cfg.stock_row}"].value)
        if v and v not in stocks:
            stocks.append(v)
    return sorted(stocks)


def copy_row_style(ws, source_row: int, target_row: int, start_col: str, end_col: str):
    for col in col_range(start_col, end_col):
        src = ws[f"{col}{source_row}"]
        dst = ws[f"{col}{target_row}"]
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.border:
            dst.border = copy(src.border)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.font:
            dst.font = copy(src.font)


def make_summary_sheet(wb, cfg: MappingConfig, selected_stocks: List[str], stock_rates: Dict[str, float]):
    if "Stock SQM Summary" in wb.sheetnames:
        del wb["Stock SQM Summary"]
    ws_sum = wb.create_sheet("Stock SQM Summary")
    headers = ["Stock / Material", "Total SQM", "Base Rate / SQM", "DS Loading %", "Total Price"]
    for c, h in enumerate(headers, 1):
        cell = ws_sum.cell(1, c, h)
        cell.fill = PatternFill("solid", fgColor="0F172A")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    stock_range = f"'{cfg.working_sheet}'!${cfg.item_start_col}${cfg.stock_row}:${cfg.item_end_col}${cfg.stock_row}"
    sqm_range = f"'{cfg.working_sheet}'!${cfg.item_start_col}${cfg.sqm_output_row}:${cfg.item_end_col}${cfg.sqm_output_row}"
    price_range = f"'{cfg.working_sheet}'!${cfg.item_start_col}${cfg.price_output_row}:${cfg.item_end_col}${cfg.price_output_row}"
    for i, stock in enumerate(selected_stocks, 2):
        ws_sum.cell(i, 1, stock)
        ws_sum.cell(i, 2, f'=SUMIF({stock_range},A{i},{sqm_range})')
        ws_sum.cell(i, 3, float(stock_rates.get(stock, 0.0)))
        ws_sum.cell(i, 4, float(cfg.ds_loading_percent) / 100.0)
        # Total price is summed from item formulas so DS items get the 20% loading correctly.
        ws_sum.cell(i, 5, f'=SUMIF({stock_range},A{i},{price_range})')
    widths = [38, 16, 16, 14, 16]
    for col in range(1, 6):
        ws_sum.column_dimensions[get_column_letter(col)].width = widths[col-1]
    for r in range(2, 2 + len(selected_stocks)):
        ws_sum.cell(r, 2).number_format = '0.00'
        ws_sum.cell(r, 3).number_format = '$#,##0.00'
        ws_sum.cell(r, 4).number_format = '0%'
        ws_sum.cell(r, 5).number_format = '$#,##0.00'
    ws_sum.freeze_panes = "A2"

def apply_formulas(uploaded_bytes: bytes, cfg: MappingConfig, selected_stocks: List[str], stock_rates: Dict[str, float]) -> bytes:
    wb = load_workbook(io.BytesIO(uploaded_bytes))
    wb_values = load_workbook(io.BytesIO(uploaded_bytes), data_only=True, read_only=True)
    ws_values = wb_values[cfg.working_sheet]
    ws = wb[cfg.working_sheet]
    selected_set = set(selected_stocks)
    stock_by_col = {col: norm_text(ws_values[f"{col}{cfg.stock_row}"].value) for col in col_range(cfg.item_start_col, cfg.item_end_col)}
    # Preserve visible style by copying from nearby total qty row.
    for target_row in [cfg.ds_ss_output_row, cfg.clean_qty_output_row, cfg.sqm_output_row, cfg.price_output_row]:
        copy_row_style(ws, cfg.total_qty_row, target_row, cfg.item_start_col, cfg.item_end_col)
    # Labels before item start.
    label_col_idx = max(1, column_index_from_string(cfg.item_start_col) - 1)
    label_col = get_column_letter(label_col_idx)
    labels = {
        cfg.ds_ss_output_row: "DS/SS Lookup",
        cfg.clean_qty_output_row: "Clean Qty",
        cfg.sqm_output_row: "SQM",
        cfg.price_output_row: "Price",
    }
    for row, label in labels.items():
        cell = ws[f"{label_col}{row}"]
        cell.value = label
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="F97316")
        cell.alignment = Alignment(horizontal="right")
    # Write formula rows.
    for col in col_range(cfg.item_start_col, cfg.item_end_col):
        ws[f"{col}{cfg.ds_ss_output_row}"] = build_ds_ss_formula(col, cfg)
        ws[f"{col}{cfg.clean_qty_output_row}"] = build_clean_qty_formula(col, cfg)
        ws[f"{col}{cfg.sqm_output_row}"] = build_sqm_formula(col, cfg)
        # Price per item uses summary rate only if exact stock selected; otherwise 0.
        stock_val = stock_by_col.get(col, "")
        if stock_val in selected_set:
            rate_index = selected_stocks.index(stock_val) + 2
            rate_cell = f"'Stock SQM Summary'!$C${rate_index}"
            ws[f"{col}{cfg.price_output_row}"] = build_price_formula(col, cfg, rate_cell)
        else:
            ws[f"{col}{cfg.price_output_row}"] = "=0"
        ws[f"{col}{cfg.sqm_output_row}"].number_format = '0.00'
        ws[f"{col}{cfg.price_output_row}"].number_format = '$#,##0.00'
    make_summary_sheet(wb, cfg, selected_stocks, stock_rates)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


def config_from_json(data: dict, sheet_names: List[str]) -> Optional[MappingConfig]:
    try:
        payload = dict(data)
        if "ignore_countries" in payload and isinstance(payload["ignore_countries"], list):
            payload["ignore_countries"] = tuple(payload["ignore_countries"])
        cfg = MappingConfig(**payload)
        if cfg.working_sheet not in sheet_names or cfg.reference_sheet not in sheet_names:
            return None
        return cfg
    except Exception:
        return None


st.title("Excel Formula Fusion — Dynamic V1.7")
st.caption("Formula-based Excel automation with rate memory, DS loading, and faster export.")

uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx"])

if not uploaded_file:
    st.info("Upload your workbook to start.")
    st.stop()

uploaded_bytes = get_workbook_bytes(uploaded_file.getvalue())
wb_formula, wb_values = load_workbooks(uploaded_bytes)
sheet_names = wb_formula.sheetnames

if "cfg" not in st.session_state:
    default_working = "DL ANZ ALLOCATION" if "DL ANZ ALLOCATION" in sheet_names else sheet_names[0]
    default_ref = "PRINT DB" if "PRINT DB" in sheet_names else sheet_names[0]
    ws_work = wb_values[default_working]
    ws_ref = wb_values[default_ref]
    item_start, item_end = detect_item_cols(ws_work)
    ref_cols = detect_reference_columns(ws_ref)
    ref_start, ref_end = detect_reference_rows(ws_ref, ref_cols["name"], ref_cols["size"])
    st.session_state.cfg = MappingConfig(
        working_sheet=default_working,
        reference_sheet=default_ref,
        item_start_col=item_start,
        item_end_col=item_end,
        country_col=detect_country_col(ws_work),
        ref_name_col=ref_cols["name"],
        ref_size_col=ref_cols["size"],
        ref_ds_col=ref_cols["ds"],
        ref_stock_col=ref_cols["stock"],
        ref_start_row=ref_start,
        ref_end_row=ref_end,
    )

cfg: MappingConfig = st.session_state.cfg

with st.sidebar:
    st.header("Mapping Control")
    st.warning("Changes are batched. Edit values, then click Apply Mapping. This prevents full processing on every widget change.")
    json_upload = st.file_uploader("Upload mapping JSON", type=["json"], key="mapping_json")
    if json_upload is not None:
        try:
            loaded = json.load(json_upload)
            new_cfg = config_from_json(loaded, sheet_names)
            if new_cfg:
                st.session_state.cfg = new_cfg
                st.success("Mapping JSON loaded. Click Apply Mapping if you change anything else.")
                st.rerun()
            else:
                st.error("JSON loaded, but sheet names do not match this workbook.")
        except Exception as e:
            st.error(f"Could not read JSON: {e}")

with st.form("mapping_form"):
    st.subheader("1) Workbook + Mapping Setup")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        working_sheet = st.selectbox("Working Sheet", sheet_names, index=sheet_names.index(cfg.working_sheet))
        reference_sheet = st.selectbox("Reference Sheet", sheet_names, index=sheet_names.index(cfg.reference_sheet))
    with c2:
        name_row = st.number_input("Name Row", min_value=1, value=int(cfg.name_row), step=1)
        size_row = st.number_input("Size Row", min_value=1, value=int(cfg.size_row), step=1)
        stock_row = st.number_input("Stock / Material Row", min_value=1, value=int(cfg.stock_row), step=1)
        total_qty_row = st.number_input("Original Total Qty Row", min_value=1, value=int(cfg.total_qty_row), step=1)
    with c3:
        item_start_col = st.text_input("Item Start Column", value=cfg.item_start_col).upper().strip()
        item_end_col = st.text_input("Item End Column", value=cfg.item_end_col).upper().strip()
        country_col = st.text_input("Country Column", value=cfg.country_col).upper().strip()
        ignore_countries_str = st.text_input("Ignore Countries", value=", ".join(cfg.ignore_countries))
    with c4:
        ds_ss_output_row = st.number_input("DS/SS Output Row", min_value=1, value=int(cfg.ds_ss_output_row), step=1)
        clean_qty_output_row = st.number_input("Clean Qty Output Row", min_value=1, value=int(cfg.clean_qty_output_row), step=1)
        sqm_output_row = st.number_input("SQM Output Row", min_value=1, value=int(cfg.sqm_output_row), step=1)
        price_output_row = st.number_input("Price Output Row", min_value=1, value=int(cfg.price_output_row), step=1)
        ds_loading_percent = st.number_input("DS Loading %", min_value=0.0, max_value=500.0, value=float(cfg.ds_loading_percent), step=1.0)

    st.subheader("2) Reference Sheet Columns")
    r1, r2, r3, r4, r5, r6 = st.columns(6)
    with r1:
        ref_name_col = st.text_input("Ref Name Column", value=cfg.ref_name_col).upper().strip()
    with r2:
        ref_size_col = st.text_input("Ref Size Column", value=cfg.ref_size_col).upper().strip()
    with r3:
        ref_ds_col = st.text_input("Ref DS/SS Column", value=cfg.ref_ds_col).upper().strip()
    with r4:
        ref_stock_col = st.text_input("Ref Stock Column", value=cfg.ref_stock_col).upper().strip()
    with r5:
        ref_start_row = st.number_input("Ref Start Row", min_value=1, value=int(cfg.ref_start_row), step=1)
    with r6:
        ref_end_row = st.number_input("Ref End Row", min_value=1, value=int(cfg.ref_end_row), step=1)

    apply_mapping = st.form_submit_button("Apply Mapping")

if apply_mapping:
    st.session_state.cfg = MappingConfig(
        working_sheet=working_sheet,
        reference_sheet=reference_sheet,
        item_start_col=item_start_col,
        item_end_col=item_end_col,
        name_row=int(name_row),
        size_row=int(size_row),
        stock_row=int(stock_row),
        total_qty_row=int(total_qty_row),
        country_col=country_col,
        clean_qty_output_row=int(clean_qty_output_row),
        ds_ss_output_row=int(ds_ss_output_row),
        sqm_output_row=int(sqm_output_row),
        price_output_row=int(price_output_row),
        ds_loading_percent=float(ds_loading_percent),
        ref_name_col=ref_name_col,
        ref_size_col=ref_size_col,
        ref_ds_col=ref_ds_col,
        ref_stock_col=ref_stock_col,
        ref_start_row=int(ref_start_row),
        ref_end_row=int(ref_end_row),
        ignore_countries=tuple([x.strip().upper() for x in ignore_countries_str.split(",") if x.strip()]),
    )
    st.success("Mapping applied.")
    st.rerun()

cfg = st.session_state.cfg
ws_values = wb_values[cfg.working_sheet]

st.subheader("Current Mapping Summary")
summary = pd.DataFrame([asdict(cfg)]).T.reset_index()
summary.columns = ["Setting", "Value"]
st.dataframe(summary, use_container_width=True, hide_index=True)

st.subheader("Stock / Material Selection")
all_stocks = unique_stock_values(ws_values, cfg)
rate_memory = load_rate_memory()
selected_stocks = st.multiselect(
    "Pick one or more stock/materials for SQM and rate calculation",
    options=all_stocks,
    default=[s for s in all_stocks if s in rate_memory][:4] or (all_stocks[:1] if all_stocks else []),
)
stock_rates = {}
if selected_stocks:
    st.write("Enter square metre rate for each selected stock/material. Rates are remembered in this app session and saved to rate memory JSON.")
    rate_cols = st.columns(min(4, len(selected_stocks)))
    for i, stock in enumerate(selected_stocks):
        with rate_cols[i % len(rate_cols)]:
            default_rate = float(rate_memory.get(stock, 0.0))
            stock_rates[stock] = st.number_input(
                f"Rate: {stock[:35]}",
                min_value=0.0,
                value=default_rate,
                step=0.10,
                key=f"rate_{stock}",
            )
    save_rates_now = st.button("Save Rates Memory")
    if save_rates_now:
        updated_memory = dict(rate_memory)
        updated_memory.update(stock_rates)
        save_rate_memory(updated_memory)
        st.success("Rates saved to memory for this app/repo session.")
else:
    st.warning("Select at least one stock/material if you want stock SQM and rate summary.")

with st.expander("Rate memory backup / restore", expanded=False):
    st.caption("Use this if Streamlit Cloud restarts or you redeploy. Download the JSON and upload it later to restore rates.")
    rate_json_upload = st.file_uploader("Upload stock rate memory JSON", type=["json"], key="rate_memory_json")
    if rate_json_upload is not None:
        try:
            loaded_rates = json.load(rate_json_upload)
            cleaned_rates = {str(k): float(v) for k, v in loaded_rates.items()}
            save_rate_memory(cleaned_rates)
            st.success("Rate memory loaded. Refresh/select stocks to see saved rates.")
        except Exception as e:
            st.error(f"Could not load rate memory JSON: {e}")
    current_memory = load_rate_memory()
    st.download_button(
        "Download Stock Rate Memory JSON",
        data=json.dumps(current_memory, indent=2, sort_keys=True),
        file_name="stock_rates_memory.json",
        mime="application/json",
    )

st.subheader("Formula Preview")
preview_col = cfg.item_start_col
st.code(build_clean_qty_formula(preview_col, cfg), language="excel")
st.code(build_ds_ss_formula(preview_col, cfg), language="excel")
st.code(build_sqm_formula(preview_col, cfg), language="excel")

with st.expander("Workbook Preview", expanded=False):
    st.dataframe(sheet_preview_df(ws_values), use_container_width=True, hide_index=True)

st.subheader("Export")
export_col1, export_col2 = st.columns([1, 2])
with export_col1:
    generate = st.button("Generate Excel Workbook")
with export_col2:
    config_json = json.dumps(asdict(cfg), indent=2)
    st.download_button(
        "Download Mapping JSON",
        data=config_json,
        file_name="excel_formula_fusion_mapping.json",
        mime="application/json",
    )

if generate:
    # Save latest entered rates automatically before generating.
    updated_memory = load_rate_memory()
    updated_memory.update(stock_rates)
    save_rate_memory(updated_memory)
    with st.spinner("Generating formula workbook..."):
        output_bytes = apply_formulas(uploaded_bytes, cfg, selected_stocks, stock_rates)
    st.success("Workbook generated.")
    st.download_button(
        "Download Excel Workbook",
        data=output_bytes,
        file_name="excel_formula_fusion_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
