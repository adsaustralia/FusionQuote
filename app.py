import io
import json
import re
from copy import copy
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string

st.set_page_config(page_title="Excel Formula Fusion", layout="wide")

st.markdown("""
<style>
.stApp { background:#f6f8fb; color:#111827; }
.block-container { padding-top:1.2rem; }
section[data-testid="stSidebar"] { background:#ffffff !important; }
[data-testid="stFileUploader"] { background:#ffffff !important; border:1px solid #cbd5e1 !important; border-radius:12px !important; padding:10px !important; }
[data-testid="stFileUploader"] * { color:#111827 !important; }
[data-testid="stFileUploader"] button { background:#ffffff !important; color:#111827 !important; border:1px solid #94a3b8 !important; }
[data-testid="stFileUploader"] button p { color:#111827 !important; }
input, textarea { color:#111827 !important; background:#ffffff !important; }
div[data-baseweb="select"] > div { color:#111827 !important; background:#ffffff !important; border-color:#94a3b8 !important; }
div[data-baseweb="select"] span { color:#111827 !important; }
div[role="listbox"], div[role="option"] { color:#111827 !important; background:#ffffff !important; }
.stButton button, .stDownloadButton button { color:#ffffff !important; background:#0f172a !important; border:1px solid #0f172a !important; font-weight:700 !important; }
.stButton button p, .stDownloadButton button p { color:#ffffff !important; }
.stButton button:hover, .stDownloadButton button:hover { background:#f97316 !important; color:#111827 !important; border-color:#f97316 !important; }
.stButton button:hover p, .stDownloadButton button:hover p { color:#111827 !important; }
</style>
""", unsafe_allow_html=True)

RATE_MEMORY_PATH = Path("data/stock_rates_memory.json")

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
    ds_ss_output_row: int = 168
    clean_qty_output_row: int = 169
    sqm_output_row: int = 170
    price_output_row: int = 171
    enable_multiplier_detection: bool = True
    ds_loading_percent: float = 20.0
    ref_name_col: str = "C"
    ref_size_col: str = "E"
    ref_ds_col: str = "F"
    ref_stock_col: str = "G"
    ref_start_row: int = 1
    ref_end_row: int = 300
    ignore_countries: Tuple[str, ...] = ("NZ",)


def norm(v) -> str:
    return "" if v is None else str(v).strip()


def quote_sheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def col_range(start_col: str, end_col: str) -> List[str]:
    try:
        s = column_index_from_string(start_col.upper())
        e = column_index_from_string(end_col.upper())
    except Exception:
        return []
    if e < s:
        s, e = e, s
    return [get_column_letter(i) for i in range(s, e + 1)]


def load_rate_memory() -> Dict[str, float]:
    if "rate_memory" not in st.session_state:
        st.session_state.rate_memory = {}
        try:
            if RATE_MEMORY_PATH.exists():
                raw = json.loads(RATE_MEMORY_PATH.read_text(encoding="utf-8"))
                st.session_state.rate_memory = {str(k): float(v) for k, v in raw.items()}
        except Exception:
            st.session_state.rate_memory = {}
    return dict(st.session_state.rate_memory)


def save_rate_memory(memory: Dict[str, float]) -> None:
    cleaned = {str(k): float(v) for k, v in memory.items() if str(k).strip()}
    st.session_state.rate_memory = cleaned
    try:
        RATE_MEMORY_PATH.parent.mkdir(parents=True, exist_ok=True)
        RATE_MEMORY_PATH.write_text(json.dumps(cleaned, indent=2, sort_keys=True), encoding="utf-8")
    except Exception:
        pass


def detect_country_col(ws) -> str:
    terms = {"AUS", "AU", "AUSTRALIA", "NZ", "NEW ZEALAND", "FIJI", "SG", "SINGAPORE"}
    best_col, best_count = "I", -1
    for c in range(1, ws.max_column + 1):
        count = 0
        for r in range(1, min(ws.max_row, 250) + 1):
            if norm(ws.cell(r, c).value).upper() in terms:
                count += 1
        if count > best_count:
            best_count = count
            best_col = get_column_letter(c)
    return best_col


def detect_item_cols(ws, name_row=4, size_row=5, qty_row=7) -> Tuple[str, str]:
    size_pat = re.compile(r"\d+\s*(mm|cm)?\s*[x×]\s*\d+", re.I)
    hits = []
    for c in range(1, ws.max_column + 1):
        score = 0
        if norm(ws.cell(name_row, c).value):
            score += 1
        if size_pat.search(norm(ws.cell(size_row, c).value)):
            score += 3
        q = ws.cell(qty_row, c).value
        if isinstance(q, (int, float)) and q > 0:
            score += 2
        if score >= 3:
            hits.append(c)
    if not hits:
        return "AC", get_column_letter(ws.max_column)
    return get_column_letter(min(hits)), get_column_letter(max(hits))


def detect_reference_columns(ws) -> Dict[str, str]:
    out = {"name": "C", "size": "E", "ds": "F", "stock": "G"}
    for r in range(1, min(ws.max_row, 30) + 1):
        for c in range(1, ws.max_column + 1):
            t = norm(ws.cell(r, c).value).upper()
            col = get_column_letter(c)
            if t in {"ARTWORK", "ARTWORK NAME", "NAME", "DESCRIPTION", "ITEM"}:
                out["name"] = col
            if "FINISH SIZE" in t or t == "SIZE":
                out["size"] = col
            if "DS" in t and "SS" in t:
                out["ds"] = col
            if "MATERIAL" in t or "STOCK" in t:
                out["stock"] = col
    return out


def detect_reference_rows(ws, name_col="C", size_col="E") -> Tuple[int, int]:
    rows = []
    try:
        nc = column_index_from_string(name_col)
        sc = column_index_from_string(size_col)
    except Exception:
        return 1, ws.max_row
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, nc).value) and norm(ws.cell(r, sc).value):
            rows.append(r)
    return (min(rows), max(rows)) if rows else (1, ws.max_row)


def build_ignore_array(countries: List[str]) -> str:
    vals = [c.strip().upper().replace('"', '') for c in countries if c.strip()]
    if not vals:
        vals = ["NZ"]
    return f'"{vals[0]}"' if len(vals) == 1 else "{" + ",".join([f'"{v}"' for v in vals]) + "}"


def clean_qty_formula(col: str, cfg: MappingConfig, multiplier: int = 1) -> str:
    countries = [c for c in cfg.ignore_countries if c.strip()]
    # Uses only the original total qty row, then subtracts matching ignored-country rows from same item column.
    # Blank country rows do not match and do not subtract.
    if len(countries) <= 1:
        country = countries[0].upper() if countries else "NZ"
        base = f'({col}${cfg.total_qty_row}-SUMIF(${cfg.country_col}:${cfg.country_col},"{country}",{col}:{col}))'
    else:
        base = f'({col}${cfg.total_qty_row}-SUM(SUMIF(${cfg.country_col}:${cfg.country_col},{build_ignore_array(countries)},{col}:{col})))'
    if int(multiplier or 1) > 1:
        return f'={base}*{int(multiplier)}'
    return f'={base}'


def detect_name_multiplier(name_text: str) -> Dict[str, object]:
    """Return controlled multiplier decision from item/name text.

    Red/confident = multiply. Orange/suspicious = flag but do not multiply.
    The patterns are intentionally conservative to avoid silent quantity corruption.
    """
    text = norm(name_text)
    upper = text.upper()
    candidates: List[Tuple[int, str]] = []

    patterns = [
        (r'\bSET\s+OF\s+(\d{1,4})\b', 'set of N'),
        (r'\bSET\s*[x×]\s*(\d{1,4})\b', 'set x N'),
        (r'\bPACK\s+OF\s+(\d{1,5})\b', 'pack of N'),
        (r'\b(?:1\s*)?PACK\s*=\s*(\d{1,5})\b', 'pack = N'),
        (r'\b(?:1\s*)?PK\s*=\s*(\d{1,5})\b', 'pk = N'),
    ]
    for pat, label in patterns:
        for m in re.finditer(pat, upper):
            try:
                val = int(m.group(1))
                if 1 < val <= 10000:
                    candidates.append((val, label))
            except Exception:
                pass

    unique_vals = sorted(set(v for v, _ in candidates))
    marker_present = bool(re.search(r'\b(SET|PACK|PACKS|PK)\b', upper))

    if len(unique_vals) == 1:
        value = unique_vals[0]
        reason = ', '.join(sorted(set(label for v, label in candidates if v == value)))
        return {"status": "CONFIDENT", "multiplier": value, "reason": reason}

    if len(unique_vals) > 1:
        return {"status": "SUSPICIOUS", "multiplier": 1, "reason": f"multiple possible multipliers: {unique_vals}"}

    if marker_present:
        return {"status": "SUSPICIOUS", "multiplier": 1, "reason": "set/pack wording but no safe multiplier pattern"}

    return {"status": "NONE", "multiplier": 1, "reason": ""}


def ds_formula(col: str, cfg: MappingConfig) -> str:
    rs = quote_sheet(cfg.reference_sheet)
    return (
        f'=IFERROR(INDEX({rs}!${cfg.ref_ds_col}${cfg.ref_start_row}:${cfg.ref_ds_col}${cfg.ref_end_row},'
        f'MATCH(1,INDEX(({rs}!${cfg.ref_name_col}${cfg.ref_start_row}:${cfg.ref_name_col}${cfg.ref_end_row}={col}${cfg.name_row})*'
        f'({rs}!${cfg.ref_size_col}${cfg.ref_start_row}:${cfg.ref_size_col}${cfg.ref_end_row}={col}${cfg.size_row}),0),0)),"")'
    )


def sqm_formula(col: str, cfg: MappingConfig) -> str:
    size_cell = f'{col}${cfg.size_row}'
    qty_cell = f'{col}${cfg.clean_qty_output_row}'
    clean = f'LOWER(SUBSTITUTE(SUBSTITUTE({size_cell}," ",""),"mm",""))'
    xclean = f'SUBSTITUTE({clean},"×","x")'
    width = f'VALUE(LEFT({xclean},FIND("x",{xclean})-1))'
    height = f'VALUE(MID({xclean},FIND("x",{xclean})+1,99))'
    return f'=IFERROR(({width}*{height}/1000000)*{qty_cell},0)'


def price_formula(col: str, cfg: MappingConfig, rate_cell: str) -> str:
    ds_cell = f'{col}${cfg.ds_ss_output_row}'
    load = 1 + float(cfg.ds_loading_percent) / 100
    return f'=IFERROR({col}${cfg.sqm_output_row}*{rate_cell}*IF(OR(UPPER({ds_cell})="DS",UPPER({ds_cell})="DOUBLE SIDED",UPPER({ds_cell})="D/S"),{load},1),0)'


def unique_stocks(ws_values, cfg: MappingConfig) -> List[str]:
    vals = []
    for col in col_range(cfg.item_start_col, cfg.item_end_col):
        v = norm(ws_values[f"{col}{cfg.stock_row}"].value)
        if v and v not in vals:
            vals.append(v)
    return sorted(vals)


def preview_df(ws, max_rows=30, max_cols=30) -> pd.DataFrame:
    rows = []
    max_r = min(ws.max_row, max_rows)
    max_c = min(ws.max_column, max_cols)
    for r in range(1, max_r + 1):
        rows.append([ws.cell(r, c).value for c in range(1, max_c + 1)])
    df = pd.DataFrame(rows, columns=[get_column_letter(c) for c in range(1, max_c + 1)])
    df.insert(0, "Row", range(1, len(df) + 1))
    return df


def copy_row_style(ws, source_row: int, target_row: int, cols: List[str]) -> None:
    for col in cols:
        src = ws[f"{col}{source_row}"]
        dst = ws[f"{col}{target_row}"]
        if src.has_style:
            dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.font = copy(src.font)


def make_summary_sheet(wb, cfg: MappingConfig, selected_stocks: List[str], stock_rates: Dict[str, float]) -> None:
    if "Stock SQM Summary" in wb.sheetnames:
        del wb["Stock SQM Summary"]
    ws = wb.create_sheet("Stock SQM Summary")
    headers = ["Stock / Material", "Total SQM", "Base Rate / SQM", "DS Loading %", "Total Price"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill = PatternFill("solid", fgColor="0F172A")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    stock_range = f"'{cfg.working_sheet}'!${cfg.item_start_col}${cfg.stock_row}:${cfg.item_end_col}${cfg.stock_row}"
    sqm_range = f"'{cfg.working_sheet}'!${cfg.item_start_col}${cfg.sqm_output_row}:${cfg.item_end_col}${cfg.sqm_output_row}"
    price_range = f"'{cfg.working_sheet}'!${cfg.item_start_col}${cfg.price_output_row}:${cfg.item_end_col}${cfg.price_output_row}"
    for row, stock in enumerate(selected_stocks, 2):
        ws.cell(row, 1, stock)
        ws.cell(row, 2, f'=SUMIF({stock_range},A{row},{sqm_range})')
        ws.cell(row, 3, float(stock_rates.get(stock, 0)))
        ws.cell(row, 4, float(cfg.ds_loading_percent) / 100)
        ws.cell(row, 5, f'=SUMIF({stock_range},A{row},{price_range})')
        ws.cell(row, 2).number_format = '0.00'
        ws.cell(row, 3).number_format = '$#,##0.00'
        ws.cell(row, 4).number_format = '0%'
        ws.cell(row, 5).number_format = '$#,##0.00'
    for i, w in enumerate([42, 16, 18, 14, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def make_multiplier_audit_sheet(wb, audit_rows: List[Dict[str, object]]) -> None:
    if "Qty Multiplier Audit" in wb.sheetnames:
        del wb["Qty Multiplier Audit"]
    ws = wb.create_sheet("Qty Multiplier Audit")
    headers = ["Column", "Name", "Status", "Multiplier", "Reason"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill = PatternFill("solid", fgColor="0F172A")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    red = PatternFill("solid", fgColor="FCA5A5")
    orange = PatternFill("solid", fgColor="FDBA74")
    for r, item in enumerate(audit_rows, 2):
        ws.cell(r, 1, item.get("column", ""))
        ws.cell(r, 2, item.get("name", ""))
        ws.cell(r, 3, item.get("status", ""))
        ws.cell(r, 4, item.get("multiplier", 1))
        ws.cell(r, 5, item.get("reason", ""))
        status = item.get("status", "")
        fill = red if status == "CONFIDENT" else orange if status == "SUSPICIOUS" else None
        if fill:
            for c in range(1, 6):
                ws.cell(r, c).fill = fill
    for i, w in enumerate([12, 80, 16, 14, 45], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def apply_formulas(uploaded_bytes: bytes, cfg: MappingConfig, selected_stocks: List[str], stock_rates: Dict[str, float], stock_by_col=None, item_name_by_col=None) -> bytes:
    # Load the editable workbook only at export time. Avoid loading a second workbook here unless absolutely needed.
    wb = load_workbook(io.BytesIO(uploaded_bytes))
    ws = wb[cfg.working_sheet]
    cols = col_range(cfg.item_start_col, cfg.item_end_col)
    if stock_by_col is None or item_name_by_col is None:
        wb_values = load_workbook(io.BytesIO(uploaded_bytes), data_only=True, read_only=True)
        ws_values = wb_values[cfg.working_sheet]
        stock_by_col = {col: norm(ws_values[f"{col}{cfg.stock_row}"].value) for col in cols}
        item_name_by_col = {col: norm(ws_values[f"{col}{cfg.name_row}"].value) for col in cols}
    selected_set = set(selected_stocks)
    multiplier_by_col: Dict[str, Dict[str, object]] = {}
    audit_rows: List[Dict[str, object]] = []
    for col in cols:
        item_name = norm(item_name_by_col.get(col, ""))
        decision = detect_name_multiplier(item_name) if cfg.enable_multiplier_detection else {"status": "NONE", "multiplier": 1, "reason": "disabled"}
        multiplier_by_col[col] = decision
        if decision.get("status") in {"CONFIDENT", "SUSPICIOUS"}:
            audit_rows.append({"column": col, "name": item_name, **decision})

    for target in [cfg.ds_ss_output_row, cfg.clean_qty_output_row, cfg.sqm_output_row, cfg.price_output_row]:
        copy_row_style(ws, cfg.total_qty_row, target, cols)

    label_col = get_column_letter(max(1, column_index_from_string(cfg.item_start_col) - 1))
    labels = {
        cfg.ds_ss_output_row: "DS/SS Lookup",
        cfg.clean_qty_output_row: "Clean Qty",
        cfg.sqm_output_row: "SQM",
        cfg.price_output_row: "Price",
    }
    for row, label in labels.items():
        c = ws[f"{label_col}{row}"]
        c.value = label
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="F97316")
        c.alignment = Alignment(horizontal="right")

    for col in cols:
        decision = multiplier_by_col.get(col, {"status": "NONE", "multiplier": 1})
        multiplier = int(decision.get("multiplier", 1) or 1) if decision.get("status") == "CONFIDENT" else 1
        ws[f"{col}{cfg.ds_ss_output_row}"] = ds_formula(col, cfg)
        ws[f"{col}{cfg.clean_qty_output_row}"] = clean_qty_formula(col, cfg, multiplier)
        ws[f"{col}{cfg.sqm_output_row}"] = sqm_formula(col, cfg)
        if decision.get("status") == "CONFIDENT":
            fill = PatternFill("solid", fgColor="FCA5A5")  # red: multiplied automatically
            for row in [cfg.name_row, cfg.clean_qty_output_row, cfg.sqm_output_row, cfg.price_output_row]:
                ws[f"{col}{row}"].fill = fill
        elif decision.get("status") == "SUSPICIOUS":
            fill = PatternFill("solid", fgColor="FDBA74")  # orange: check manually, no multiply
            for row in [cfg.name_row, cfg.clean_qty_output_row, cfg.sqm_output_row, cfg.price_output_row]:
                ws[f"{col}{row}"].fill = fill
        stock = stock_by_col.get(col, "")
        if stock in selected_set:
            rate_row = selected_stocks.index(stock) + 2
            ws[f"{col}{cfg.price_output_row}"] = price_formula(col, cfg, f"'Stock SQM Summary'!$C${rate_row}")
        else:
            ws[f"{col}{cfg.price_output_row}"] = "=0"
        ws[f"{col}{cfg.sqm_output_row}"].number_format = '0.00'
        ws[f"{col}{cfg.price_output_row}"].number_format = '$#,##0.00'

    make_summary_sheet(wb, cfg, selected_stocks, stock_rates)
    make_multiplier_audit_sheet(wb, audit_rows)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def config_from_json(data: dict, sheet_names: List[str]):
    payload = dict(data)
    if isinstance(payload.get("ignore_countries"), list):
        payload["ignore_countries"] = tuple(payload["ignore_countries"])
    cfg = MappingConfig(**payload)
    if cfg.working_sheet not in sheet_names or cfg.reference_sheet not in sheet_names:
        return None
    return cfg


st.title("Excel Formula Fusion — V1.9")
st.caption("Adds controlled stock-rate updates to reduce Streamlit reruns and memory spikes.")

uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx"], key="main_workbook")
if not uploaded_file:
    st.info("Upload your workbook to start.")
    st.stop()

uploaded_bytes = uploaded_file.getvalue()
try:
    wb_formula = load_workbook(io.BytesIO(uploaded_bytes), read_only=False)
    wb_values = load_workbook(io.BytesIO(uploaded_bytes), data_only=True, read_only=False)
except Exception as e:
    st.error(f"Could not open workbook: {e}")
    st.stop()

sheet_names = wb_formula.sheetnames

if "cfg" not in st.session_state or st.session_state.get("last_file_name") != uploaded_file.name:
    default_work = "DL ANZ ALLOCATION" if "DL ANZ ALLOCATION" in sheet_names else sheet_names[0]
    default_ref = "PRINT DB" if "PRINT DB" in sheet_names else sheet_names[0]
    ws_work = wb_values[default_work]
    ws_ref = wb_values[default_ref]
    start_col, end_col = detect_item_cols(ws_work)
    ref_cols = detect_reference_columns(ws_ref)
    ref_start, ref_end = detect_reference_rows(ws_ref, ref_cols["name"], ref_cols["size"])
    st.session_state.cfg = MappingConfig(
        working_sheet=default_work,
        reference_sheet=default_ref,
        item_start_col=start_col,
        item_end_col=end_col,
        country_col=detect_country_col(ws_work),
        ref_name_col=ref_cols["name"],
        ref_size_col=ref_cols["size"],
        ref_ds_col=ref_cols["ds"],
        ref_stock_col=ref_cols["stock"],
        ref_start_row=ref_start,
        ref_end_row=ref_end,
    )
    st.session_state.last_file_name = uploaded_file.name

cfg: MappingConfig = st.session_state.cfg

with st.sidebar:
    st.header("Mapping / Memory")
    st.caption("Edit values in the main form, then press Apply Mapping.")
    mapping_json = st.file_uploader("Upload Mapping JSON", type=["json"], key="mapping_json_uploader")
    if mapping_json is not None:
        try:
            new_cfg = config_from_json(json.load(mapping_json), sheet_names)
            if new_cfg:
                st.session_state.cfg = new_cfg
                st.success("Mapping JSON loaded.")
                st.rerun()
            else:
                st.error("JSON sheet names do not match this workbook.")
        except Exception as e:
            st.error(f"Could not load mapping JSON: {e}")

with st.form("mapping_form", clear_on_submit=False):
    st.subheader("1) Mapping Setup")
    a, b, c, d = st.columns(4)
    with a:
        working_sheet = st.selectbox("Working Sheet", sheet_names, index=sheet_names.index(cfg.working_sheet))
        reference_sheet = st.selectbox("Reference Sheet", sheet_names, index=sheet_names.index(cfg.reference_sheet))
    with b:
        name_row = st.number_input("Name Row", min_value=1, value=int(cfg.name_row), step=1)
        size_row = st.number_input("Size Row", min_value=1, value=int(cfg.size_row), step=1)
        stock_row = st.number_input("Stock / Material Row", min_value=1, value=int(cfg.stock_row), step=1)
        total_qty_row = st.number_input("Original Total Qty Row", min_value=1, value=int(cfg.total_qty_row), step=1)
    with c:
        item_start_col = st.text_input("Item Start Column", value=cfg.item_start_col).upper().strip()
        item_end_col = st.text_input("Item End Column", value=cfg.item_end_col).upper().strip()
        country_col = st.text_input("Country Column", value=cfg.country_col).upper().strip()
        ignore_countries_str = st.text_input("Ignore Countries", value=", ".join(cfg.ignore_countries))
    with d:
        ds_ss_output_row = st.number_input("DS/SS Output Row", min_value=1, value=int(cfg.ds_ss_output_row), step=1)
        clean_qty_output_row = st.number_input("Clean Qty Output Row", min_value=1, value=int(cfg.clean_qty_output_row), step=1)
        sqm_output_row = st.number_input("SQM Output Row", min_value=1, value=int(cfg.sqm_output_row), step=1)
        price_output_row = st.number_input("Price Output Row", min_value=1, value=int(cfg.price_output_row), step=1)
        ds_loading_percent = st.number_input("DS Loading %", min_value=0.0, max_value=500.0, value=float(cfg.ds_loading_percent), step=1.0)
        enable_multiplier_detection = st.checkbox("Detect set/pack quantity multipliers", value=bool(cfg.enable_multiplier_detection))

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
    try:
        # Validate columns before storing.
        for col in [item_start_col, item_end_col, country_col, ref_name_col, ref_size_col, ref_ds_col, ref_stock_col]:
            column_index_from_string(col)
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
            ds_ss_output_row=int(ds_ss_output_row),
            clean_qty_output_row=int(clean_qty_output_row),
            sqm_output_row=int(sqm_output_row),
            price_output_row=int(price_output_row),
            enable_multiplier_detection=bool(enable_multiplier_detection),
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
    except Exception as e:
        st.error(f"Mapping not applied: {e}")

cfg = st.session_state.cfg
ws_values = wb_values[cfg.working_sheet]

st.subheader("Current Mapping Summary")
st.dataframe(pd.DataFrame([asdict(cfg)]).T.reset_index().rename(columns={"index": "Setting", 0: "Value"}), use_container_width=True, hide_index=True)

st.subheader("Quantity Multiplier Detection Preview")
if cfg.enable_multiplier_detection:
    audit_preview = []
    for col in col_range(cfg.item_start_col, cfg.item_end_col):
        item_name = norm(ws_values[f"{col}{cfg.name_row}"].value)
        decision = detect_name_multiplier(item_name)
        if decision.get("status") in {"CONFIDENT", "SUSPICIOUS"}:
            audit_preview.append({"Column": col, "Name": item_name, "Status": decision.get("status"), "Multiplier": decision.get("multiplier"), "Reason": decision.get("reason")})
    if audit_preview:
        st.dataframe(pd.DataFrame(audit_preview), use_container_width=True, hide_index=True)
        st.caption("CONFIDENT rows will be multiplied and highlighted red. SUSPICIOUS rows will be highlighted orange but not multiplied.")
    else:
        st.info("No set/pack multiplier wording detected in the selected item columns.")
else:
    st.info("Multiplier detection is disabled.")

st.subheader("Stock / Material Rates")
st.info("Rates are staged inside a form. Changing a price will not regenerate the workbook. Press Refresh / Update Rates when finished.")
all_stocks = unique_stocks(ws_values, cfg)
rate_memory = load_rate_memory()
selected_defaults = [s for s in all_stocks if s in rate_memory]
if not selected_defaults and all_stocks:
    selected_defaults = all_stocks[:1]
selected_stocks = st.multiselect("Pick one or more stocks/materials", all_stocks, default=selected_defaults, key="selected_stocks")

stock_rates: Dict[str, float] = {stock: float(rate_memory.get(stock, 0.0)) for stock in selected_stocks}
if selected_stocks:
    with st.form("stock_rates_form", clear_on_submit=False):
        st.caption("Enter or edit all rates, then press Refresh / Update Rates once.")
        cols = st.columns(min(4, max(1, len(selected_stocks))))
        staged_rates: Dict[str, float] = {}
        for i, stock in enumerate(selected_stocks):
            with cols[i % len(cols)]:
                key = "rate_input_" + re.sub(r"[^A-Za-z0-9_]+", "_", stock)[:80]
                staged_rates[stock] = st.number_input(
                    f"Rate: {stock[:35]}",
                    min_value=0.0,
                    value=float(rate_memory.get(stock, 0.0)),
                    step=0.10,
                    key=key,
                )
        update_rates = st.form_submit_button("Refresh / Update Rates")
    if update_rates:
        merged = load_rate_memory()
        merged.update(staged_rates)
        save_rate_memory(merged)
        st.success("Rates updated and saved to memory.")
        rate_memory = load_rate_memory()
        stock_rates = {stock: float(rate_memory.get(stock, 0.0)) for stock in selected_stocks}

    st.dataframe(
        pd.DataFrame([{"Stock / Material": s, "Saved Rate": float(load_rate_memory().get(s, 0.0))} for s in selected_stocks]),
        use_container_width=True,
        hide_index=True,
    )
else:
    st.warning("No stocks selected. Price row will be 0 unless you select stock/materials.")

with st.expander("Rate memory backup / restore"):
    rate_upload = st.file_uploader("Upload Stock Rate Memory JSON", type=["json"], key="rate_json")
    if rate_upload is not None:
        try:
            uploaded_rates = {str(k): float(v) for k, v in json.load(rate_upload).items()}
            save_rate_memory(uploaded_rates)
            st.success("Rate memory restored.")
        except Exception as e:
            st.error(f"Could not restore rates: {e}")
    st.download_button("Download Stock Rate Memory JSON", json.dumps(load_rate_memory(), indent=2, sort_keys=True), "stock_rates_memory.json", "application/json")

st.subheader("Formula Preview")
preview_col = cfg.item_start_col
example_decision = detect_name_multiplier(norm(ws_values[f"{preview_col}{cfg.name_row}"].value)) if cfg.enable_multiplier_detection else {"multiplier": 1, "status": "NONE"}
example_mult = int(example_decision.get("multiplier", 1) or 1) if example_decision.get("status") == "CONFIDENT" else 1
st.code(clean_qty_formula(preview_col, cfg, example_mult), language="excel")
st.code(ds_formula(preview_col, cfg), language="excel")
st.code(sqm_formula(preview_col, cfg), language="excel")
if selected_stocks:
    st.code(price_formula(preview_col, cfg, "'Stock SQM Summary'!$C$2"), language="excel")

with st.expander("Workbook Preview"):
    st.dataframe(preview_df(ws_values), use_container_width=True, hide_index=True)

st.subheader("Export")
e1, e2 = st.columns(2)
with e1:
    generate = st.button("Generate Excel Workbook")
with e2:
    st.download_button("Download Mapping JSON", json.dumps(asdict(cfg), indent=2), "excel_formula_fusion_mapping.json", "application/json")

if generate:
    try:
        merged = load_rate_memory()
        merged.update(stock_rates)
        save_rate_memory(merged)
        with st.spinner("Generating formula workbook..."):
            cols_for_export = col_range(cfg.item_start_col, cfg.item_end_col)
            stock_by_col = {col: norm(ws_values[f"{col}{cfg.stock_row}"].value) for col in cols_for_export}
            item_name_by_col = {col: norm(ws_values[f"{col}{cfg.name_row}"].value) for col in cols_for_export}
            latest_rates = {stock: float(load_rate_memory().get(stock, 0.0)) for stock in selected_stocks}
            output = apply_formulas(uploaded_bytes, cfg, selected_stocks, latest_rates, stock_by_col, item_name_by_col)
        st.success("Workbook generated.")
        st.download_button("Download Excel Workbook", output, "excel_formula_fusion_output.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Export failed: {e}")
