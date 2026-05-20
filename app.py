import io
import json
import re
import traceback
import zipfile
import xml.etree.ElementTree as ET
from copy import copy

import streamlit as st

APP_VERSION = "V2.2 Ultra Stable"

st.set_page_config(page_title="Excel Formula Fusion", layout="wide")

st.markdown("""
<style>
.stApp { background: #ffffff; color: #111827; }
[data-testid="stSidebar"] { background: #f7f8fb; }
.stButton button, .stDownloadButton button {
    background-color: #f36f21 !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    border: 1px solid #c95512 !important;
}
.stButton button *, .stDownloadButton button * { color: #ffffff !important; }
input, textarea { color: #111827 !important; background-color: #ffffff !important; }
div[data-baseweb="select"] > div { color: #111827 !important; background-color: #ffffff !important; }
</style>
""", unsafe_allow_html=True)

# ---------------- lightweight helpers ----------------

def norm(v):
    return "" if v is None else str(v).strip()


def safe_col(v, default):
    v = (v or "").strip().upper()
    return v if re.fullmatch(r"[A-Z]{1,3}", v) else default


def quote_sheet(name):
    return str(name).replace("'", "''")


def sheet_default(sheets, wanted, fallback=0):
    for s in sheets:
        if s.strip().lower() == wanted.strip().lower():
            return s
    if not sheets:
        return ""
    return sheets[min(fallback, len(sheets) - 1)]


def get_sheet_names_fast(xlsx_bytes):
    """Read sheet names without openpyxl. Keeps upload path lightweight."""
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
        root = ET.fromstring(z.read("xl/workbook.xml"))
    return [el.attrib.get("name", "") for el in root.findall("main:sheets/main:sheet", ns) if el.attrib.get("name")]


def detect_multiplier(name):
    t = norm(name).upper()
    m = re.search(r"\bSET\s*(?:OF)?\s*(\d{1,3})\b", t)
    if m:
        n = int(m.group(1))
        if 1 < n <= 500:
            return n, "red", f"set of {n}"
    m = re.search(r"\b(?:1\s*)?PACK\s*=\s*(\d{1,5})\b", t)
    if m:
        n = int(m.group(1))
        if 1 < n <= 10000:
            return n, "red", f"pack = {n}"
    m = re.search(r"\bPACK\s+OF\s+(\d{1,5})\b", t)
    if m:
        n = int(m.group(1))
        if 1 < n <= 10000:
            return n, "red", f"pack of {n}"
    if "SET" in t or "PACK" in t:
        return 1, "orange", "set/pack wording unclear"
    return 1, "none", ""


def parse_size_to_sqm(size_text):
    s = norm(size_text).lower().replace("×", "x").replace(",", "")
    nums = re.findall(r"(\d+(?:\.\d+)?)", s)
    if len(nums) >= 2:
        w, h = float(nums[0]), float(nums[1])
        if "cm" in s and "mm" not in s:
            return (w / 100.0) * (h / 100.0)
        if re.search(r"\bm\b", s) and "mm" not in s and "cm" not in s:
            return w * h
        return (w / 1000.0) * (h / 1000.0)
    if len(nums) == 1 and ("dia" in s or "ø" in s or "round" in s):
        d = float(nums[0])
        if "cm" in s and "mm" not in s:
            dm = d / 100.0
        elif re.search(r"\bm\b", s) and "mm" not in s and "cm" not in s:
            dm = d
        else:
            dm = d / 1000.0
        return 3.141592653589793 * (dm / 2.0) ** 2
    return 0.0


def ignored_array(countries):
    vals = [c.strip().upper().replace('"', "") for c in countries if c.strip()]
    if not vals:
        vals = ["NZ"]
    return "{" + ",".join(f'"{v}"' for v in vals) + "}"


def build_clean_qty_formula(col, total_row, scan_start, scan_end, country_col, ignored, multiplier):
    arr = ignored_array(ignored)
    base = (
        f"({col}${total_row}-SUMPRODUCT(({col}${scan_start}:{col}${scan_end})*"
        f"(--ISNUMBER(MATCH(UPPER(${country_col}${scan_start}:${country_col}${scan_end}),{arr},0)))))"
    )
    return f"={base}*{multiplier}" if multiplier != 1 else f"={base}"


def build_ds_formula(col, name_row, size_row, ref_sheet, ref_name_col, ref_size_col, ref_ds_col, ref_start, ref_end):
    rs = quote_sheet(ref_sheet)
    return (
        f"=IFERROR(INDEX('{rs}'!${ref_ds_col}${ref_start}:${ref_ds_col}${ref_end},"
        f"MATCH(1,INDEX(('{rs}'!${ref_name_col}${ref_start}:${ref_name_col}${ref_end}={col}${name_row})*"
        f"('{rs}'!${ref_size_col}${ref_start}:${ref_size_col}${ref_end}={col}${size_row}),0),0)),\"\")"
    )


def load_openpyxl():
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import column_index_from_string, get_column_letter
    return load_workbook, Font, PatternFill, column_index_from_string, get_column_letter


def get_stock_options(upload_bytes, working_sheet, stock_row, start_col, end_col):
    load_workbook, _, _, column_index_from_string, _ = load_openpyxl()
    wb = load_workbook(io.BytesIO(upload_bytes), read_only=True, data_only=True)
    try:
        ws = wb[working_sheet]
        start_i = column_index_from_string(start_col)
        end_i = column_index_from_string(end_col)
        out, seen = [], set()
        for c in range(start_i, end_i + 1):
            s = norm(ws.cell(stock_row, c).value)
            if s and s.upper() not in seen:
                seen.add(s.upper())
                out.append(s)
        return out
    finally:
        wb.close()


def detect_country_end(upload_bytes, working_sheet, country_col, start_row, limit=400):
    load_workbook, _, _, column_index_from_string, _ = load_openpyxl()
    wb = load_workbook(io.BytesIO(upload_bytes), read_only=True, data_only=True)
    try:
        ws = wb[working_sheet]
        tokens = {"NZ", "AUS", "AU", "AUSTRALIA", "NEW ZEALAND", "FIJI", "SG", "SINGAPORE"}
        cidx = column_index_from_string(country_col)
        last = start_row
        max_r = min(ws.max_row, limit)
        for r in range(start_row, max_r + 1):
            if norm(ws.cell(r, cidx).value).upper() in tokens:
                last = r
        return max(last, start_row)
    finally:
        wb.close()


def build_workbook(upload_bytes, cfg, rates):
    load_workbook, Font, PatternFill, column_index_from_string, get_column_letter = load_openpyxl()
    wb = load_workbook(io.BytesIO(upload_bytes))
    wbv = load_workbook(io.BytesIO(upload_bytes), data_only=True, read_only=True)
    try:
        ws = wb[cfg["working_sheet"]]
        wsv = wbv[cfg["working_sheet"]]

        start_col = column_index_from_string(cfg["item_start_col"])
        end_col = column_index_from_string(cfg["item_end_col"])
        selected = set(cfg.get("selected_stocks", []))
        ds_factor = 1.0 + (float(cfg.get("ds_loading", 20.0)) / 100.0)

        red_fill = PatternFill("solid", fgColor="FFC7CE")
        orange_fill = PatternFill("solid", fgColor="FCE4D6")
        green_fill = PatternFill("solid", fgColor="C6EFCE")
        yellow_fill = PatternFill("solid", fgColor="FFF2CC")
        blue_fill = PatternFill("solid", fgColor="1F4E78")
        white_bold = Font(color="FFFFFF", bold=True)

        def heading(cell):
            cell.fill = blue_fill
            cell.font = white_bold

        for row, label in [
            (cfg["ds_row"], "DS/SS Lookup"),
            (cfg["clean_qty_row"], "Clean Qty"),
            (cfg["multiplier_row"], "Qty Multiplier"),
            (cfg["sqm_row"], "SQM"),
            (cfg["price_row"], "Price"),
        ]:
            ws.cell(row, 1).value = label
            heading(ws.cell(row, 1))

        audit = []
        for c in range(start_col, end_col + 1):
            col = get_column_letter(c)
            name = norm(wsv.cell(cfg["name_row"], c).value or ws.cell(cfg["name_row"], c).value)
            size = norm(wsv.cell(cfg["size_row"], c).value or ws.cell(cfg["size_row"], c).value)
            stock = norm(wsv.cell(cfg["stock_row"], c).value or ws.cell(cfg["stock_row"], c).value)
            mult, flag, reason = detect_multiplier(name)

            ws.cell(cfg["ds_row"], c).value = build_ds_formula(
                col, cfg["name_row"], cfg["size_row"], cfg["reference_sheet"], cfg["ref_name_col"], cfg["ref_size_col"], cfg["ref_ds_col"], cfg["ref_start_row"], cfg["ref_end_row"]
            )
            ws.cell(cfg["clean_qty_row"], c).value = build_clean_qty_formula(
                col, cfg["total_qty_row"], cfg["country_scan_start"], cfg["country_scan_end"], cfg["country_col"], cfg["ignored_countries"], mult
            )
            ws.cell(cfg["multiplier_row"], c).value = mult
            sqm_each = parse_size_to_sqm(size)
            ws.cell(cfg["sqm_row"], c).value = f"={col}${cfg['clean_qty_row']}*{sqm_each:.6f}" if sqm_each else ""

            rate = float(rates.get(stock, 0.0))
            if stock in selected and rate > 0:
                ws.cell(cfg["price_row"], c).value = f'=IF(UPPER({col}${cfg["ds_row"]})="DS",{col}${cfg["sqm_row"]}*{rate}*{ds_factor},{col}${cfg["sqm_row"]}*{rate})'
            else:
                ws.cell(cfg["price_row"], c).value = ""

            if flag == "red":
                ws.cell(cfg["name_row"], c).fill = red_fill
                ws.cell(cfg["clean_qty_row"], c).fill = red_fill
                ws.cell(cfg["multiplier_row"], c).fill = red_fill
                audit.append([col, name, size, stock, mult, "RED", reason])
            elif flag == "orange":
                ws.cell(cfg["name_row"], c).fill = orange_fill
                ws.cell(cfg["multiplier_row"], c).fill = orange_fill
                audit.append([col, name, size, stock, mult, "ORANGE", reason])
            else:
                ws.cell(cfg["multiplier_row"], c).fill = green_fill
            if stock in selected:
                ws.cell(cfg["stock_row"], c).fill = yellow_fill

        if "Qty Multiplier Audit" in wb.sheetnames:
            del wb["Qty Multiplier Audit"]
        aud = wb.create_sheet("Qty Multiplier Audit")
        aud.append(["Column", "Name", "Size", "Stock", "Multiplier", "Flag", "Reason"])
        for cell in aud[1]:
            heading(cell)
        for row in audit:
            aud.append(row)
            fill = red_fill if row[5] == "RED" else orange_fill
            for cell in aud[aud.max_row]:
                cell.fill = fill

        if "Stock SQM Summary" in wb.sheetnames:
            del wb["Stock SQM Summary"]
        sm = wb.create_sheet("Stock SQM Summary")
        sm.append(["Stock", "Rate / sqm", "Total SQM", "Estimated Price", "DS Loading %"])
        for cell in sm[1]:
            heading(cell)
        work = quote_sheet(cfg["working_sheet"])
        r = 2
        for stock in cfg.get("selected_stocks", []):
            sm.cell(r, 1).value = stock
            sm.cell(r, 2).value = float(rates.get(stock, 0.0))
            sm.cell(r, 3).value = f'=SUMIF(\'{work}\'!${cfg["item_start_col"]}${cfg["stock_row"]}:${cfg["item_end_col"]}${cfg["stock_row"]},A{r},\'{work}\'!${cfg["item_start_col"]}${cfg["sqm_row"]}:${cfg["item_end_col"]}${cfg["sqm_row"]})'
            sm.cell(r, 4).value = f'=SUMIF(\'{work}\'!${cfg["item_start_col"]}${cfg["stock_row"]}:${cfg["item_end_col"]}${cfg["stock_row"]},A{r},\'{work}\'!${cfg["item_start_col"]}${cfg["price_row"]}:${cfg["item_end_col"]}${cfg["price_row"]})'
            sm.cell(r, 5).value = cfg.get("ds_loading", 20.0)
            r += 1

        for sheet in [aud, sm]:
            for col_idx in range(1, sheet.max_column + 1):
                sheet.column_dimensions[get_column_letter(col_idx)].width = 26

        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()
    finally:
        try:
            wbv.close()
        except Exception:
            pass
        try:
            wb.close()
        except Exception:
            pass

# ---------------- UI ----------------

st.title("Excel Formula Fusion")
st.caption(APP_VERSION)
st.info("This build is deliberately conservative: upload responds first, then you choose when to read sheets, refresh stock, update rates, and generate workbook.")

for k, v in {
    "uploaded_bytes": None,
    "uploaded_name": "",
    "sheet_names": [],
    "rate_memory": {},
    "cfg": None,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

with st.sidebar:
    st.header("1. Upload")
    uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx", "xlsm"])
    if uploaded is not None:
        st.session_state.uploaded_bytes = uploaded.getvalue()
        st.session_state.uploaded_name = uploaded.name
        st.session_state.sheet_names = []
        st.session_state.cfg = None
        st.success(f"Uploaded: {uploaded.name}")

    rates_file = st.file_uploader("Optional stock rate JSON", type=["json"])
    if rates_file is not None:
        try:
            st.session_state.rate_memory.update(json.loads(rates_file.getvalue().decode("utf-8")))
            st.success("Rate memory loaded")
        except Exception as exc:
            st.error(f"Could not read JSON: {exc}")

if not st.session_state.uploaded_bytes:
    st.stop()

st.success(f"Workbook selected: {st.session_state.uploaded_name}")
st.write(f"File size: {len(st.session_state.uploaded_bytes) / 1024 / 1024:.2f} MB")

if st.button("Read workbook / detect sheets", type="primary"):
    try:
        st.session_state.sheet_names = get_sheet_names_fast(st.session_state.uploaded_bytes)
        st.success("Sheet names detected.")
    except Exception:
        st.error("Could not read sheet names. This may not be a valid .xlsx/.xlsm file.")
        st.code(traceback.format_exc())

if not st.session_state.sheet_names:
    st.warning("Press **Read workbook / detect sheets** to continue.")
    st.stop()

sheets = st.session_state.sheet_names
st.write("Sheets detected:", ", ".join(sheets))
wd = sheet_default(sheets, "DL ANZ ALLOCATION", 0)
rd = sheet_default(sheets, "PRINT DB", 1 if len(sheets) > 1 else 0)

with st.form("mapping_form"):
    st.subheader("2. Mapping")
    c1, c2 = st.columns(2)
    with c1:
        working_sheet = st.selectbox("Working sheet", sheets, index=sheets.index(wd))
        name_row = int(st.number_input("Name row", min_value=1, value=4))
        size_row = int(st.number_input("Size row", min_value=1, value=5))
        stock_row = int(st.number_input("Stock/material row", min_value=1, value=6))
        total_qty_row = int(st.number_input("Original total qty row", min_value=1, value=7))
        country_col = safe_col(st.text_input("Country column", value="I"), "I")
        ignored = st.text_input("Ignore countries", value="NZ")
    with c2:
        reference_sheet = st.selectbox("Reference sheet", sheets, index=sheets.index(rd))
        ref_name_col = safe_col(st.text_input("Reference name column", value="C"), "C")
        ref_size_col = safe_col(st.text_input("Reference size column", value="E"), "E")
        ref_ds_col = safe_col(st.text_input("Reference DS/SS column", value="F"), "F")
        ref_start_row = int(st.number_input("Reference start row", min_value=1, value=12))
        ref_end_row = int(st.number_input("Reference end row", min_value=1, value=141))

    st.subheader("3. Columns and output rows")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        item_start_col = safe_col(st.text_input("Start column", value="AC"), "AC")
        item_end_col = safe_col(st.text_input("End column", value="IG"), "IG")
    with m2:
        ds_row = int(st.number_input("DS/SS output row", min_value=1, value=168))
        clean_qty_row = int(st.number_input("Clean qty row", min_value=1, value=169))
    with m3:
        multiplier_row = int(st.number_input("Multiplier row", min_value=1, value=170))
        sqm_row = int(st.number_input("SQM row", min_value=1, value=171))
    with m4:
        price_row = int(st.number_input("Price row", min_value=1, value=172))
        ds_loading = float(st.number_input("DS loading %", min_value=0.0, value=20.0, step=1.0))
    default_rate = float(st.number_input("Default new stock rate", min_value=0.0, value=0.0, step=0.1, format="%.2f"))
    apply_mapping = st.form_submit_button("Apply mapping / refresh stock list")

if apply_mapping:
    try:
        load_workbook, _, _, column_index_from_string, _ = load_openpyxl()
        if column_index_from_string(item_end_col) < column_index_from_string(item_start_col):
            raise ValueError("End column must be after Start column.")
        scan_start = total_qty_row + 1
        scan_end = detect_country_end(st.session_state.uploaded_bytes, working_sheet, country_col, scan_start)
        stocks = get_stock_options(st.session_state.uploaded_bytes, working_sheet, stock_row, item_start_col, item_end_col)
        st.session_state.cfg = {
            "working_sheet": working_sheet,
            "reference_sheet": reference_sheet,
            "name_row": name_row,
            "size_row": size_row,
            "stock_row": stock_row,
            "total_qty_row": total_qty_row,
            "country_col": country_col,
            "ignored_countries": [x.strip().upper() for x in ignored.split(",") if x.strip()],
            "ref_name_col": ref_name_col,
            "ref_size_col": ref_size_col,
            "ref_ds_col": ref_ds_col,
            "ref_start_row": ref_start_row,
            "ref_end_row": ref_end_row,
            "item_start_col": item_start_col,
            "item_end_col": item_end_col,
            "ds_row": ds_row,
            "clean_qty_row": clean_qty_row,
            "multiplier_row": multiplier_row,
            "sqm_row": sqm_row,
            "price_row": price_row,
            "ds_loading": ds_loading,
            "country_scan_start": scan_start,
            "country_scan_end": scan_end,
            "stock_options": stocks,
            "selected_stocks": [],
            "default_rate": default_rate,
        }
        st.success(f"Mapping applied. Stocks found from {item_start_col}:{item_end_col}: {len(stocks)}. Country rows scanned: {scan_start}:{scan_end}.")
    except Exception:
        st.error("Mapping failed. Details below.")
        st.code(traceback.format_exc())

if not st.session_state.cfg:
    st.warning("Apply mapping to continue.")
    st.stop()

cfg = st.session_state.cfg
with st.expander("Current mapping", expanded=False):
    st.json({k: v for k, v in cfg.items() if k != "stock_options"})

st.subheader("4. Stock/material rates")
st.caption(f"Stock names are read from row {cfg['stock_row']} between {cfg['item_start_col']} and {cfg['item_end_col']} only.")
selected_stocks = st.multiselect("Pick stock/material to calculate SQM and rate", cfg.get("stock_options", []), default=cfg.get("selected_stocks", []))
st.session_state.cfg["selected_stocks"] = selected_stocks

if selected_stocks:
    with st.form("rates_form"):
        staged = {}
        st.write("Enter rates below, then press **Refresh / Update Rates**. Typing alone does not generate the workbook.")
        for stock in selected_stocks:
            current = float(st.session_state.rate_memory.get(stock, cfg.get("default_rate", 0.0)))
            staged[stock] = st.number_input(f"Rate $/sqm — {stock}", min_value=0.0, value=current, step=0.1, format="%.2f", key="rate_" + re.sub(r"[^A-Za-z0-9]+", "_", stock)[:80])
        if st.form_submit_button("Refresh / Update Rates"):
            st.session_state.rate_memory.update(staged)
            st.success("Rates saved for this session.")
else:
    st.info("Select one or more stocks to show rate entry fields.")

st.download_button(
    "Download stock rate memory JSON",
    data=json.dumps(st.session_state.rate_memory, indent=2).encode("utf-8"),
    file_name="stock_rate_memory.json",
    mime="application/json",
)

st.subheader("5. Generate")
st.warning("Only press Generate after mapping and rates are correct. Narrow the Start/End columns if Streamlit Cloud memory is limited.")
if st.button("Generate Excel Workbook", type="primary"):
    try:
        with st.spinner("Generating workbook..."):
            output = build_workbook(st.session_state.uploaded_bytes, st.session_state.cfg, st.session_state.rate_memory)
        st.success("Workbook generated.")
        st.download_button(
            "Download Excel Workbook",
            data=output,
            file_name="formula_fusion_" + st.session_state.uploaded_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception:
        st.error("Generation failed. Details below.")
        st.code(traceback.format_exc())
