import io
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(page_title="Fusion Excel Reformat + Packaging Calculator", layout="wide")

ROLL_KEYWORDS_DEFAULT = [
    "roll", "banner", "vinyl", "sav", "fabric", "magnet", "magnetic", "ferrous",
    "backlit", "duratran", "silicone", "silicon edged", "pocket", "window"
]
FLAT_KEYWORDS_DEFAULT = [
    "foam", "pvc", "screenboard", "gsm", "paper", "card", "b-flute", "corflute",
    "forex", "palight", "pailight", "acm", "poly", "board"
]

@dataclass
class Detection:
    header_row: int
    size_row: int
    material_row: int
    side_row: int
    first_data_row: int
    country_col: Optional[int]
    zone_col: Optional[int]
    store_col: Optional[int]
    item_start_col: int
    item_end_col: int


def norm(v: Any) -> str:
    if v is None:
        return ""
    return str(v).replace("\n", " ").replace("\r", " ").strip()


def lower(v: Any) -> str:
    return norm(v).lower()


def is_qty(v: Any) -> bool:
    if v is None or v == "":
        return False
    if isinstance(v, (int, float)) and not pd.isna(v):
        return float(v) != 0
    s = norm(v).replace(",", "")
    if s.upper() in {"X", "Y", "YES"}:
        return True
    try:
        return float(s) != 0
    except Exception:
        return False


def qty_value(v: Any) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)) and not pd.isna(v):
        return float(v)
    s = norm(v).replace(",", "")
    if s.upper() in {"X", "Y", "YES"}:
        return 1.0
    try:
        return float(s)
    except Exception:
        return 0.0


def parse_size_mm(text: Any) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    """Return width, height, area_m2 from messy strings."""
    s = lower(text)
    if not s:
        return None, None, None
    s = s.replace("×", "x").replace("*", "x").replace("diameter", "dia").replace("ø", " dia ")
    # cm conversion if explicit cm and no mm nearby
    unit_factor = 10.0 if "cm" in s and "mm" not in s else 1.0

    # diameter / round
    dia = re.search(r"(?:dia|round|circle)[^0-9]{0,8}([0-9]+(?:\.[0-9]+)?)|([0-9]+(?:\.[0-9]+)?)\s*(?:mm|cm)?\s*(?:dia|round|circle)", s)
    if dia:
        val = dia.group(1) or dia.group(2)
        d = float(val) * unit_factor
        area = 3.141592653589793 * (d / 1000.0 / 2) ** 2
        return d, d, area

    # Prefer W/H labels when present
    w = re.search(r"w\s*([0-9][0-9,]*(?:\.[0-9]+)?)|([0-9][0-9,]*(?:\.[0-9]+)?)\s*(?:mm|cm)?\s*w", s)
    h = re.search(r"h\s*([0-9][0-9,]*(?:\.[0-9]+)?)|([0-9][0-9,]*(?:\.[0-9]+)?)\s*(?:mm|cm)?\s*h", s)
    if w and h:
        width = float((w.group(1) or w.group(2)).replace(",", "")) * unit_factor
        height = float((h.group(1) or h.group(2)).replace(",", "")) * unit_factor
        return width, height, (width * height) / 1_000_000

    # Generic A x B
    m = re.search(r"([0-9][0-9,]*(?:\.[0-9]+)?)\s*(?:mm|cm)?\s*x\s*([0-9][0-9,]*(?:\.[0-9]+)?)", s)
    if m:
        a = float(m.group(1).replace(",", "")) * unit_factor
        b = float(m.group(2).replace(",", "")) * unit_factor
        return a, b, (a * b) / 1_000_000

    return None, None, None


def find_row_by_keywords(ws, keywords: List[str], max_rows: int = 20) -> int:
    best_row, best_score = 1, -1
    for r in range(1, min(ws.max_row, max_rows) + 1):
        vals = [lower(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        row_text = " ".join(vals)
        score = sum(1 for k in keywords if k in row_text)
        # extra score for many dimension-looking cells
        score += sum(1 for v in vals if re.search(r"\d+\s*(mm|cm)?\s*[x×*]\s*\d+|\b[wh]\s*\d+", v))
        if score > best_score:
            best_score, best_row = score, r
    return best_row


def find_col(ws, row: int, keywords: List[str], max_cols: Optional[int] = None) -> Optional[int]:
    max_cols = max_cols or ws.max_column
    for c in range(1, max_cols + 1):
        v = lower(ws.cell(row, c).value)
        if any(k in v for k in keywords):
            return c
    # fallback scan first 15 rows
    for r in range(1, min(ws.max_row, 15) + 1):
        for c in range(1, max_cols + 1):
            v = lower(ws.cell(r, c).value)
            if any(k in v for k in keywords):
                return c
    return None


def detect_sheet(ws) -> Detection:
    header_row = find_row_by_keywords(ws, ["store", "zone", "country", "total"], 25)
    size_row = find_row_by_keywords(ws, ["spec", "size", "w", "h"], 15)
    material_row = find_row_by_keywords(ws, ["material", "stock", "gsm", "pvc", "vinyl", "magnet"], 15)
    side_row = find_row_by_keywords(ws, ["colour", "print", "ss", "ds", "double", "single"], 15)
    country_col = find_col(ws, header_row, ["country", "co"])
    zone_col = find_col(ws, header_row, ["zone"])
    store_col = find_col(ws, header_row, ["store name", "store"])

    item_start_col = 1
    # first item col is after location/admin fields; detect first col in size row with dimensions
    for c in range(1, ws.max_column + 1):
        if parse_size_mm(ws.cell(size_row, c).value)[2] is not None:
            item_start_col = c
            break
    item_end_col = ws.max_column
    for c in range(ws.max_column, item_start_col - 1, -1):
        if any(norm(ws.cell(r, c).value) for r in [size_row, material_row, side_row, header_row] if r <= ws.max_row):
            item_end_col = c
            break
    return Detection(header_row, size_row, material_row, side_row, header_row + 1, country_col, zone_col, store_col, item_start_col, item_end_col)


def classify_material(material: str, size_text: str, side_text: str, roll_keywords: List[str], flat_keywords: List[str]) -> str:
    txt = f"{material} {size_text} {side_text}".lower()
    if any(k.lower() in txt for k in roll_keywords):
        return "Roll"
    if any(k.lower() in txt for k in flat_keywords):
        return "Flat Pack"
    return "Flat Pack"


def analyse_workbook(file_bytes: bytes, selected_sheets: List[str], roll_keywords: List[str], flat_keywords: List[str], exclude_nz: bool) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Detection]]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
    item_rows: List[Dict[str, Any]] = []
    summary_rows: Dict[Tuple[str, str], Dict[str, Any]] = {}
    detections: Dict[str, Detection] = {}

    for sheet in selected_sheets:
        ws = wb[sheet]
        det = detect_sheet(ws)
        detections[sheet] = det
        for c in range(det.item_start_col, det.item_end_col + 1):
            size_text = norm(ws.cell(det.size_row, c).value)
            material = norm(ws.cell(det.material_row, c).value)
            side = norm(ws.cell(det.side_row, c).value)
            item = norm(ws.cell(det.header_row, c).value) or norm(ws.cell(max(1, det.header_row - 1), c).value)
            width, height, area = parse_size_mm(size_text)
            if not (size_text or material or item):
                continue
            if area is None and not material:
                continue
            pack_type = classify_material(material, size_text, side, roll_keywords, flat_keywords)
            total_qty = 0.0
            stores = 0
            zones_seen = set()
            for r in range(det.first_data_row, ws.max_row + 1):
                country = lower(ws.cell(r, det.country_col).value) if det.country_col else ""
                if exclude_nz and country in {"nz", "new zealand"}:
                    continue
                q = qty_value(ws.cell(r, c).value)
                if q:
                    total_qty += q
                    stores += 1
                    zone = norm(ws.cell(r, det.zone_col).value) if det.zone_col else "Unknown"
                    zones_seen.add(zone or "Unknown")
                    key = (sheet, zone or "Unknown")
                    row = summary_rows.setdefault(key, {"Sheet": sheet, "Zone": zone or "Unknown", "Stores": set(), "Flat Items": 0, "Roll Items": 0, "Total Qty": 0.0, "Max Roll Width mm": 0.0})
                    if det.store_col:
                        row["Stores"].add(norm(ws.cell(r, det.store_col).value))
                    if pack_type == "Roll":
                        row["Roll Items"] += 1
                        row["Max Roll Width mm"] = max(row["Max Roll Width mm"], width or height or 0)
                    else:
                        row["Flat Items"] += 1
                    row["Total Qty"] += q
            if total_qty:
                item_rows.append({
                    "Sheet": sheet,
                    "Column": get_column_letter(c),
                    "Item": item,
                    "Size/Specs": size_text,
                    "Width mm": width,
                    "Height mm": height,
                    "Area Each m²": area,
                    "Material/Stock": material,
                    "Side/Print": side,
                    "Pack Type": pack_type,
                    "Total Qty": total_qty,
                    "Store Count": stores,
                    "Zones": ", ".join(sorted(zones_seen)),
                })

    summary = []
    for row in summary_rows.values():
        stores_set = row.pop("Stores")
        row["No. Stores"] = len([s for s in stores_set if s])
        summary.append(row)
    return pd.DataFrame(item_rows), pd.DataFrame(summary), detections


def write_df_to_sheet(ws, df: pd.DataFrame):
    for c_idx, col in enumerate(df.columns, 1):
        cell = ws.cell(1, c_idx, col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r_idx, row in enumerate(df.itertuples(index=False), 2):
        for c_idx, value in enumerate(row, 1):
            ws.cell(r_idx, c_idx, value)
    max_row = max(2, len(df) + 1)
    max_col = max(1, len(df.columns))
    ref = f"A1:{get_column_letter(max_col)}{max_row}"
    tab = Table(displayName=re.sub(r"\W+", "_", ws.title)[:20] + "Table", ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws.add_table(tab)
    ws.freeze_panes = "A2"
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        max_len = max(len(str(ws.cell(r, col).value or "")) for r in range(1, min(max_row, 80) + 1))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 42)


def make_export(file_bytes: bytes, item_df: pd.DataFrame, summary_df: pd.DataFrame, detections: Dict[str, Detection], reformat: bool) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))
    for name in ["PACKAGING SUMMARY", "ITEM PACK TYPE", "DETECTION SETTINGS"]:
        if name in wb.sheetnames:
            del wb[name]
    ws1 = wb.create_sheet("PACKAGING SUMMARY", 0)
    ws2 = wb.create_sheet("ITEM PACK TYPE", 1)
    ws3 = wb.create_sheet("DETECTION SETTINGS", 2)
    write_df_to_sheet(ws1, summary_df if not summary_df.empty else pd.DataFrame(columns=["Sheet", "Zone", "Flat Items", "Roll Items", "Total Qty", "Max Roll Width mm", "No. Stores"]))
    write_df_to_sheet(ws2, item_df if not item_df.empty else pd.DataFrame(columns=["Sheet", "Column", "Item", "Size/Specs", "Material/Stock", "Pack Type", "Total Qty"]))
    det_df = pd.DataFrame([
        {"Sheet": s, "Header Row": d.header_row, "Size Row": d.size_row, "Material Row": d.material_row, "Side Row": d.side_row,
         "First Data Row": d.first_data_row, "Country Column": get_column_letter(d.country_col) if d.country_col else "", "Zone Column": get_column_letter(d.zone_col) if d.zone_col else "", "Item Columns": f"{get_column_letter(d.item_start_col)}:{get_column_letter(d.item_end_col)}"}
        for s, d in detections.items()
    ])
    write_df_to_sheet(ws3, det_df)

    if reformat:
        thin = Side(style="thin", color="D9E2F3")
        for ws in wb.worksheets:
            if ws.title in {"PACKAGING SUMMARY", "ITEM PACK TYPE", "DETECTION SETTINGS"}:
                continue
            ws.freeze_panes = "A10"
            for col in range(1, min(ws.max_column, 80) + 1):
                ws.column_dimensions[get_column_letter(col)].width = min(max(ws.column_dimensions[get_column_letter(col)].width or 12, 10), 22)
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 12), max_col=ws.max_column):
                for cell in row:
                    if cell.value not in (None, ""):
                        cell.alignment = Alignment(wrap_text=True, vertical="center")
                        cell.border = Border(bottom=thin)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


st.title("Fusion Excel Reformat + Flat/Roll Packaging Calculator")
st.caption("Upload a Fusion allocation workbook, review detection, calculate flat packs / rolls, then export the workbook with summary sheets added.")

with st.sidebar:
    st.header("Rules")
    exclude_nz = st.checkbox("Exclude NZ rows", value=True)
    reformat = st.checkbox("Apply light reformat to copied workbook", value=False)
    roll_text = st.text_area("Roll keywords", value=", ".join(ROLL_KEYWORDS_DEFAULT), height=110)
    flat_text = st.text_area("Flat-pack keywords", value=", ".join(FLAT_KEYWORDS_DEFAULT), height=110)
    st.warning("Do not trust automation blindly. Check the detected rows before using exported counts.")

uploaded = st.file_uploader("Upload Fusion Excel workbook", type=["xlsx"])

if uploaded:
    file_bytes = uploaded.getvalue()
    wb_preview = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=False)
    sheets = wb_preview.sheetnames
    default_sheets = [s for s in sheets if not any(x in s.lower() for x in ["summary", "pricing"])]
    selected_sheets = st.multiselect("Sheets to process", sheets, default=default_sheets[:6])

    roll_keywords = [x.strip() for x in roll_text.split(",") if x.strip()]
    flat_keywords = [x.strip() for x in flat_text.split(",") if x.strip()]

    if selected_sheets:
        item_df, summary_df, detections = analyse_workbook(file_bytes, selected_sheets, roll_keywords, flat_keywords, exclude_nz)

        st.subheader("Detected sheet settings")
        det_show = pd.DataFrame([
            {"Sheet": s, "Header Row": d.header_row, "Size Row": d.size_row, "Material Row": d.material_row, "Side Row": d.side_row,
             "First Data Row": d.first_data_row, "Country Column": get_column_letter(d.country_col) if d.country_col else "", "Zone Column": get_column_letter(d.zone_col) if d.zone_col else "", "Item Columns": f"{get_column_letter(d.item_start_col)}:{get_column_letter(d.item_end_col)}"}
            for s, d in detections.items()
        ])
        st.dataframe(det_show, use_container_width=True, hide_index=True)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Flat-pack items", int(summary_df["Flat Items"].sum()) if not summary_df.empty else 0)
        c2.metric("Roll items", int(summary_df["Roll Items"].sum()) if not summary_df.empty else 0)
        c3.metric("Total qty", int(summary_df["Total Qty"].sum()) if not summary_df.empty else 0)
        c4.metric("Max roll width mm", int(summary_df["Max Roll Width mm"].max()) if not summary_df.empty else 0)

        tab1, tab2, tab3 = st.tabs(["Packaging Summary", "Item Classification", "Excel Preview"])
        with tab1:
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(item_df, use_container_width=True, hide_index=True)
        with tab3:
            sheet_to_preview = st.selectbox("Preview sheet", selected_sheets)
            preview = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_to_preview, header=None, nrows=80)
            preview.columns = [get_column_letter(i + 1) for i in range(len(preview.columns))]
            preview.index = preview.index + 1
            st.dataframe(preview, use_container_width=True)

        out_bytes = make_export(file_bytes, item_df, summary_df, detections, reformat)
        st.download_button(
            "Download updated workbook",
            data=out_bytes,
            file_name="fusion_packaging_calculated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Upload an `.xlsx` workbook to start. The app will not overwrite your original file.")
