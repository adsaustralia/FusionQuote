import io
import json
import re
from copy import copy
from dataclasses import asdict, dataclass
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

APP_TITLE = "Excel Formula Fusion — Dynamic V1"

DEFAULTS = {
    "working_sheet": "DL ANZ ALLOCATION",
    "reference_sheet": "PRINT DB",
    "active_start_col": "AC",
    "active_end_col": "EU",
    "working_name_row": 4,
    "working_size_row": 5,
    "qty_start_row": 8,
    "qty_end_row": 166,
    "country_col": "I",
    "dsss_output_row": 168,
    "clean_qty_output_row": 169,
    "ref_start_row": 12,
    "ref_end_row": 141,
    "ref_name_col": "C",
    "ref_size_col": "E",
    "ref_dsss_col": "F",
    "ignore_countries": ["NZ"],
}

KNOWN_COUNTRIES = {"AUS", "AU", "AUSTRALIA", "NZ", "NEW ZEALAND", "FIJI", "SG", "SINGAPORE"}


@dataclass
class MappingConfig:
    working_sheet: str
    reference_sheet: str
    active_start_col: str
    active_end_col: str
    working_name_row: int
    working_size_row: int
    qty_start_row: int
    qty_end_row: int
    country_col: str
    dsss_output_row: int
    clean_qty_output_row: int
    ref_start_row: int
    ref_end_row: int
    ref_name_col: str
    ref_size_col: str
    ref_dsss_col: str
    ignore_countries: List[str]


def norm(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text


def safe_sheet_name(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def clean_col(value: str) -> str:
    value = str(value or "").strip().upper()
    value = re.sub(r"[^A-Z]", "", value)
    if not value:
        raise ValueError("Column cannot be blank")
    column_index_from_string(value)
    return value


def load_workbook_from_upload(uploaded_file):
    data = uploaded_file.getvalue()
    return load_workbook(io.BytesIO(data), data_only=False), data


def sheet_preview(ws, max_rows: int = 40, max_cols: int = 35) -> pd.DataFrame:
    max_r = min(ws.max_row, max_rows)
    max_c = min(ws.max_column, max_cols)
    headers = [get_column_letter(c) for c in range(1, max_c + 1)]
    rows = []
    for r in range(1, max_r + 1):
        rows.append([ws.cell(r, c).value for c in range(1, max_c + 1)])
    df = pd.DataFrame(rows, columns=headers)
    df.insert(0, "Row", range(1, max_r + 1))
    return df


def detect_country_column(ws, scan_rows: int = 220) -> str:
    best_col = DEFAULTS["country_col"]
    best_score = 0
    for c in range(1, min(ws.max_column, 40) + 1):
        score = 0
        for r in range(1, min(ws.max_row, scan_rows) + 1):
            if norm(ws.cell(r, c).value) in KNOWN_COUNTRIES:
                score += 1
        if score > best_score:
            best_score = score
            best_col = get_column_letter(c)
    return best_col


def detect_active_columns(ws, name_row: int, size_row: int) -> Tuple[str, str]:
    populated = []
    size_pattern = re.compile(r"(\d+\s*(X|×)\s*\d+|\d+\s?MM|\d+\s?CM|DIA|Ø)", re.I)
    for c in range(1, ws.max_column + 1):
        name = ws.cell(name_row, c).value
        size = ws.cell(size_row, c).value
        if name not in (None, "") or (size is not None and size_pattern.search(str(size))):
            populated.append(c)
    if not populated:
        return DEFAULTS["active_start_col"], DEFAULTS["active_end_col"]
    # Ignore early admin columns where possible. HOKA campaign item columns start at AC.
    populated = [c for c in populated if c >= column_index_from_string("J")] or populated
    return get_column_letter(min(populated)), get_column_letter(max(populated))


def detect_qty_end_row(ws, qty_start_row: int, country_col: str) -> int:
    c = column_index_from_string(country_col)
    last_country_row = qty_start_row
    for r in range(qty_start_row, min(ws.max_row, 500) + 1):
        if norm(ws.cell(r, c).value) in KNOWN_COUNTRIES:
            last_country_row = r
    return max(qty_start_row, last_country_row)


def detect_ref_end_row(ws_ref, start_row: int, name_col: str, size_col: str) -> int:
    nc = column_index_from_string(name_col)
    sc = column_index_from_string(size_col)
    last = start_row
    for r in range(start_row, ws_ref.max_row + 1):
        if ws_ref.cell(r, nc).value not in (None, "") or ws_ref.cell(r, sc).value not in (None, ""):
            last = r
    return last


def detect_countries(ws, country_col: str, start_row: int, end_row: int) -> List[str]:
    c = column_index_from_string(country_col)
    vals = []
    for r in range(start_row, end_row + 1):
        v = norm(ws.cell(r, c).value)
        if v:
            vals.append(v)
    return sorted(set(vals))


def build_clean_qty_formula(col_letter: str, cfg: MappingConfig) -> str:
    col_letter = clean_col(col_letter)
    country_col = clean_col(cfg.country_col)
    qty_range = f"{col_letter}${cfg.qty_start_row}:{col_letter}${cfg.qty_end_row}"
    country_range = f"${country_col}${cfg.qty_start_row}:${country_col}${cfg.qty_end_row}"
    conditions = "".join(
        f"*(--(UPPER({country_range})<>\"{norm(country)}\"))"
        for country in cfg.ignore_countries
        if norm(country)
    )
    if not conditions:
        return f"=SUM({qty_range})"
    return f"=SUMPRODUCT(({qty_range}){conditions})"


def build_dsss_formula(col_letter: str, cfg: MappingConfig) -> str:
    col_letter = clean_col(col_letter)
    ref_sheet = safe_sheet_name(cfg.reference_sheet)
    ref_name_col = clean_col(cfg.ref_name_col)
    ref_size_col = clean_col(cfg.ref_size_col)
    ref_dsss_col = clean_col(cfg.ref_dsss_col)
    name_rng = f"{ref_sheet}!${ref_name_col}${cfg.ref_start_row}:${ref_name_col}${cfg.ref_end_row}"
    size_rng = f"{ref_sheet}!${ref_size_col}${cfg.ref_start_row}:${ref_size_col}${cfg.ref_end_row}"
    dsss_rng = f"{ref_sheet}!${ref_dsss_col}${cfg.ref_start_row}:${ref_dsss_col}${cfg.ref_end_row}"
    name_cell = f"{col_letter}${cfg.working_name_row}"
    size_cell = f"{col_letter}${cfg.working_size_row}"
    return f"=IFERROR(INDEX({dsss_rng},MATCH(1,INDEX(({name_rng}={name_cell})*({size_rng}={size_cell}),0),0)),\"\")"


def copy_cell_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def apply_formulas(wb, cfg: MappingConfig) -> Tuple[int, int]:
    ws = wb[cfg.working_sheet]
    start = column_index_from_string(clean_col(cfg.active_start_col))
    end = column_index_from_string(clean_col(cfg.active_end_col))
    if end < start:
        raise ValueError("End column must be after start column")
    count = 0
    for col_idx in range(start, end + 1):
        col = get_column_letter(col_idx)
        dsss_cell = ws.cell(cfg.dsss_output_row, col_idx)
        copy_cell_style(ws.cell(cfg.working_name_row, col_idx), dsss_cell)
        dsss_cell.value = build_dsss_formula(col, cfg)
        count += 1

        qty_cell = ws.cell(cfg.clean_qty_output_row, col_idx)
        copy_cell_style(ws.cell(cfg.working_size_row, col_idx), qty_cell)
        qty_cell.value = build_clean_qty_formula(col, cfg)
        count += 1
    return count, end - start + 1


def build_reference_index(ws_ref, cfg: MappingConfig) -> Dict[Tuple[str, str], List[Tuple[int, str]]]:
    idx: Dict[Tuple[str, str], List[Tuple[int, str]]] = {}
    nc = column_index_from_string(clean_col(cfg.ref_name_col))
    sc = column_index_from_string(clean_col(cfg.ref_size_col))
    dc = column_index_from_string(clean_col(cfg.ref_dsss_col))
    for r in range(cfg.ref_start_row, cfg.ref_end_row + 1):
        key = (norm(ws_ref.cell(r, nc).value), norm(ws_ref.cell(r, sc).value))
        if not key[0] and not key[1]:
            continue
        idx.setdefault(key, []).append((r, ws_ref.cell(r, dc).value))
    return idx


def validate_matches(wb, cfg: MappingConfig) -> pd.DataFrame:
    ws = wb[cfg.working_sheet]
    ws_ref = wb[cfg.reference_sheet]
    idx = build_reference_index(ws_ref, cfg)
    start = column_index_from_string(clean_col(cfg.active_start_col))
    end = column_index_from_string(clean_col(cfg.active_end_col))
    rows = []
    for col_idx in range(start, end + 1):
        col = get_column_letter(col_idx)
        item_name = ws.cell(cfg.working_name_row, col_idx).value
        item_size = ws.cell(cfg.working_size_row, col_idx).value
        matches = idx.get((norm(item_name), norm(item_size)), [])
        if len(matches) == 1:
            status, ref_row, dsss = "Exact Match", matches[0][0], matches[0][1]
        elif len(matches) > 1:
            status = "Duplicate Match"
            ref_row = ", ".join(str(m[0]) for m in matches[:5])
            dsss = ", ".join(str(m[1]) for m in matches[:5])
        else:
            status, ref_row, dsss = "No Match", "", ""
        rows.append({
            "Column": col,
            "Working Name": item_name,
            "Working Size": item_size,
            "Status": status,
            "Reference Row": ref_row,
            "DS/SS": dsss,
        })
    return pd.DataFrame(rows)


def qty_range_validation(wb, cfg: MappingConfig) -> pd.DataFrame:
    ws = wb[cfg.working_sheet]
    start = column_index_from_string(clean_col(cfg.active_start_col))
    end = min(column_index_from_string(clean_col(cfg.active_end_col)), start + 14)
    cc = column_index_from_string(clean_col(cfg.country_col))
    rows = []
    for col_idx in range(start, end + 1):
        col = get_column_letter(col_idx)
        end_val = ws.cell(cfg.qty_end_row, col_idx).value
        end_country = ws.cell(cfg.qty_end_row, cc).value
        warning = ""
        if isinstance(end_val, str) and end_val.startswith("="):
            warning = "End row contains formula — likely total/subtotal. Reduce Qty End Row."
        elif end_val not in (None, "") and norm(end_country) not in KNOWN_COUNTRIES:
            warning = "End row has qty but country is blank/unknown. Confirm this is a store row."
        rows.append({
            "Column": col,
            "Formula": build_clean_qty_formula(col, cfg),
            "Qty End Row Value": end_val,
            "Qty End Row Country": end_country,
            "Warning": warning,
        })
    return pd.DataFrame(rows)


def save_workbook_bytes(wb) -> bytes:
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def cfg_json(cfg: MappingConfig) -> str:
    return json.dumps(asdict(cfg), indent=2)


def json_to_cfg(data: str, sheet_names: List[str]) -> dict:
    raw = json.loads(data)
    result = DEFAULTS.copy()
    result.update(raw)
    if result["working_sheet"] not in sheet_names:
        result["working_sheet"] = sheet_names[0]
    if result["reference_sheet"] not in sheet_names:
        result["reference_sheet"] = sheet_names[0]
    return result


def set_page_style():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(
        """
        <style>
        .stApp { background:#f7f8fb; }
        section[data-testid="stSidebar"] { background:#0f2742; }
        section[data-testid="stSidebar"] * { color:white; }
        .block-container { padding-top:1.4rem; }
        div[data-testid="stMetric"] { background:white; padding:0.8rem; border-radius:14px; border-left:5px solid #f58220; box-shadow:0 2px 12px rgba(0,0,0,0.06); }
        </style>
        """,
        unsafe_allow_html=True,
    )


def get_initial_values(wb):
    sheets = wb.sheetnames
    vals = DEFAULTS.copy()
    vals["working_sheet"] = DEFAULTS["working_sheet"] if DEFAULTS["working_sheet"] in sheets else sheets[0]
    vals["reference_sheet"] = DEFAULTS["reference_sheet"] if DEFAULTS["reference_sheet"] in sheets else sheets[0]
    try:
        ws = wb[vals["working_sheet"]]
        vals["country_col"] = detect_country_column(ws)
        vals["active_start_col"], vals["active_end_col"] = detect_active_columns(ws, vals["working_name_row"], vals["working_size_row"])
        vals["qty_end_row"] = detect_qty_end_row(ws, vals["qty_start_row"], vals["country_col"])
        vals["ignore_countries"] = ["NZ"]
    except Exception:
        pass
    try:
        vals["ref_end_row"] = detect_ref_end_row(wb[vals["reference_sheet"]], vals["ref_start_row"], vals["ref_name_col"], vals["ref_size_col"])
    except Exception:
        pass
    return vals


def sidebar_config(wb) -> MappingConfig:
    sheets = wb.sheetnames
    initial = get_initial_values(wb)

    with st.sidebar:
        st.header("Workbook")
        mapping_upload = st.file_uploader("Optional: load mapping JSON", type=["json"])
        if mapping_upload:
            try:
                initial = json_to_cfg(mapping_upload.getvalue().decode("utf-8"), sheets)
                st.success("Mapping loaded")
            except Exception as exc:
                st.error(f"Could not load mapping JSON: {exc}")

        working_sheet = st.selectbox("Working sheet", sheets, index=sheets.index(initial["working_sheet"]))
        reference_sheet = st.selectbox("Reference sheet", sheets, index=sheets.index(initial["reference_sheet"]))

        st.header("Working Sheet Mapping")
        c1, c2 = st.columns(2)
        active_start_col = c1.text_input("Start item column", initial["active_start_col"]).strip().upper()
        active_end_col = c2.text_input("End item column", initial["active_end_col"]).strip().upper()
        working_name_row = st.number_input("Name row", min_value=1, value=int(initial["working_name_row"]), step=1)
        working_size_row = st.number_input("Size row", min_value=1, value=int(initial["working_size_row"]), step=1)
        qty_start_row = st.number_input("Qty start row", min_value=1, value=int(initial["qty_start_row"]), step=1)
        qty_end_row = st.number_input("Qty end row", min_value=1, value=int(initial["qty_end_row"]), step=1)
        country_col = st.text_input("Country column", initial["country_col"]).strip().upper()
        dsss_output_row = st.number_input("DS/SS output row", min_value=1, value=int(initial["dsss_output_row"]), step=1)
        clean_qty_output_row = st.number_input("Clean qty output row", min_value=1, value=int(initial["clean_qty_output_row"]), step=1)

        st.header("Reference Sheet Mapping")
        ref_start_row = st.number_input("Reference start row", min_value=1, value=int(initial["ref_start_row"]), step=1)
        ref_end_row = st.number_input("Reference end row", min_value=1, value=int(initial["ref_end_row"]), step=1)
        ref_name_col = st.text_input("Reference name column", initial["ref_name_col"]).strip().upper()
        ref_size_col = st.text_input("Reference size column", initial["ref_size_col"]).strip().upper()
        ref_dsss_col = st.text_input("Reference DS/SS column", initial["ref_dsss_col"]).strip().upper()

        st.header("Country Exclusion")
        detected = []
        try:
            detected = detect_countries(wb[working_sheet], country_col, int(qty_start_row), int(qty_end_row))
        except Exception:
            pass
        options = sorted(set(detected + initial["ignore_countries"] + ["NZ"]))
        ignore_countries = st.multiselect("Ignore countries", options=options, default=[c for c in initial["ignore_countries"] if c in options])
        extra = st.text_input("Extra ignore countries, comma separated", "")
        ignore_countries += [x.strip().upper() for x in extra.split(",") if x.strip()]

    return MappingConfig(
        working_sheet=working_sheet,
        reference_sheet=reference_sheet,
        active_start_col=active_start_col,
        active_end_col=active_end_col,
        working_name_row=int(working_name_row),
        working_size_row=int(working_size_row),
        qty_start_row=int(qty_start_row),
        qty_end_row=int(qty_end_row),
        country_col=country_col,
        dsss_output_row=int(dsss_output_row),
        clean_qty_output_row=int(clean_qty_output_row),
        ref_start_row=int(ref_start_row),
        ref_end_row=int(ref_end_row),
        ref_name_col=ref_name_col,
        ref_size_col=ref_size_col,
        ref_dsss_col=ref_dsss_col,
        ignore_countries=sorted(set([norm(x) for x in ignore_countries if norm(x)])),
    )


def main():
    set_page_style()
    st.title(APP_TITLE)
    st.caption("Dynamic formula generator for clean quantity rows and DS/SS lookup. Upload workbook → confirm mappings → preview formulas → export formula workbook.")

    uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx"])
    if not uploaded:
        st.info("Upload your workbook to begin. Defaults are tuned for HOKA Clifton 11 MASTER DBDL.")
        st.stop()

    try:
        wb, _ = load_workbook_from_upload(uploaded)
    except Exception as exc:
        st.error(f"Could not open workbook: {exc}")
        st.stop()

    cfg = sidebar_config(wb)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Workbook", uploaded.name)
    m2.metric("Sheets", len(wb.sheetnames))
    m3.metric("Working Sheet", cfg.working_sheet)
    m4.metric("Reference Sheet", cfg.reference_sheet)

    with st.expander("1. Workbook Preview", expanded=True):
        preview_sheet = st.selectbox("Preview sheet", wb.sheetnames, index=wb.sheetnames.index(cfg.working_sheet))
        a, b = st.columns(2)
        rows_to_show = a.slider("Rows to preview", 10, min(120, wb[preview_sheet].max_row), min(45, wb[preview_sheet].max_row))
        cols_to_show = b.slider("Columns to preview", 8, min(90, wb[preview_sheet].max_column), min(35, wb[preview_sheet].max_column))
        st.dataframe(sheet_preview(wb[preview_sheet], rows_to_show, cols_to_show), use_container_width=True, hide_index=True)

    with st.expander("2. Formula Preview", expanded=True):
        sample_col = clean_col(cfg.active_start_col)
        st.write("Clean quantity formula")
        st.code(build_clean_qty_formula(sample_col, cfg), language="excel")
        st.write("DS/SS lookup formula")
        st.code(build_dsss_formula(sample_col, cfg), language="excel")

    with st.expander("3. Quantity Range Validation", expanded=True):
        st.warning("Check this table before export. If the end row is a total/subtotal row, reduce Qty End Row. For the HOKA file, 8:166 is usually correct; 167 may double-count totals.")
        try:
            st.dataframe(qty_range_validation(wb, cfg), use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(f"Quantity validation failed: {exc}")

    with st.expander("4. DS/SS Match Validation", expanded=True):
        try:
            val = validate_matches(wb, cfg)
            counts = val["Status"].value_counts().to_dict()
            c1, c2, c3 = st.columns(3)
            c1.metric("Exact", counts.get("Exact Match", 0))
            c2.metric("No Match", counts.get("No Match", 0))
            c3.metric("Duplicate", counts.get("Duplicate Match", 0))
            st.dataframe(val, use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(f"DS/SS validation failed: {exc}")

    with st.expander("5. Export", expanded=True):
        st.write("Export writes formulas into the working sheet. Excel calculates values when opened.")
        if st.button("Generate formula workbook", type="primary"):
            try:
                formula_count, item_count = apply_formulas(wb, cfg)
                out = save_workbook_bytes(wb)
                st.success(f"Generated {formula_count} formulas across {item_count} item columns.")
                st.download_button(
                    "Download updated workbook",
                    data=out,
                    file_name=uploaded.name.replace(".xlsx", "_DYNAMIC_FORMULAS.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.download_button(
                    "Download mapping JSON",
                    data=cfg_json(cfg),
                    file_name="mapping_config_dynamic_v1.json",
                    mime="application/json",
                )
            except Exception as exc:
                st.error(f"Export failed: {exc}")

    with st.expander("Current Mapping JSON"):
        st.code(cfg_json(cfg), language="json")


if __name__ == "__main__":
    main()
