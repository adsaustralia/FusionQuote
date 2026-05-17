import re
from io import BytesIO
from copy import copy

import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

st.set_page_config(page_title="Excel Formula Fusion", layout="wide")
st.title("Excel Formula Fusion – Sheet Picker + Formula Builder")
st.caption("Build formula-based DS/SS lookup and country-excluded quantity rows without changing the original input cells.")


def norm_sheet_name(name: str) -> str:
    return name.replace("'", "''")


def safe_cell(sheet_name: str, cell_ref: str) -> str:
    return f"'{norm_sheet_name(sheet_name)}'!{cell_ref}"


def parse_ignore_countries(raw: str):
    return [x.strip().upper() for x in re.split(r"[,;\n]+", raw or "") if x.strip()]


def copy_row_style(ws, source_row: int, target_row: int, min_col: int, max_col: int):
    for col in range(min_col, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.font:
            dst.font = copy(src.font)
        if src.border:
            dst.border = copy(src.border)


def ds_ss_formula(ref_sheet, current_col_letter, name_row, size_row, ref_name_col, ref_size_col, ref_ds_col, ref_start_row, ref_end_row):
    # Match both name/artwork and size. Uses INDEX/MATCH so workbook remains formula based and editable in Excel.
    name_ref = f"{current_col_letter}${name_row}"
    size_ref = f"{current_col_letter}${size_row}"
    ref = norm_sheet_name(ref_sheet)
    return (
        f'=IFERROR(INDEX(\'{ref}\'!${ref_ds_col}${ref_start_row}:${ref_ds_col}${ref_end_row},'
        f'MATCH(1,INDEX((\'{ref}\'!${ref_name_col}${ref_start_row}:${ref_name_col}${ref_end_row}={name_ref})*'
        f'(\'{ref}\'!${ref_size_col}${ref_start_row}:${ref_size_col}${ref_end_row}={size_ref}),0),0)),"")'
    )


def country_excluded_qty_formula(current_col_letter, data_start_row, data_end_row, country_col, ignored_countries):
    qty_range = f"{current_col_letter}${data_start_row}:{current_col_letter}${data_end_row}"
    country_range = f"${country_col}${data_start_row}:${country_col}${data_end_row}"
    if not ignored_countries:
        return f"=SUM({qty_range})"
    # Formula: SUM all qty where country does not match each ignored country.
    # For multiple countries, SUMPRODUCT avoids fragile nested SUMIFS subtraction.
    conditions = "*".join([f"(--(UPPER({country_range})<>\"{c}\"))" for c in ignored_countries])
    return f"=SUMPRODUCT(({qty_range})*{conditions})"


def apply_formulas(wb, cfg):
    ws = wb[cfg["working_sheet"]]

    start_col = column_index_from_string(cfg["first_item_col"])
    end_col = column_index_from_string(cfg["last_item_col"])
    copy_row_style(ws, cfg["qty_source_row"], cfg["clean_qty_output_row"], start_col, end_col)
    copy_row_style(ws, cfg["size_row"], cfg["ds_output_row"], start_col, end_col)

    ws.cell(cfg["clean_qty_output_row"], max(1, start_col - 1)).value = cfg["clean_qty_label"]
    ws.cell(cfg["ds_output_row"], max(1, start_col - 1)).value = cfg["ds_label"]

    for col in range(start_col, end_col + 1):
        col_letter = get_column_letter(col)
        ws.cell(cfg["clean_qty_output_row"], col).value = country_excluded_qty_formula(
            col_letter,
            cfg["store_start_row"],
            cfg["store_end_row"],
            cfg["country_col"],
            cfg["ignored_countries"],
        )
        ws.cell(cfg["ds_output_row"], col).value = ds_ss_formula(
            cfg["reference_sheet"],
            col_letter,
            cfg["name_row"],
            cfg["size_row"],
            cfg["ref_name_col"],
            cfg["ref_size_col"],
            cfg["ref_ds_col"],
            cfg["ref_start_row"],
            cfg["ref_end_row"],
        )

    return wb


uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx"])

if uploaded:
    file_bytes = uploaded.read()
    wb_preview = load_workbook(BytesIO(file_bytes), data_only=False)
    sheets = wb_preview.sheetnames

    st.subheader("1) Choose sheets")
    c1, c2 = st.columns(2)
    with c1:
        working_sheet = st.selectbox(
            "Working sheet", sheets, index=sheets.index("DL ANZ ALLOCATION") if "DL ANZ ALLOCATION" in sheets else 0
        )
    with c2:
        reference_sheet = st.selectbox(
            "Reference sheet for DS/SS", sheets, index=sheets.index("PRINT DB") if "PRINT DB" in sheets else 0
        )

    st.subheader("2) Working sheet mapping")
    st.caption("Defaults are set for your HOKA Clifton 11 file. Change these for other workbooks.")
    a, b, c, d = st.columns(4)
    with a:
        first_item_col = st.text_input("First item column", "AC").upper()
        name_row = st.number_input("Item/name row", min_value=1, value=4)
    with b:
        last_item_col = st.text_input("Last item column", "EU").upper()
        size_row = st.number_input("Size row", min_value=1, value=5)
    with c:
        qty_source_row = st.number_input("Original total qty row", min_value=1, value=7)
        country_col = st.text_input("Country column", "I").upper()
    with d:
        store_start_row = st.number_input("Store data start row", min_value=1, value=8)
        store_end_row = st.number_input("Store data end row", min_value=1, value=167)

    st.subheader("3) Reference sheet mapping")
    r1, r2, r3, r4, r5 = st.columns(5)
    with r1:
        ref_name_col = st.text_input("Reference artwork/name column", "C").upper()
    with r2:
        ref_size_col = st.text_input("Reference size column", "E").upper()
    with r3:
        ref_ds_col = st.text_input("Reference DS/SS column", "F").upper()
    with r4:
        ref_start_row = st.number_input("Reference start row", min_value=1, value=12)
    with r5:
        ref_end_row = st.number_input("Reference end row", min_value=1, value=141)

    st.subheader("4) Formula output")
    f1, f2, f3 = st.columns(3)
    with f1:
        ignored_raw = st.text_area("Countries to ignore", "NZ", help="Use comma or new line separated country codes/names.")
    with f2:
        clean_qty_output_row = st.number_input("Output row for country-excluded qty", min_value=1, value=169)
        clean_qty_label = st.text_input("Clean qty label", "AUS Qty (formula)")
    with f3:
        ds_output_row = st.number_input("Output row for DS/SS lookup", min_value=1, value=168)
        ds_label = st.text_input("DS/SS label", "DS/SS from PRINT DB")

    ignored_countries = parse_ignore_countries(ignored_raw)

    st.subheader("5) Preview formulas")
    first_col_letter = first_item_col
    st.code(
        "Clean qty formula example:\n"
        + country_excluded_qty_formula(first_col_letter, store_start_row, store_end_row, country_col, ignored_countries)
        + "\n\nDS/SS lookup formula example:\n"
        + ds_ss_formula(reference_sheet, first_col_letter, name_row, size_row, ref_name_col, ref_size_col, ref_ds_col, ref_start_row, ref_end_row),
        language="excel",
    )

    if st.button("Generate formula-based workbook", type="primary"):
        try:
            wb = load_workbook(BytesIO(file_bytes), data_only=False)
            cfg = {
                "working_sheet": working_sheet,
                "reference_sheet": reference_sheet,
                "first_item_col": first_item_col,
                "last_item_col": last_item_col,
                "name_row": int(name_row),
                "size_row": int(size_row),
                "qty_source_row": int(qty_source_row),
                "country_col": country_col,
                "store_start_row": int(store_start_row),
                "store_end_row": int(store_end_row),
                "ref_name_col": ref_name_col,
                "ref_size_col": ref_size_col,
                "ref_ds_col": ref_ds_col,
                "ref_start_row": int(ref_start_row),
                "ref_end_row": int(ref_end_row),
                "ignored_countries": ignored_countries,
                "clean_qty_output_row": int(clean_qty_output_row),
                "ds_output_row": int(ds_output_row),
                "clean_qty_label": clean_qty_label,
                "ds_label": ds_label,
            }
            wb = apply_formulas(wb, cfg)
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            st.success("Workbook generated. Open in Excel to calculate/refresh formulas.")
            st.download_button(
                "Download updated workbook",
                data=output,
                file_name="formula_based_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Failed to generate workbook: {e}")
else:
    st.info("Upload your workbook to start.")
