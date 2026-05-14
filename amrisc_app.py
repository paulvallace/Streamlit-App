# Amrisc_SOV_app.py
# Streamlit UI wrapper for AmRisc SOV processing
# Revised to fix Arrow / float64 None crashes

import io
import re
from typing import Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# =========================
# Page Setup
# =========================
st.set_page_config(page_title="AmRisc SOV Builder", page_icon="📄", layout="wide")
st.title("📄 AmRisc SOV Builder")
st.caption("Normalize a client SOV into the AmRisc SOV-APP template.")

# =========================
# Helpers
# =========================
def normalize_alias(x: Optional[str]) -> str:
    if x is None:
        return ""
    s = str(x).strip().lower().replace("&", "and")
    s = re.sub(r"[\*\(\)]", "", s)
    s = re.sub(r"\s+", " ", s)
    return re.sub(r"[^a-z0-9]", "", s)

def split_lines_safe(s: Optional[str]):
    if not isinstance(s, str):
        return []
    return [p.strip() for p in re.split(r"[\r\n]+", s) if p.strip()]

def is_blank_series(series: pd.Series) -> pd.Series:
    return series.isna() | (series.astype(str).str.strip() == "")

def split_city_state_zip_col(series: pd.Series) -> pd.DataFrame:
    s = series.astype("string").str.strip()
    pat = r"^\s*(?P<City>[^,]+?)\s*,\s*(?P<State>[A-Za-z]{2})\s*,\s*(?P<Zip>\d{5}(?:-\d{4})?)\s*$"
    parts = s.str.extract(pat)
    if "State" in parts:
        parts["State"] = parts["State"].str.upper()
    return parts

def first_empty_row_under(ws, column_index: int, start: int = 3) -> int:
    r = start
    while True:
        if ws.cell(row=r, column=column_index).value in (None, ""):
            return r
        r += 1

def safe_write(ws, row: int, col: int, value):
    if pd.isna(value):
        value = None
    c = ws.cell(row=row, column=col)
    if isinstance(c, MergedCell):
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                ws.cell(row=rng.min_row, column=rng.min_col, value=value)
                return
    else:
        c.value = value

def non_empty(s) -> bool:
    return isinstance(s, str) and s.strip() != ""

def find_sprinkler_col(df: pd.DataFrame) -> Optional[str]:
    for c in df.columns:
        if "sprink" in str(c).lower():
            return c
    return None

# =========================
# SAFE NUMERIC NORMALIZER ✅
# =========================
def coerce_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    int_cols = [
        "*# of Bldgs", "*# of Stories", "*Orig Year Built",
        "Yr Bldg upgraded - Major Exterior Update (mandatory if >25 yrs old)",
        "*Year Roof covering last fully replaced",
        "*# of Units", "*Square Footage",
        "Percent Sprinklered", "ISO Prot Class"
    ]

    float_cols = [
        "*Real Property Value ($)",
        "Personal Property Value ($)",
        "M&E (Complete M&E Tech Summary Sheet)",
        "Other Value $ (outdoor prop & Eqpt must be sch'd)",
        "BI/Rental Income ($)",
        "% Occupied"
    ]

    for c in int_cols:
        if c in df.columns:
            s = (
                df[c]
                .astype("string")
                .str.replace(",", "", regex=False)
                .str.replace("%", "", regex=False)
                .str.strip()
            )
            df[c] = pd.to_numeric(s, errors="coerce").round().astype("Int64")

    for c in float_cols:
        if c in df.columns:
            s = (
                df[c]
                .astype("string")
                .str.replace(",", "", regex=False)
                .str.replace("%", "", regex=False)
                .str.strip()
            )
            df[c] = pd.to_numeric(s, errors="coerce")

    return df

# =========================
# TARGETS
# =========================
TARGETS_IN_ORDER = [
    "* Bldg No.", "*Property Type", "Location Name", "AddressNum",
    "*Street Address", "*City", "*State Code", "*Zip", "County",
    "Is Prop within 1000 ft of saltwater", "*# of Bldgs",
    "*ISO Const",
    "Construction Description (provide further details on construction features)",
    "*# of Stories", "*Orig Year Built",
    "Yr Bldg upgraded - Major Exterior Update (mandatory if >25 yrs old)",
    "*Year Roof covering last fully replaced",
    "*Real Property Value ($)", "Personal Property Value ($)",
    "M&E (Complete M&E Tech Summary Sheet)",
    "Other Value $ (outdoor prop & Eqpt must be sch'd)",
    "BI/Rental Income ($)", "*Occupancy Description",
    "Is Prop Mgd or Owned?", "Date Added to Sched.",
    "*# of Units", "*Square Footage", "% Occupied",
    "Percent Sprinklered", "Sprinklered (Y/N)",
    "ISO Prot Class", "Flood Zone",
]

RAW_COLUMN_MAPPING = {
    "address": "*Street Address",
    "street": "*Street Address",
    "city": "*City",
    "state": "*State Code",
    "zip": "*Zip",
    "county": "County",
    "building value": "*Real Property Value ($)",
    "bpp": "Personal Property Value ($)",
    "business income": "BI/Rental Income ($)",
    "square feet": "*Square Footage",
    "occupancy": "*Occupancy Description",
    "stories": "*# of Stories",
    "year built": "*Orig Year Built",
}

# =========================
# UI
# =========================
with st.sidebar:
    named_insured = st.text_input("Named Insured")
    source_sov = st.file_uploader("Upload Source SOV (.xlsx)", type=["xlsx"])
    template_path = st.text_input("Template path", value="AmRisc_SOV_Schedule.xlsx")
    template_sheet_name = st.text_input("Template sheet", value="SOV-APP")

process_button = st.button("🚀 Process SOV", use_container_width=True)

# =========================
# PROCESSING
# =========================
if process_button:
    try:
        src_df = pd.read_excel(source_sov, engine="openpyxl")

        new_data = pd.DataFrame(columns=TARGETS_IN_ORDER)

        src_cols_norm = {normalize_alias(c): c for c in src_df.columns}
        for src_key, tgt in RAW_COLUMN_MAPPING.items():
            k = normalize_alias(src_key)
            if k in src_cols_norm:
                new_data[tgt] = src_df[src_cols_norm[k]].values

        # AddressNum
        if "*Street Address" in new_data:
            new_data["AddressNum"] = new_data["*Street Address"].astype(str).str.extract(r"^(\d+)")

        # ✅ FIX: numeric coercion BEFORE preview
        new_data = coerce_numeric_columns(new_data)

        st.markdown("**Preview (safe):**")
        st.dataframe(new_data.head().astype("string"), use_container_width=True)

        # Template load
        wb = load_workbook(template_path)
        ws = wb[template_sheet_name]

        start_row = 3
        written = 0

        for _, r in new_data.iterrows():
            if not non_empty(r.get("*Street Address")):
                continue
            for col_idx, tgt in enumerate(TARGETS_IN_ORDER, start=1):
                safe_write(ws, start_row + written, col_idx, r.get(tgt))
            written += 1

        with io.BytesIO() as out:
            wb.save(out)
            out.seek(0)
            st.download_button(
                "⬇️ Download AmRisc SOV",
                data=out.getvalue(),
                file_name=f"{named_insured or 'Named Insured'} - Amrisc SOV.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.success(f"Complete ✅ Rows written: {written}")

    except Exception as e:
        st.error(f"Processing failed: {e}")
