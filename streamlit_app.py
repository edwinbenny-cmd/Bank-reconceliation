"""
streamlit_app.py
Run with:  streamlit run streamlit_app.py

Lets a colleague upload the raw Tally-style ledger export and instantly
get back a clean, filterable table (plus a download button for the
structured .xlsx).
"""

import io
import streamlit as st
import pandas as pd
from ledger_parser import parse_workbook

st.set_page_config(page_title="Ledger Structurer", layout="wide")
st.title("📒 Ledger → Structured Table")

st.write(
    "Upload the raw ledger export (the multi-sheet CASH/BANK workbook) and "
    "get back one flat, filterable table with narration split into its own column."
)

uploaded = st.file_uploader("Upload the raw ledger .xlsx", type=["xlsx"])

if uploaded is not None:
    # parse_workbook needs a path or a file-like openpyxl can read; a BytesIO works.
    df = parse_workbook(io.BytesIO(uploaded.read()))
    # Keep Vch No as text so Streamlit/pandas don't turn "805" into 805.0
    df["Vch No"] = df["Vch No"].astype("string")

    st.success(f"Parsed {len(df)} rows across {df['Sheet'].nunique()} sheet(s) "
               f"and {df['Ledger'].nunique()} ledger(s).")

    col1, col2, col3 = st.columns(3)
    with col1:
        sheet_filter = st.multiselect("Sheet", sorted(df["Sheet"].dropna().unique()))
    with col2:
        ledger_filter = st.multiselect("Ledger", sorted(df["Ledger"].dropna().unique()))
    with col3:
        rowtype_filter = st.multiselect("Row Type", sorted(df["Row Type"].dropna().unique()))

    filtered = df.copy()
    if sheet_filter:
        filtered = filtered[filtered["Sheet"].isin(sheet_filter)]
    if ledger_filter:
        filtered = filtered[filtered["Ledger"].isin(ledger_filter)]
    if rowtype_filter:
        filtered = filtered[filtered["Row Type"].isin(rowtype_filter)]

    search = st.text_input("Search narration / particulars")
    if search:
        mask = (
            filtered["Narration"].fillna("").str.contains(search, case=False)
            | filtered["Particulars"].fillna("").str.contains(search, case=False)
        )
        filtered = filtered[mask]

    st.dataframe(filtered, use_container_width=True, height=600)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        filtered.to_excel(writer, sheet_name="Structured", index=False)
    st.download_button(
        "⬇️ Download structured .xlsx",
        data=buf.getvalue(),
        file_name="structured_ledger.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Waiting for a file...")
