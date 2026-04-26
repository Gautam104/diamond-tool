import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl.styles import Font

st.title("Diamond Tool")

# Upload files
cost_file = st.file_uploader("Upload Cost File", type=["xlsx"])
panding_file = st.file_uploader("Upload Pending File", type=["xlsx"])

# Lab file supports XLS + XLSX
lab_file = st.file_uploader(
    "Upload Lab File",
    type=["xls", "xlsx"]
)

if cost_file and panding_file and lab_file:

    # ================= READ FILES =================
    cost = pd.read_excel(cost_file)
    panding = pd.read_excel(panding_file)

    # Lab file header starts from row 3
    lab = pd.read_excel(lab_file, header=2)

    # Clean column names
    cost.columns = cost.columns.str.strip()
    panding.columns = panding.columns.str.strip()
    lab.columns = lab.columns.str.strip()

    # ================= COST CLEAN =================
cost = cost[[
    "Lot #",
    "Shape",
    "Color",
    "Clarity",
    "Cts.",
    "GIA #",
    "Lab",
    "Quality",
    "Price / Cts",
    "Cost / Cts.",
    "Rapnet Note"
]]

    # ================= LAB FILTER =================
    # Keep only GIA / IGI / GCAL
    cost = cost[cost["Lab"].isin(["GIA", "IGI", "GCAL"])]

    # ================= COLOR FILTER =================
    # Keep only valid one-letter colors
    valid_colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]

    cost["Color"] = cost["Color"].astype(str).str.strip()
    cost = cost[cost["Color"].isin(valid_colors)]

    # ================= VP SERIES DELETE =================
    # Delete rows where Lot # starts with VP

    cost["Lot #"] = cost["Lot #"].astype(str).str.strip()

    cost = cost[
        ~cost["Lot #"].str.upper().str.startswith("VP")
    ]

    # ================= QUALITY FIX =================
    # Fill Quality from Rapnet Note if blank

    cost["Quality"] = cost["Quality"].fillna("").astype(str).str.strip()
    cost["Rapnet Note"] = cost["Rapnet Note"].fillna("").astype(str).str.upper()

    cost["Quality"] = cost["Quality"].replace(
        ["Blank", "blank", "BLANK", "nan", "NaN"],
        ""
    )

    # Fill CVD
    cost.loc[
        (cost["Quality"] == "") &
        (cost["Rapnet Note"].str.contains("CVD", na=False)),
        "Quality"
    ] = "CVD"

    # Fill HPHT
    cost.loc[
        (cost["Quality"] == "") &
        (cost["Rapnet Note"].str.contains("HPHT", na=False)),
        "Quality"
    ] = "HPHT"

    # ================= PENDING FILE FIX =================

    panding["Customer"] = panding["Customer"].fillna("").astype(str).str.strip().str.upper()
    panding["Status"] = panding["Status"].fillna("").astype(str).str.strip()

    panding.loc[
        (
            (
                panding["Customer"] == "GOODS IN TRANSIT FROM OVERSEAS"
            ) |
            (
                panding["Customer"] == "GOODS IN OFFICE - PARCEL PAPERS BEING MADE"
            )
        ) &
        (
            panding["Status"].str.upper() == "ONMEMO"
        ),
        "Status"
    ] = "Inhand"

    # ================= PENDING MERGE =================

    panding = panding[["Lot #", "Status"]]
    cost = cost.merge(panding, on="Lot #", how="left")

    # ================= LAB GROWN FILE CLEAN =================

    stock_col = [c for c in lab.columns if "stock" in c.lower()][0]
    days_col = [c for c in lab.columns if "old" in c.lower()][0]

    lab = lab[[stock_col, days_col]]

    lab = lab.rename(columns={
        stock_col: "Lot #",
        days_col: "No of Days"
    })

    # ================= MERGE =================

    cost = cost.merge(lab, on="Lot #", how="left")

    # ================= NO OF DAYS FIX =================
    # DM / DC series + No of Days = 0 → blank

    cost["Lot #"] = cost["Lot #"].astype(str).str.strip()
    cost["No of Days"] = pd.to_numeric(cost["No of Days"], errors="coerce")

    cost.loc[
        (
            cost["Lot #"].str.upper().str.startswith(("DM", "DC"))
        ) &
        (
            cost["No of Days"] == 0
        ),
        "No of Days"
    ] = np.nan

    # ================= FINAL FORMAT =================

   cost = cost[[
    "Lot #",
    "Status",
    "Shape",
    "Color",
    "Clarity",
    "Cts.",
    "No of Days",
    "Price / Cts",
    "Cost / Cts.",
    "GIA #",
    "Lab",
    "Quality"
]]

    # ================= SIMPLE TOTAL DIAMOND COUNT =================

    st.success("Processing Completed Successfully ✅")

    total_diamond = len(cost)

    st.markdown(f"## Total Diamonds: {total_diamond}")
    st.markdown("---")

    # ================= OUTPUT =================

    st.dataframe(cost)

    # ================= DOWNLOAD EXCEL WITH BOLD HEADER =================

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        cost.to_excel(writer, index=False, sheet_name="Final Output")

        worksheet = writer.sheets["Final Output"]

        for cell in worksheet[1]:
            cell.font = Font(bold=True)

    buffer.seek(0)

    st.download_button(
        label="Download Final Excel File",
        data=buffer,
        file_name="Final_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
