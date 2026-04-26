import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl.styles import Font

st.title("Diamond Tool")

# Upload files
cost_file = st.file_uploader("Upload Cost File", type=["xlsx"])
panding_file = st.file_uploader("Upload Pending File", type=["xlsx"])
lab_file = st.file_uploader("Upload Lab File", type=["xls", "xlsx"])

if cost_file and panding_file and lab_file:

    # ================= READ FILES =================
    cost = pd.read_excel(cost_file)
    panding = pd.read_excel(panding_file)
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

    # Keep only GIA / IGI / GCAL
    cost = cost[cost["Lab"].isin(["GIA", "IGI", "GCAL"])]

    # ================= COLOR FILTER =================
    valid_colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
    cost["Color"] = cost["Color"].astype(str).str.strip()
    cost = cost[cost["Color"].isin(valid_colors)]

    # ================= VP SERIES DELETE =================
    cost["Lot #"] = cost["Lot #"].astype(str).str.strip()
    cost = cost[
        ~cost["Lot #"].str.upper().str.startswith("VP")
    ]

    # ================= QUALITY FIX =================
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
    panding["Customer"] = (
        panding["Customer"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    panding["Status"] = (
        panding["Status"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    panding.loc[
        (
            (panding["Customer"] == "GOODS IN TRANSIT FROM OVERSEAS") |
            (panding["Customer"] == "GOODS IN OFFICE - PARCEL PAPERS BEING MADE")
        ) &
        (panding["Status"].str.upper() == "ONMEMO"),
        "Status"
    ] = "Inhand"

    # ================= PENDING MERGE =================
    panding = panding[["Lot #", "Status"]]
    cost = cost.merge(panding, on="Lot #", how="left")

    # ================= LAB FILE CLEAN =================
    stock_col = [c for c in lab.columns if "stock" in c.lower()][0]
    days_col = [c for c in lab.columns if "old" in c.lower()][0]

    lab = lab[[stock_col, days_col]]

    lab = lab.rename(columns={
        stock_col: "Lot #",
        days_col: "No of Days"
    })

    # ================= LAB MERGE =================
    cost = cost.merge(lab, on="Lot #", how="left")

    # ================= NO OF DAYS FIX =================
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

    # ================= SIZE GROUP =================
    def get_size_grp(cts):
        if pd.isna(cts):
            return ""

        cts = float(cts)

        if 0.30 <= cts <= 0.39:
            return "0.30 - 0.39"
        elif 0.40 <= cts <= 0.49:
            return "0.40 - 0.49"
        elif 0.50 <= cts <= 0.59:
            return "0.50 - 0.59"
        elif 0.60 <= cts <= 0.69:
            return "0.60 - 0.69"
        elif 0.70 <= cts <= 0.79:
            return "0.70 - 0.79"
        elif 0.80 <= cts <= 0.89:
            return "0.80 - 0.89"
        elif 0.90 <= cts <= 0.99:
            return "0.90 - 0.99"
        elif 1.00 <= cts <= 1.05:
            return "1.00 - 1.05"
        else:
            return ""

    cost["Size Grp"] = cost["Cts."].apply(get_size_grp)

    # ================= EXTRA COLUMNS =================
    cost["UPDATED PRICE"] = ""
    cost["DIFFERENCE"] = ""
    cost["Cost Amt"] = ""
    cost["Sale Amt"] = ""
    cost["Differance"] = ""

        # ================= FINAL FORMAT =================
    cost = cost[[
        "Lot #",
        "Status",
        "Shape",
        "Color",
        "Clarity",
        "Cts.",
        "Size Grp",
        "No of Days",
        "Price / Cts",
        "Cost / Cts.",
        "GIA #",
        "Lab",
        "Quality",
        "UPDATED PRICE",
        "DIFFERENCE",
        "Cost Amt",
        "Sale Amt",
        "Differance"
    ]]

    # ================= TOTAL COUNT =================
    st.success("Processing Completed Successfully ✅")
    total_diamond = len(cost)

    st.markdown(f"## Total Diamonds: {total_diamond}")
    st.markdown("---")

    # ================= OUTPUT =================
    st.dataframe(cost)

    # ================= DOWNLOAD =================
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
