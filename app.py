import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl.styles import Font

st.title("Diamond Tool")

# ================= SIZE GROUP FUNCTION =================

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
    elif 1.06 <= cts <= 1.10:
        return "1.06 - 1.10"
    elif 1.11 <= cts <= 1.49:
        return "1.11 - 1.49"
    elif 1.50 <= cts <= 1.55:
        return "1.50 - 1.55"
    elif 1.56 <= cts <= 1.59:
        return "1.56 - 1.59"
    elif 1.60 <= cts <= 1.99:
        return "1.60 - 1.99"
    elif 2.00 <= cts <= 2.05:
        return "2.00 - 2.05"
    elif 2.06 <= cts <= 2.10:
        return "2.06 - 2.10"
    elif 2.11 <= cts <= 2.49:
        return "2.11 - 2.49"
    elif 2.50 <= cts <= 2.55:
        return "2.50 - 2.55"
    elif 2.56 <= cts <= 2.59:
        return "2.56 - 2.59"
    elif 2.60 <= cts <= 2.99:
        return "2.60 - 2.99"
    elif 3.00 <= cts <= 3.05:
        return "3.00 - 3.05"
    elif 3.06 <= cts <= 3.10:
        return "3.06 - 3.10"
    elif 3.11 <= cts <= 3.49:
        return "3.11 - 3.49"
    elif 3.50 <= cts <= 3.55:
        return "3.50 - 3.55"
    elif 3.56 <= cts <= 3.59:
        return "3.56 - 3.59"
    elif 3.60 <= cts <= 3.99:
        return "3.60 - 3.99"
    elif 4.00 <= cts <= 4.05:
        return "4.00 - 4.05"
    elif 4.06 <= cts <= 4.10:
        return "4.06 - 4.10"
    elif 4.11 <= cts <= 4.49:
        return "4.11 - 4.49"
    elif 4.50 <= cts <= 4.55:
        return "4.50 - 4.55"
    elif 4.56 <= cts <= 4.59:
        return "4.56 - 4.59"
    elif 4.60 <= cts <= 4.99:
        return "4.60 - 4.99"
    elif 5.00 <= cts <= 5.49:
        return "5.00 - 5.49"
    elif 5.50 <= cts <= 5.99:
        return "5.50 - 5.99"
    elif 6.00 <= cts <= 6.99:
        return "6.00 - 6.99"
    elif 7.00 <= cts <= 7.99:
        return "7.00 - 7.99"
    elif 8.00 <= cts <= 8.99:
        return "8.00 - 8.99"
    elif 9.00 <= cts <= 9.99:
        return "9.00 - 9.99"
    elif 10.00 <= cts <= 10.99:
        return "10.00 - 10.99"
    elif 11.00 <= cts <= 11.99:
        return "11.00 - 11.99"
    elif 12.00 <= cts <= 12.99:
        return "12.00 - 12.99"
    else:
        return ""

# ================= FILE UPLOAD =================

cost_file = st.file_uploader("Upload Cost File", type=["xlsx"])
panding_file = st.file_uploader("Upload Pending File", type=["xlsx"])
lab_file = st.file_uploader("Upload Lab File", type=["xls", "xlsx"])

if cost_file and panding_file and lab_file:

    # READ FILES
    cost = pd.read_excel(cost_file)
    panding = pd.read_excel(panding_file)

    if lab_file.name.endswith(".xls"):
        lab = pd.read_excel(lab_file, header=2, engine="xlrd")
    else:
        lab = pd.read_excel(lab_file, header=2, engine="openpyxl")

    # CLEAN COLUMN NAMES
    cost.columns = cost.columns.str.strip()
    panding.columns = panding.columns.str.strip()
    lab.columns = lab.columns.str.strip()

    # COST FILE REQUIRED COLUMNS
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

    # LAB FILTER
    cost = cost[cost["Lab"].isin(["GIA", "IGI", "GCAL"])]

    # COLOR FILTER
    valid_colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
    cost["Color"] = cost["Color"].astype(str).str.strip()
    cost = cost[cost["Color"].isin(valid_colors)]

    # REMOVE VP SERIES
    cost["Lot #"] = cost["Lot #"].astype(str).str.strip()
    cost = cost[
        ~cost["Lot #"].str.upper().str.startswith("VP")
    ]

    # QUALITY FIX
    cost["Quality"] = cost["Quality"].fillna("").astype(str).str.strip()
    cost["Rapnet Note"] = cost["Rapnet Note"].fillna("").astype(str).str.upper()

    cost["Quality"] = cost["Quality"].replace(
        ["Blank", "blank", "BLANK", "nan", "NaN"],
        ""
    )

    cost.loc[
        (cost["Quality"] == "") &
        (cost["Rapnet Note"].str.contains("CVD", na=False)),
        "Quality"
    ] = "CVD"

    cost.loc[
        (cost["Quality"] == "") &
        (cost["Rapnet Note"].str.contains("HPHT", na=False)),
        "Quality"
    ] = "HPHT"

    # PENDING FILE FIX
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

    # MERGE STATUS
    panding = panding[["Lot #", "Status"]]
    cost = cost.merge(panding, on="Lot #", how="left")

    # LAB FILE CLEAN
    stock_col = [c for c in lab.columns if "stock" in c.lower()][0]
    days_col = [c for c in lab.columns if "old" in c.lower()][0]

    lab = lab[[stock_col, days_col]]

    lab = lab.rename(columns={
        stock_col: "Lot #",
        days_col: "No of Days"
    })

    # MERGE LAB
    cost = cost.merge(lab, on="Lot #", how="left")

    # NO OF DAYS FIX
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

    # SIZE GROUP
    cost["Cts."] = pd.to_numeric(cost["Cts."], errors="coerce")
    cost["Size Grp"] = cost["Cts."].apply(get_size_grp)

    # EXTRA HEADER COLUMNS ONLY
    cost["UPDATED PRICE"] = ""
    cost["DIFFERENCE"] = ""
    cost["Cost Amt"] = ""
    cost["Sale Amt"] = ""
    cost["Differance"] = ""

    # FINAL FORMAT
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

    # OUTPUT
    st.success("Processing Completed Successfully ✅")

    total_diamond = len(cost)
    st.markdown(f"## Total Diamonds: {total_diamond}")
    st.markdown("---")

    st.dataframe(cost)

    # DOWNLOAD EXCEL
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
