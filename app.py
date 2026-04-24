import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Diamond Tool")

# Upload files
cost_file = st.file_uploader("Upload Cost File", type=["xlsx"])
panding_file = st.file_uploader("Upload Pending File", type=["xlsx"])

# Last file allow XLS + XLSX
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

    # ================= QUALITY FIX =================
    cost["Quality"] = cost["Quality"].fillna("").astype(str).str.strip()
    cost["Rapnet Note"] = cost["Rapnet Note"].fillna("").astype(str).str.upper()

    # Treat Blank also as empty
    cost["Quality"] = cost["Quality"].replace(
        ["Blank", "blank", "BLANK", "nan", "NaN"],
        ""
    )

    # Fill CVD from Rapnet Note
    cost.loc[
        (cost["Quality"] == "") &
        (cost["Rapnet Note"].str.contains("CVD", na=False)),
        "Quality"
    ] = "CVD"

    # Fill HPHT from Rapnet Note
    cost.loc[
        (cost["Quality"] == "") &
        (cost["Rapnet Note"].str.contains("HPHT", na=False)),
        "Quality"
    ] = "HPHT"

    # ================= PANDING FILE FIX =================
    # If Customer = GOODS IN TRANSIT FROM OVERSEAS
    # then change Status = OnMemo → Inhand

    panding["Customer"] = panding["Customer"].fillna("").astype(str).str.strip().str.upper()
    panding["Status"] = panding["Status"].fillna("").astype(str).str.strip()

    panding.loc[
        (panding["Customer"] == "GOODS IN TRANSIT FROM OVERSEAS") &
        (panding["Status"].str.upper() == "ONMEMO"),
        "Status"
    ] = "Inhand"

    # ================= PANDING MERGE =================
    panding = panding[["Lot #", "Status"]]
    cost = cost.merge(panding, on="Lot #", how="left")

    # ================= LAB GROWN FILE CLEAN =================

    # Auto detect Stock# column
    stock_col = [c for c in lab.columns if "stock" in c.lower()][0]

    # Auto detect How old stone column
    days_col = [c for c in lab.columns if "old" in c.lower()][0]

    lab = lab[[stock_col, days_col]]

    lab = lab.rename(columns={
        stock_col: "Lot #",
        days_col: "No of Days"
    })

    # ================= MERGE =================
    cost = cost.merge(lab, on="Lot #", how="left")

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
        "Lab",
        "Quality"
    ]]

    # ================= OUTPUT =================
    st.success("Done ✅")
    st.dataframe(cost)

    # ================= DOWNLOAD EXCEL WITH BOLD HEADER =================
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        cost.to_excel(writer, index=False, sheet_name="Final Output")

        worksheet = writer.sheets["Final Output"]

        # Make header bold
        from openpyxl.styles import Font

        for cell in worksheet[1]:
            cell.font = Font(bold=True)

    buffer.seek(0)

    st.download_button(
        label="Download Final Excel File",
        data=buffer,
        file_name="Final_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
