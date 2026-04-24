import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Diamond Automation Tool")

# Upload files
cost_file = st.file_uploader("Upload Cost File", type=["xlsx"])
panding_file = st.file_uploader("Upload Panding File", type=["xlsx"])
lab_file = st.file_uploader("Upload Lab File", type=["xlsx"])

if cost_file and panding_file and lab_file:

    # ================= READ FILES =================
    cost = pd.read_excel(cost_file)
    panding = pd.read_excel(panding_file)

    # Lab file has header in 3rd row
    lab = pd.read_excel(lab_file, header=2)

    # Clean column names
    cost.columns = cost.columns.str.strip()
    panding.columns = panding.columns.str.strip()
    lab.columns = lab.columns.str.strip()

    # ================= COST CLEAN =================
    cost = cost[["Lot #", "Shape", "Color", "Clarity", "Cts.", "Lab", "Quality","Price / Cts","Cost / Cts."]]
    cost = cost[cost["Lab"].isin(["GIA", "IGI", "GCAL"])]

    # ================= PANDING MERGE =================
    panding = panding[["Lot #", "Status"]]
    cost = cost.merge(panding, on="Lot #", how="left")

    # ================= LAB CLEAN =================

    # Find correct columns automatically
    stock_col = [c for c in lab.columns if "stock" in c.lower()][0]
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

    # ================= DOWNLOAD FIX =================
    buffer = BytesIO()
    cost.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)

    st.download_button(
        label="Download Final Excel File",
        data=buffer,
        file_name="Final_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
