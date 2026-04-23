import streamlit as st
import pandas as pd

st.title("Diamond Automation Tool")

cost_file = st.file_uploader("Upload Cost File", type=["xlsx"])
panding_file = st.file_uploader("Upload Panding File", type=["xlsx"])
lab_file = st.file_uploader("Upload Lab File", type=["xlsx"])

if cost_file and panding_file and lab_file:

    # ================= READ FILES =================
    cost = pd.read_excel(cost_file)
    panding = pd.read_excel(panding_file)

    # IMPORTANT: Lab file header fix (your file starts from row 3)
    lab = pd.read_excel(lab_file, header=2)

    # Clean column names
    cost.columns = cost.columns.str.strip()
    panding.columns = panding.columns.str.strip()
    lab.columns = lab.columns.str.strip()

    # ================= COST CLEAN =================
    cost = cost[["Lot #", "Shape", "Color", "Clarity", "Cts.", "Lab"]]
    cost = cost[cost["Lab"].isin(["GIA", "IGI", "GCAL"])]

    # ================= PANDING MERGE =================
    panding = panding[["Lot #", "Status"]]
    cost = cost.merge(panding, on="Lot #", how="left")

    # ================= LAB FIX =================

    # Show columns (for debug)
    st.write("Lab Columns:", lab.columns)

    # Find Stock column (contains 'Stock')
    stock_col = [c for c in lab.columns if "stock" in c.lower()][0]

    # Find Days column (contains 'old')
    days_col = [c for c in lab.columns if "old" in c.lower()][0]

    lab = lab[[stock_col, days_col]]

    lab = lab.rename(columns={
        stock_col: "Lot #",
        days_col: "No of Days"
    })

    # ================= MERGE =================
    cost = cost.merge(lab, on="Lot #", how="left")

    # ================= OUTPUT =================
    st.success("Done ✅")
    st.dataframe(cost)

    # Download
    output = cost.to_excel(index=False)
    st.download_button("Download Final File", output, file_name="Final_Output.xlsx")
