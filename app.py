import streamlit as st
import pandas as pd

st.title("Diamond Automation Tool")

cost_file = st.file_uploader("Upload Cost File", type=["xlsx"])
panding_file = st.file_uploader("Upload Panding File", type=["xlsx"])
lab_file = st.file_uploader("Upload Lab File", type=["xlsx"])

if cost_file and panding_file and lab_file:

    cost = pd.read_excel(cost_file)
    panding = pd.read_excel(panding_file)
    lab = pd.read_excel(lab_file)

    # Clean column names (remove spaces)
    cost.columns = cost.columns.str.strip()
    panding.columns = panding.columns.str.strip()
    lab.columns = lab.columns.str.strip()

    # --- COST CLEAN ---
    cost = cost[["Lot #", "Shape", "Color", "Clarity", "Cts.", "Lab"]]
    cost = cost[cost["Lab"].isin(["GIA", "IGI", "GCAL"])]

    # --- PANDING MERGE ---
    panding = panding[["Lot #", "Status"]]
    cost = cost.merge(panding, on="Lot #", how="left")

    # --- LAB MERGE (SAFE WAY) ---
    # Auto detect columns
    stock_col = [col for col in lab.columns if "stock" in col.lower()][0]
    days_col = [col for col in lab.columns if "old" in col.lower() or "day" in col.lower()][0]

    lab = lab[[stock_col, days_col]]
    lab = lab.rename(columns={stock_col: "Lot #", days_col: "No of Days"})

    cost = cost.merge(lab, on="Lot #", how="left")

    st.success("Done ✅")
    st.dataframe(cost)

    output = cost.to_excel(index=False)
    st.download_button("Download File", output, file_name="Final.xlsx")
