import streamlit as st
import pandas as pd

st.title("Diamond Automation Tool")

# Upload files
cost_file = st.file_uploader("Upload Cost File", type=["xlsx"])
panding_file = st.file_uploader("Upload Panding File", type=["xlsx"])
lab_file = st.file_uploader("Upload Lab File", type=["xlsx"])

if cost_file and panding_file and lab_file:

    # Read files
    cost = pd.read_excel(cost_file)
    panding = pd.read_excel(panding_file)
    lab = pd.read_excel(lab_file)

    # Clean column names (remove extra spaces)
    cost.columns = cost.columns.str.strip()
    panding.columns = panding.columns.str.strip()
    lab.columns = lab.columns.str.strip()

    # ===================== COST CLEAN =====================
    cost = cost[["Lot #", "Shape", "Color", "Clarity", "Cts.", "Lab"]]
    cost = cost[cost["Lab"].isin(["GIA", "IGI", "GCAL"])]

    # ===================== PANDING MERGE =====================
    panding = panding[["Lot #", "Status"]]
    cost = cost.merge(panding, on="Lot #", how="left")

    # ===================== LAB MERGE =====================
    # Show columns for safety
    st.write("Lab File Columns:", list(lab.columns))

    # Use your exact column names
    lab = lab[["Stock#", "No of Days / How old stone in stock"]]

    lab = lab.rename(columns={
        "Stock#": "Lot #",
        "No of Days / How old stone in stock": "No of Days"
    })

    cost = cost.merge(lab, on="Lot #", how="left")

    # ===================== OUTPUT =====================
    st.success("Done ✅")
    st.dataframe(cost)

    # Download file
    output = cost.to_excel(index=False)
    st.download_button("Download Final File", output, file_name="Final_Output.xlsx")
