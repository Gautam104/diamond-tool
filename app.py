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

    cost = cost[["Lot #", "Shape", "Color", "Clarity", "Cts.", "Lab"]]
    cost = cost[cost["Lab"].isin(["GIA", "IGI", "GCAL"])]

    panding = panding[["Lot #", "Status"]]
    cost = cost.merge(panding, on="Lot #", how="left")

    lab = lab[["Stock#", "How old stone in stock"]]
    lab = lab.rename(columns={"Stock#": "Lot #"})
    cost = cost.merge(lab, on="Lot #", how="left")

    cost.rename(columns={"How old stone in stock": "No of Days"}, inplace=True)

    st.dataframe(cost)

    output = cost.to_excel(index=False)
    st.download_button("Download File", output, file_name="Final.xlsx")
