import streamlit as st
from excel_project import process_excel  # wrap your logic into a function

st.title("Daily Production Sheet Updater")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
date_input = st.text_input("Enter Date (e.g., 8th March)")
mwh_value = st.number_input("Enter MWh value", step=0.01)

if st.button("Process Report"):
    if uploaded_file:
        # Save to temp file, then pass to processing logic
        with open("temp_report.xlsx", "wb") as f:
            f.write(uploaded_file.read())
        
        result = process_excel("temp_report.xlsx", date_input, mwh_value)
        message = print("successfully updated")
        st.success(message)
