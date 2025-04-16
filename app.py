import streamlit as st
from excel_project import generate_daily_report
from datetime import datetime
import pandas as pd

st.title("Daily Production Report Generator")

# File uploaders
uploaded_source = st.file_uploader("Upload Gross Generation File", type="xlsx")
uploaded_template = st.file_uploader("Upload Daily production Report Template", type="xlsx") 
uploaded_annual = st.file_uploader("Upload Gross Gen Summary File", type="xlsx")

# Date input
report_date = st.date_input("Select Report Date", datetime.today())

if st.button("Generate Report" , type ="primary"):
    if not all([uploaded_source, uploaded_template, uploaded_annual]):
        str.error("Please upload all required files")
    else:
        try:
            with st.spinner("Generating Report..."):
            # Call function with CORRECT parameter names
                report_data = generate_daily_report(
                    source_file=uploaded_source,
                    report_file=uploaded_template,
                    gross_wb=uploaded_annual,
                    target_date_input=report_date
            )
            
            st.success("Report generated successfully!")
            st.download_button(
                label="Download Updated Report",
                data=report_data,
                file_name=f"Daily_production_report_{report_date.strftime('%B_%Y')}_updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except ValueError as e:
            st.error(str(e))
            st.warning("Available dates in your source file:")


            # Show available dates to help user
            try:
                df = pd.read_excel(uploaded_source, sheet_name="Summary", header=None)
                dates = df.iloc[10:, 0]  # Adjust based on your actual data structure
                st.write(pd.to_datetime(dates, errors='coerce').dropna().dt.strftime("%Y-%m-%d").unique())
            except:
                st.write("Could not extract dates from the uploaded file")
                
        except Exception as e:
            st.error(f"Unexpected error: {str(e)}")
            st.exception(e)