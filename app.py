import streamlit as st
import pandas as pd

st.set_page_config(page_title="MCF Admission Analyzer", layout="wide")

st.title("📊 MCF Admission Analyzer")
st.write("Upload your Admission Template Excel file to generate analysis.")

# File Upload
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read Excel
        df = pd.read_excel(uploaded_file, header=1)

        st.subheader("🔍 Raw Data Preview")
        st.dataframe(df)

        # ---- CLEAN DATA ----
        df = df.dropna(how='all')

        # Rename columns (adjust based on your file)
        df.columns = df.columns.str.strip()

        # Example important columns (EDIT if needed)
        # You must match these with your actual Excel columns
        name_col = "Employee Name"
        camp_col = "Camp Type"
        fees_col = "Fees"

        # If columns not matching, auto detect (basic fallback)
        if name_col not in df.columns:
            name_col = df.columns[2]
        if camp_col not in df.columns:
            camp_col = df.columns[5]
        if fees_col not in df.columns:
            fees_col = df.columns[-1]

        # Convert fees to numeric
        df[fees_col] = pd.to_numeric(df[fees_col], errors='coerce')

        # ---- ANALYSIS ----
        summary = df.groupby(name_col)[fees_col].sum().reset_index()

        summary.columns = ["Employee Name", "Total Fees"]

        st.subheader("📊 Admission Summary")
        st.dataframe(summary)

        # ---- CAMP WISE ANALYSIS ----
        camp_summary = df.groupby(camp_col)[fees_col].sum().reset_index()
        camp_summary.columns = ["Camp Type", "Total Fees"]

        st.subheader("🏕️ Camp-wise Summary")
        st.dataframe(camp_summary)

        # ---- DOWNLOAD ----
        def convert_to_excel(df1, df2):
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df1.to_excel(writer, index=False, sheet_name='Employee Summary')
                df2.to_excel(writer, index=False, sheet_name='Camp Summary')
            return output.getvalue()

        excel_data = convert_to_excel(summary, camp_summary)

        st.download_button(
            label="📥 Download Report",
            data=excel_data,
            file_name="MCF_Admission_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
