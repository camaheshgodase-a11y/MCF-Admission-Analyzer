import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------ PAGE CONFIG ------------------
st.set_page_config(page_title="MCF Admission Analyzer", layout="wide")

st.title("📊 MCF Admission Analyzer")
st.write("Upload your MCF Admission Template Excel file to generate Admission Format.")

# ------------------ FILE UPLOAD ------------------
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        # ------------------ READ FILE ------------------
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.astype(str).str.strip()

        st.subheader("🔍 Raw Data Preview")
        st.dataframe(df)

        # ------------------ VALIDATION ------------------
        if df.empty:
            st.warning("⚠️ Uploaded file is empty")
            st.stop()

        # ------------------ AUTO COLUMN DETECTION ------------------
        def find_column(keywords):
            for col in df.columns:
                for k in keywords:
                    if k in col.lower():
                        return col
            return None

        name_col = find_column(["name", "student"])
        camp_col = find_column(["camp", "days", "duration"])
        fees_col = find_column(["fee", "amount", "paid"])

        # ------------------ FALLBACK (NO CRASH) ------------------
        if name_col is None:
            name_col = df.columns[0]

        if camp_col is None:
            camp_col = df.columns[1]

        if fees_col is None:
            fees_col = df.columns[-1]

        st.write("✅ Detected Columns:")
        st.write(f"Name: {name_col}")
        st.write(f"Camp: {camp_col}")
        st.write(f"Fees: {fees_col}")

        # ------------------ CLEAN DATA ------------------
        df = df[[name_col, camp_col, fees_col]].copy()

        df[fees_col] = pd.to_numeric(df[fees_col], errors='coerce').fillna(0)
        df[name_col] = df[name_col].astype(str).str.strip()
        df[camp_col] = df[camp_col].astype(str).str.strip()

        df = df.dropna(subset=[name_col])

        # ------------------ CREATE ADMISSION FORMAT ------------------
        pivot_df = pd.pivot_table(
            df,
            index=name_col,
            columns=camp_col,
            values=fees_col,
            aggfunc='sum',
            fill_value=0
        )

        pivot_df = pivot_df.reset_index()

        # ------------------ ADD TOTAL COLUMN ------------------
        numeric_cols = pivot_df.select_dtypes(include='number').columns
        pivot_df["Total Fees"] = pivot_df[numeric_cols].sum(axis=1)

        pivot_df.rename(columns={name_col: "Student Name"}, inplace=True)

        st.subheader("📊 Admission Format Output")
        st.dataframe(pivot_df)

        # ------------------ CAMP SUMMARY ------------------
        camp_summary = df.groupby(camp_col)[fees_col].sum().reset_index()
        camp_summary.columns = ["Camp", "Total Fees"]

        st.subheader("🏕️ Camp Summary")
        st.dataframe(camp_summary)

        # ------------------ DOWNLOAD EXCEL ------------------
        def convert_to_excel(df1, df2):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df1.to_excel(writer, index=False, sheet_name='Admission Format')
                df2.to_excel(writer, index=False, sheet_name='Camp Summary')
            return output.getvalue()

        excel_file = convert_to_excel(pivot_df, camp_summary)

        st.download_button(
            label="📥 Download Admission Format Excel",
            data=excel_file,
            file_name="Admission_Format_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error occurred while processing file")
        st.exception(e)

else:
    st.info("👆 Please upload an Excel file to start analysis.")
