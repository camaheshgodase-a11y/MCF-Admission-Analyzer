import streamlit as st
import pandas as pd

st.title("MCF Admission Analyzer")

uploaded_file = st.file_uploader("Upload Admission Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.write("Columns in your file:", df.columns.tolist())

    # Select columns
    employee_col = st.selectbox("Select Employee Column", df.columns)
    camp_col = st.selectbox("Select Camp Column", df.columns)

    # Clean Data (IMPORTANT)
    df[employee_col] = df[employee_col].astype(str)
    df[camp_col] = df[camp_col].astype(str)

    df = df.dropna(subset=[employee_col, camp_col])

    # Pivot Table
    pivot_table = pd.pivot_table(
        df,
        index=camp_col,
        columns=employee_col,
        aggfunc='size',
        fill_value=0
    )

    st.subheader("Camp vs Employee Admission Count")
    st.dataframe(pivot_table)

    # Download Excel
    output_file = "Admission_Pivot.xlsx"
    pivot_table.to_excel(output_file)

    with open(output_file, "rb") as file:
        st.download_button(
            label="Download Pivot Excel",
            data=file,
            file_name="Admission_Pivot.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
