import streamlit as st
import pandas as pd

st.title("MCF Admission Analyzer")

uploaded_file = st.file_uploader("Upload Admission Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # SHOW COLUMN NAMES (IMPORTANT)
    st.write("Columns in your file:", df.columns.tolist())

    # SIMPLE AUTO ANALYSIS (NO ERROR)
    st.subheader("Total Admissions")
    st.write(len(df))

    st.subheader("Employee Wise Admissions")
    emp = df.groupby(df.columns[0]).size().reset_index(name="Admissions")
    st.dataframe(emp)

    st.subheader("Camp Wise Admissions")
    camp = df.groupby(df.columns[1]).size().reset_index(name="Admissions")
    st.dataframe(camp)
