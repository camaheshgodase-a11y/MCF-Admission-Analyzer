import streamlit as st
import pandas as pd

st.title("MCF Admission MIS Analyzer")

uploaded_file = st.file_uploader("Upload Admission Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Remove blank rows
    df = df.dropna(how='all')

    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

    st.subheader("Select Columns")
    employee_col = st.selectbox("Employee Name Column", df.columns)
    camp_col = st.selectbox("Camp Name Column", df.columns)
    date_col = st.selectbox("Admission Date Column", df.columns)
    fees_col = st.selectbox("Fees Column", df.columns)
    balance_col = st.selectbox("Balance Column", df.columns)

    # Clean text
    df[employee_col] = df[employee_col].astype(str).str.strip()
    df[camp_col] = df[camp_col].astype(str).str.strip()

    # Convert Date
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df[df[date_col].notna()]

    # Convert numeric
    df[fees_col] = pd.to_numeric(df[fees_col], errors='coerce').fillna(0)
    df[balance_col] = pd.to_numeric(df[balance_col], errors='coerce').fillna(0)

    # Month Column
    df['Month'] = df[date_col].apply(lambda x: x.strftime('%Y-%m'))

    # Pivot Table Camp vs Employee
    pivot_table = pd.pivot_table(
        df,
        index=camp_col,
        columns=employee_col,
        aggfunc='size',
        fill_value=0
    )

    st.subheader("Camp vs Employee Admission Count")
    st.dataframe(pivot_table)

    # Summaries
    employee_summary = df.groupby(employee_col).size().reset_index()
    employee_summary.columns = [employee_col, "Admissions"]

    camp_summary = df.groupby(camp_col).size().reset_index()
    camp_summary.columns = [camp_col, "Admissions"]

    date_summary = df.groupby(date_col).size().reset_index()
    date_summary.columns = ["Date", "Admissions"]

    month_summary = df.groupby('Month').size().reset_index()
    month_summary.columns = ["Month", "Admissions"]

    fees_summary = df.groupby(employee_col, as_index=False)[fees_col].sum()
    balance_summary = df.groupby(employee_col, as_index=False)[balance_col].sum()

    # Save Excel
    output_file = "MCF_Admission_MIS_Report.xlsx"

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        pivot_table.to_excel(writer, sheet_name='Camp vs Employee')
        employee_summary.to_excel(writer, sheet_name='Employee Summary', index=False)
        camp_summary.to_excel(writer, sheet_name='Camp Summary', index=False)
        date_summary.to_excel(writer, sheet_name='Date Summary', index=False)
        month_summary.to_excel(writer, sheet_name='Monthly Summary', index=False)
        fees_summary.to_excel(writer, sheet_name='Fees Summary', index=False)
        balance_summary.to_excel(writer, sheet_name='Balance Summary', index=False)

    with open(output_file, "rb") as file:
        st.download_button(
            label="Download Full MIS Excel Report",
            data=file,
            file_name="MCF_Admission_MIS_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
