import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="MCF Admission Analyzer", layout="wide")

st.title("📊 MCF Summer Camp Admission 2026")

uploaded_file = st.file_uploader("📂 Upload Admission Template", type=["xlsx"])

if uploaded_file is not None:
    try:
        # ---------------- READ FILE ----------------
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.astype(str).str.strip()

        st.subheader("🔍 Raw Data")
        st.dataframe(df)

        # ---------------- COLUMN DETECTION ----------------
        def find_col(keywords):
            for col in df.columns:
                for k in keywords:
                    if k in col.lower():
                        return col
            return None

        name_col = find_col(["name", "employee", "student"])
        camp_col = find_col(["camp", "program", "course"])

        if name_col is None:
            name_col = df.columns[0]
        if camp_col is None:
            camp_col = df.columns[1]

        # ---------------- CLEAN DATA ----------------
        df = df[[name_col, camp_col]].copy()
        df[name_col] = df[name_col].astype(str).str.strip().str.upper()
        df[camp_col] = df[camp_col].astype(str).str.strip()

        # ---------------- FIXED CAMP ORDER ----------------
        camp_order = [
            "MCF SUMMER BOOT CAMP- 45 DAY'S",
            "ADVANCE ADVENTURE CAMP - 10 DAY'S",
            "ADVENTURE TRAINING CAMP - 7 DAY'S",
            "COMMANDO TRANING CAMP -15  DAY'S",
            "SUMMER MILITARY TRAINING CAMP - 30 DAY'S",
            "BASIC ADVENTURE  CAMP - 5 DAY'S",
            "PERSONALITY DEVELOPMENT CAMP - 21 DAY'S",
            "COMMANDO TRANING CAMP -15 DAY'S"
        ]

        # ---------------- CREATE PIVOT (COUNT) ----------------
        pivot = pd.pivot_table(
            df,
            index=name_col,
            columns=camp_col,
            aggfunc='size',
            fill_value=0
        )

        # ---------------- ENSURE ALL COLUMNS EXIST ----------------
        for camp in camp_order:
            if camp not in pivot.columns:
                pivot[camp] = 0

        pivot = pivot[camp_order]  # reorder columns
        pivot = pivot.reset_index()

        pivot.rename(columns={name_col: "Employee Name"}, inplace=True)

        # ---------------- ADD ROW TOTAL ----------------
        pivot["Total"] = pivot[camp_order].sum(axis=1)

        # ---------------- ADD GRAND TOTAL ROW ----------------
        total_row = pd.DataFrame(pivot[camp_order + ["Total"]].sum()).T
        total_row.insert(0, "Employee Name", "Total")

        final_df = pd.concat([pivot, total_row], ignore_index=True)

        # ---------------- DISPLAY ----------------
        st.subheader("📊 Final Output")
        st.dataframe(final_df)

        # ---------------- DOWNLOAD ----------------
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Admission Report')
            return output.getvalue()

        excel_file = to_excel(final_df)

        st.download_button(
            "📥 Download Excel",
            data=excel_file,
            file_name="MCF_Summer_Camp_Admission_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error occurred")
        st.exception(e)

else:
    st.info("👆 Upload your file to generate report")
