import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="MCF Admission Auditor", layout="wide")

st.title("📊 MCF Admission Analyzer + Audit System")

uploaded_file = st.file_uploader("📂 Upload Admission File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_raw.columns = df_raw.columns.astype(str).str.strip()

        st.subheader("🔍 Raw Data Preview")
        st.dataframe(df_raw)

        # -------- FIND COLUMNS --------
        def find_col(keys):
            for col in df_raw.columns:
                for k in keys:
                    if k in col.lower():
                        return col
            return None

        emp_col = find_col(["employee", "staff", "counsellor"])
        camp_col = find_col(["camp"])

        if emp_col is None or camp_col is None:
            st.error("❌ Required columns not found")
            st.stop()

        df = df_raw[[emp_col, camp_col]].copy()

        # -------- CLEAN EMPLOYEE --------
        df[emp_col] = (
            df[emp_col]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        # -------- CLEAN CAMP --------
        df[camp_col] = (
            df[camp_col]
            .astype(str)
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)
        )

        # -------- STANDARD CAMP MAP --------
        camp_map = {
            "MCF SUMMER BOOT CAMP- 45 DAY'S": "MCF SUMMER BOOT CAMP- 45 DAY'S",
            "ADVANCE ADVENTURE CAMP - 10 DAY'S": "ADVANCE ADVENTURE CAMP - 10 DAY'S",
            "ADVENTURE TRAINING CAMP - 7 DAY'S": "ADVENTURE TRAINING CAMP - 7 DAY'S",
            "COMMANDO TRANING CAMP -15 DAY'S": "COMMANDO TRANING CAMP -15 DAY'S",
            "COMMANDO TRANING CAMP -15  DAY'S": "COMMANDO TRANING CAMP -15 DAY'S",
            "SUMMER MILITARY TRAINING CAMP - 30 DAY'S": "SUMMER MILITARY TRAINING CAMP - 30 DAY'S",
            "BASIC ADVENTURE CAMP - 5 DAY'S": "BASIC ADVENTURE CAMP - 5 DAY'S",
            "BASIC ADVENTURE  CAMP - 5 DAY'S": "BASIC ADVENTURE CAMP - 5 DAY'S",
            "PERSONALITY DEVELOPMENT CAMP - 21 DAY'S": "PERSONALITY DEVELOPMENT CAMP - 21 DAY'S"
        }

        df["Original Camp"] = df[camp_col]
        df[camp_col] = df[camp_col].replace(camp_map)

        # -------- AUDIT REPORT --------
        st.subheader("🧾 Audit Report")

        # Unknown camps
        unknown_camps = df[~df[camp_col].isin(camp_map.values())]
        st.write("❗ Unknown / Incorrect Camp Names:", unknown_camps.shape[0])
        if not unknown_camps.empty:
            st.dataframe(unknown_camps.head(10))

        # Blank employees
        blank_emp = df[df[emp_col] == ""]
        st.write("❗ Blank Employee Names:", blank_emp.shape[0])

        # Total rows
        st.write(f"📊 Total Input Rows: {len(df)}")

        # -------- FINAL CAMP ORDER --------
        camp_order = list(set(camp_map.values()))

        # -------- PIVOT --------
        pivot = pd.pivot_table(
            df,
            index=emp_col,
            columns=camp_col,
            aggfunc="size",
            fill_value=0
        )

        for camp in camp_order:
            if camp not in pivot.columns:
                pivot[camp] = 0

        pivot = pivot[camp_order]
        pivot = pivot.reset_index()
        pivot.rename(columns={emp_col: "Employee Name"}, inplace=True)

        # -------- TOTAL --------
        pivot["Total"] = pivot[camp_order].sum(axis=1)

        total_row = pd.DataFrame(pivot[camp_order + ["Total"]].sum()).T
        total_row.insert(0, "Employee Name", "Total")

        final_df = pd.concat([pivot, total_row], ignore_index=True)

        # -------- VALIDATION --------
        output_total = int(final_df.iloc[-1]["Total"])
        st.write(f"📊 Output Total Count: {output_total}")

        if output_total != len(df):
            st.error("❌ Mismatch detected between input rows and output count")
        else:
            st.success("✅ Perfect Match: No data loss")

        # -------- DISPLAY --------
        st.subheader("📋 Final Clean Report")
        st.dataframe(final_df, use_container_width=True)

        # -------- DOWNLOAD --------
        def to_excel(df):
            output = BytesIO()
            df.to_excel(output, index=False)
            return output.getvalue()

        st.download_button(
            "📥 Download Final Excel",
            data=to_excel(final_df),
            file_name="MCF_Audit_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error occurred")
        st.exception(e)

else:
    st.info("👆 Upload file to start audit")
