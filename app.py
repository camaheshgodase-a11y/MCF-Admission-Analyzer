import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference

st.set_page_config(page_title="MCF Admission Dashboard", layout="wide")

st.title("📊 MCF Summer Camp Admission 2026")

uploaded_file = st.file_uploader("📂 Upload Admission File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.astype(str).str.strip()

        # ---------------- COLUMN DETECTION ----------------
        def find_col(keys):
            for col in df.columns:
                for k in keys:
                    if k in col.lower():
                        return col
            return None

        emp_col = find_col(["employee", "staff", "counsellor"])
        camp_col = find_col(["camp"])
        date_col = find_col(["date"])
        branch_col = find_col(["branch"])

        if emp_col is None or camp_col is None:
            st.error("❌ Employee or Camp column not found")
            st.stop()

        # ---------------- CLEAN DATA ----------------
        df = df[[emp_col, camp_col] + ([date_col] if date_col else []) + ([branch_col] if branch_col else [])]

        df[emp_col] = df[emp_col].astype(str).str.strip().str.upper()
        df[camp_col] = df[camp_col].astype(str).str.strip()

        # ---------------- FILTERS ----------------
        st.sidebar.header("🔍 Filters")

        if date_col:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            min_date = df[date_col].min()
            max_date = df[date_col].max()

            selected_date = st.sidebar.date_input("Select Date Range", [min_date, max_date])

            if len(selected_date) == 2:
                df = df[(df[date_col] >= pd.to_datetime(selected_date[0])) &
                        (df[date_col] <= pd.to_datetime(selected_date[1]))]

        if branch_col:
            branches = st.sidebar.multiselect("Select Branch", df[branch_col].dropna().unique())
            if branches:
                df = df[df[branch_col].isin(branches)]

        # ---------------- CAMP ORDER ----------------
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

        # ---------------- PIVOT ----------------
        pivot = pd.pivot_table(
            df,
            index=emp_col,
            columns=camp_col,
            aggfunc='size',
            fill_value=0
        )

        for camp in camp_order:
            if camp not in pivot.columns:
                pivot[camp] = 0

        pivot = pivot[camp_order]
        pivot = pivot.reset_index()
        pivot.rename(columns={emp_col: "Employee Name"}, inplace=True)

        # ---------------- TOTAL ----------------
        pivot["Total"] = pivot[camp_order].sum(axis=1)

        total_row = pd.DataFrame(pivot[camp_order + ["Total"]].sum()).T
        total_row.insert(0, "Employee Name", "Total")

        final_df = pd.concat([pivot, total_row], ignore_index=True)

        # ---------------- VALIDATION ----------------
        grand_total = int(final_df.iloc[-1]["Total"])
        st.success(f"✅ Total Admissions Count: {grand_total}")

        # ---------------- DISPLAY ----------------
        st.subheader("📋 Final Report")
        st.dataframe(final_df, use_container_width=True)

        # ---------------- DASHBOARD ----------------
        st.subheader("📊 Dashboard")

        col1, col2 = st.columns(2)

        with col1:
            st.bar_chart(final_df.set_index("Employee Name").iloc[:-1]["Total"])

        with col2:
            st.bar_chart(final_df.iloc[-1][1:-1])

        # ---------------- EXCEL EXPORT ----------------
        def create_excel_with_dashboard(df):
            wb = Workbook()
            ws = wb.active
            ws.title = "Admission Report"

            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            bold_font = Font(bold=True)
            center = Alignment(horizontal="center")

            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

            # Header
            for col_num, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=1, column=col_num, value=col_name)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center
                cell.border = border

            # Data
            for r, row in enumerate(df.values, 2):
                for c, val in enumerate(row, 1):
                    cell = ws.cell(row=r, column=c, value=val)
                    cell.border = border
                    if row[0] == "Total":
                        cell.font = bold_font

            # ---------------- DASHBOARD SHEET ----------------
            ws2 = wb.create_sheet(title="Dashboard")

            chart1 = BarChart()
            chart1.title = "Employee-wise Admissions"
            data = Reference(ws, min_col=len(df.columns), min_row=1, max_row=len(df))
            cats = Reference(ws, min_col=1, min_row=2, max_row=len(df)-1)
            chart1.add_data(data, titles_from_data=True)
            chart1.set_categories(cats)
            ws2.add_chart(chart1, "A1")

            chart2 = BarChart()
            chart2.title = "Camp-wise Admissions"
            data2 = Reference(ws, min_col=2, max_col=len(df.columns)-1,
                              min_row=len(df), max_row=len(df))
            cats2 = Reference(ws, min_col=2, max_col=len(df.columns)-1,
                              min_row=1, max_row=1)
            chart2.add_data(data2, titles_from_data=True)
            chart2.set_categories(cats2)
            ws2.add_chart(chart2, "A20")

            output = BytesIO()
            wb.save(output)
            return output.getvalue()

        excel_file = create_excel_with_dashboard(final_df)

        st.download_button(
            "📥 Download Final Excel with Dashboard",
            data=excel_file,
            file_name="MCF_Admission_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error occurred")
        st.exception(e)

else:
    st.info("👆 Upload file to start")
