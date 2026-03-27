import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

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
            st.error("❌ Required columns missing")
            st.stop()

        # ---------------- CLEAN ----------------
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
            branches = st.sidebar.multiselect("Select Branch", df[branch_col].unique())
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

        # ---------------- REMOVE DUPLICATES (IMPORTANT) ----------------
        df = df.drop_duplicates()

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

        pivot["Total"] = pivot[camp_order].sum(axis=1)

        # ---------------- GRAND TOTAL ----------------
        total_row = pd.DataFrame(pivot[camp_order + ["Total"]].sum()).T
        total_row.insert(0, "Employee Name", "Total")

        final_df = pd.concat([pivot, total_row], ignore_index=True)

        st.subheader("📋 Final Report")
        st.dataframe(final_df, use_container_width=True)

        # ---------------- DASHBOARD ----------------
        st.subheader("📊 Dashboard")

        col1, col2 = st.columns(2)

        with col1:
            st.bar_chart(final_df.set_index("Employee Name")["Total"])

        with col2:
            camp_totals = final_df.iloc[-1, 1:-1]
            st.bar_chart(camp_totals)

        # ---------------- EXCEL STYLING ----------------
        def create_styled_excel(df):
            wb = Workbook()
            ws = wb.active
            ws.title = "Admission Report"

            # Styles
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            bold_font = Font(bold=True)
            center_align = Alignment(horizontal="center", vertical="center")
            border = Border(left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin'))

            # Write header
            for col_num, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=1, column=col_num, value=col_name)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
                cell.border = border

            # Write data
            for row_num, row in enumerate(df.values, 2):
                for col_num, value in enumerate(row, 1):
                    cell = ws.cell(row=row_num, column=col_num, value=value)
                    cell.border = border

                    # Bold total row
                    if row[0] == "Total":
                        cell.font = bold_font

            # Auto width
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[col_letter].width = max_length + 2

            output = BytesIO()
            wb.save(output)
            return output.getvalue()

        excel_file = create_styled_excel(final_df)

        st.download_button(
            "📥 Download Styled Excel",
            data=excel_file,
            file_name="MCF_Admission_Report_Styled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error occurred")
        st.exception(e)

else:
    st.info("👆 Upload file to start")
