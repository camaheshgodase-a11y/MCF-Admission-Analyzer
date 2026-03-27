import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image

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

        # -------- CLEAN DATA --------
        df[emp_col] = df[emp_col].astype(str).str.strip().str.upper()
        df[camp_col] = df[camp_col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

        # -------- CAMP STANDARDIZATION --------
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

        df[camp_col] = df[camp_col].replace(camp_map)

        # -------- AUDIT --------
        st.subheader("🧾 Audit Report")

        unknown = df[~df[camp_col].isin(camp_map.values())]
        st.write("❗ Unknown Camps:", len(unknown))

        blank_emp = df[df[emp_col] == ""]
        st.write("❗ Blank Employees:", len(blank_emp))

        st.write("📊 Input Rows:", len(df))

        # -------- PIVOT --------
        camp_order = list(set(camp_map.values()))

        pivot = pd.pivot_table(df, index=emp_col, columns=camp_col, aggfunc="size", fill_value=0)

        for c in camp_order:
            if c not in pivot.columns:
                pivot[c] = 0

        pivot = pivot[camp_order].reset_index()
        pivot.rename(columns={emp_col: "Employee Name"}, inplace=True)

        pivot["Total"] = pivot[camp_order].sum(axis=1)

        total_row = pd.DataFrame(pivot[camp_order + ["Total"]].sum()).T
        total_row.insert(0, "Employee Name", "Total")

        final_df = pd.concat([pivot, total_row], ignore_index=True)

        st.subheader("📋 Final Report")
        st.dataframe(final_df, use_container_width=True)

        # -------- EXCEL FUNCTION --------
        def to_excel(df, raw_df):
            wb = Workbook()

            header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            bold_font = Font(bold=True)
            center = Alignment(horizontal="center", vertical="center")

            border = Border(left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin'))

            def auto_width(ws):
                for col in ws.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    ws.column_dimensions[col_letter].width = max_len + 3

            # ===== SHEET 1 =====
            ws = wb.active
            ws.title = "Admission Report"

            try:
                ws.add_image(Image("logo.png"), "A1")
            except:
                pass

            ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=len(df.columns))
            ws.cell(row=1, column=2, value="MCF Summer Camp Admission 2026").font = Font(size=16, bold=True)

            start_row = 4

            for col_num, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=start_row, column=col_num, value=col_name)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = center

            for r, row in enumerate(df.values, start_row + 1):
                for c, val in enumerate(row, 1):
                    cell = ws.cell(row=r, column=c, value=val)
                    cell.border = border
                    cell.alignment = center

                    if str(row[0]).upper() == "TOTAL":
                        cell.font = bold_font
                        cell.fill = PatternFill(start_color="D9D9D9", fill_type="solid")

            ws.freeze_panes = "A5"
            auto_width(ws)

            # ===== SHEET 6 RAW DATA =====
            ws6 = wb.create_sheet("Raw Data (Input)")

            headers = list(raw_df.columns) + ["Data Status"]
            ws6.append(headers)

            for row in raw_df.values:
                row_list = list(row)
                status = "Incomplete" if any(pd.isna(x) or str(x).strip() == "" for x in row_list) else "Complete"
                row_list.append(status)
                ws6.append(row_list)

                fill = PatternFill(start_color="FFC7CE" if status=="Incomplete" else "C6EFCE", fill_type="solid")

                for cell in ws6[ws6.max_row]:
                    cell.fill = fill

            auto_width(ws6)

            # ===== SAVE =====
            output = BytesIO()
            wb.save(output)
            return output.getvalue()

        # -------- DOWNLOAD --------
        st.download_button(
            "📥 Download Final MIS Excel",
            data=to_excel(final_df, df_raw),
            file_name="MCF_Final_MIS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error occurred")
        st.exception(e)

else:
    st.info("👆 Upload file to start")
