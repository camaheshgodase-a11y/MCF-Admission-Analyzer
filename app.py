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

        output_total = int(final_df.iloc[-1]["Total"])
        st.success(f"✅ Output Total: {output_total}")

        # -------- DISPLAY --------
        st.subheader("📋 Final Report")
        st.dataframe(final_df, use_container_width=True)

        # -------- EXCEL FUNCTION --------
        def to_excel(df, raw_df):
            wb = Workbook()

            # ===== SHEET 1 =====
            ws = wb.active
            ws.title = "Admission Report"

            try:
                logo = Image("logo.png")
                ws.add_image(logo, "A1")
            except:
                pass

            ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=len(df.columns))
            ws.cell(row=1, column=2, value="MCF Summer Camp Admission 2026").font = Font(size=16, bold=True)

            header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            border = Border(left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin'))

            start_row = 4

            for col_num, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=start_row, column=col_num, value=col_name)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = Alignment(horizontal="center")

            for r, row in enumerate(df.values, start_row + 1):
                for c, val in enumerate(row, 1):
                    cell = ws.cell(row=r, column=c, value=val)
                    cell.border = border
                    cell.alignment = Alignment(horizontal="center")

            # ===== SHEET 2 DASHBOARD =====
            ws2 = wb.create_sheet("Dashboard")

            chart = BarChart()
            data = Reference(ws, min_col=len(df.columns), min_row=4, max_row=len(df)+3)
            cats = Reference(ws, min_col=1, min_row=5, max_row=len(df)+2)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            ws2.add_chart(chart, "A1")

            # ===== SHEET 3 CAMP ANALYSIS =====
            ws3 = wb.create_sheet("Camp Analysis")

            ws3.append(["Camp", "Total", "%"])

            total = df.iloc[-1]["Total"]

            for col in df.columns[1:-1]:
                val = df.iloc[-1][col]
                ws3.append([col, val, round(val/total*100,2)])

            # ===== SHEET 4 TOP PERFORMERS =====
            ws4 = wb.create_sheet("Top Performers")

            temp = df.iloc[:-1].sort_values(by="Total", ascending=False)

            ws4.append(["Rank", "Employee", "Total"])

            for i, row in enumerate(temp.values, 1):
                ws4.append([i, row[0], row[-1]])

            # ===== SHEET 5 AUDIT =====
            ws5 = wb.create_sheet("Audit")

            ws5.append(["Metric", "Value"])
            ws5.append(["Total Rows", len(raw_df)])
            ws5.append(["Output Total", df.iloc[-1]["Total"]])

            # ===== SHEET 6 BASE DATA =====
            ws6 = wb.create_sheet("Raw Data (Input)")

            ws6.append(list(raw_df.columns))
            for row in raw_df.values:
                ws6.append(list(row))

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
