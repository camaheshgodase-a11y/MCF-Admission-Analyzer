import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

st.set_page_config(page_title="MCF Admission Auditor", layout="wide")

st.title("📊 MCF Admission Analyzer + Audit System")
st.caption("Created by CA Mahesh Godase")

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

        # -------- AUTO WIDTH --------
        def auto_width(ws):
            for col_cells in ws.iter_cols():
                col_letter = get_column_letter(col_cells[0].column)
                max_length = max((len(str(c.value)) for c in col_cells if c.value), default=0)
                ws.column_dimensions[col_letter].width = max_length + 3

        # -------- EXCEL FUNCTION --------
        def to_excel(df, raw_df):
            wb = Workbook()

            header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            center = Alignment(horizontal="center", vertical="center")

            thin = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )

            # ===== MAIN REPORT =====
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
                cell.alignment = center
                cell.border = thin

            for r, row in enumerate(df.values, start_row + 1):
                ws.append(list(row))
                for c in range(1, len(df.columns)+1):
                    ws.cell(row=r, column=c).border = thin

            # Highlight total
            for c in range(1, len(df.columns)+1):
                ws.cell(row=ws.max_row, column=c).font = Font(bold=True)

            ws.freeze_panes = "A5"
            auto_width(ws)

            # ===== DASHBOARD =====
            ws2 = wb.create_sheet("Dashboard")

            total_adm = df.iloc[-1]["Total"]
            ws2.append(["Total Admissions", total_adm])

            chart = BarChart()
            data = Reference(ws, min_col=len(df.columns), min_row=4, max_row=ws.max_row)
            cats = Reference(ws, min_col=1, min_row=5, max_row=ws.max_row-1)

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            ws2.add_chart(chart, "A5")

            auto_width(ws2)

            # ===== FEES COLLECTION WITH BALANCE =====
            ws7 = wb.create_sheet("Fees Collection")

            fee_col = None
            for col in raw_df.columns:
                if "fee" in col.lower() or "amount" in col.lower():
                    fee_col = col

            EXPECTED = 5000

            if fee_col:
                temp_fee = raw_df.copy()
                temp_fee[fee_col] = pd.to_numeric(temp_fee[fee_col], errors="coerce").fillna(0)

                summary = temp_fee.groupby(emp_col)[fee_col].sum().reset_index()

                ws7.append(["Employee", "Collected", "Expected", "Balance"])

                for row in summary.values:
                    name = row[0]
                    collected = row[1]

                    admissions = df[df["Employee Name"] == name]["Total"].values
                    admissions = admissions[0] if len(admissions) else 0

                    expected = admissions * EXPECTED
                    balance = expected - collected

                    ws7.append([name, collected, expected, balance])

            auto_width(ws7)

            # ===== FEES ANALYSIS + PIE =====
            ws8 = wb.create_sheet("Fees Analysis")

            if fee_col:
                summary = temp_fee.groupby(camp_col)[fee_col].sum().reset_index()
                total_fee = summary[fee_col].sum()

                ws8.append(["Camp", "Fees"])

                for row in summary.values:
                    ws8.append([row[0], row[1]])

                pie = PieChart()
                data = Reference(ws8, min_col=2, min_row=1, max_row=len(summary)+1)
                labels = Reference(ws8, min_col=1, min_row=2, max_row=len(summary)+1)

                pie.add_data(data, titles_from_data=True)
                pie.set_categories(labels)

                ws8.add_chart(pie, "E2")

            auto_width(ws8)

            # ===== STUDENT AUDIT =====
            ws9 = wb.create_sheet("Student Audit")
            ws9.append(list(raw_df.columns))

            for row in raw_df.values:
                ws9.append(list(row))

            auto_width(ws9)

            output = BytesIO()
            wb.save(output)
            return output.getvalue()

        st.download_button(
            "📥 Download Advanced MIS Excel",
            data=to_excel(final_df, df_raw),
            file_name="MCF_Advanced_MIS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error occurred")
        st.exception(e)

else:
    st.info("👆 Upload file to start")
