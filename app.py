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

            # ===== SHEET 1: MAIN REPORT =====
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

            for r, row in enumerate(df.values, start_row + 1):
                ws.append(list(row))

            auto_width(ws)

            # ===== SHEET 2: DASHBOARD =====
            ws2 = wb.create_sheet("Dashboard")

            chart1 = BarChart()
            chart1.title = "Employee-wise Admissions"

            data = Reference(ws, min_col=len(df.columns), min_row=4, max_row=len(df)+3)
            cats = Reference(ws, min_col=1, min_row=5, max_row=len(df)+2)

            chart1.add_data(data, titles_from_data=True)
            chart1.set_categories(cats)

            ws2.add_chart(chart1, "A1")

            # ===== SHEET 3: CAMP ANALYSIS =====
            ws3 = wb.create_sheet("Camp Analysis")
            ws3.append(["Camp", "Total", "%"])

            total = df.iloc[-1]["Total"]

            for col in df.columns[1:-1]:
                val = df.iloc[-1][col]
                percent = (val/total) if total else 0
                ws3.append([col, val, percent])

            # % format
            for row in ws3.iter_rows(min_row=2, min_col=3, max_col=3):
                for cell in row:
                    cell.number_format = '0.00%'

            auto_width(ws3)

            # ===== SHEET 4: TOP PERFORMERS =====
            ws4 = wb.create_sheet("Top Performers")

            temp = df.iloc[:-1].sort_values(by="Total", ascending=False)
            ws4.append(["Rank", "Employee", "Total"])

            for i, row in enumerate(temp.values, 1):
                ws4.append([i, row[0], row[-1]])

            auto_width(ws4)

            # ===== SHEET 5: AUDIT =====
            ws5 = wb.create_sheet("Audit Summary")
            ws5.append(["Metric", "Value"])
            ws5.append(["Total Records", len(raw_df)])
            ws5.append(["Final Count", df.iloc[-1]["Total"]])
            ws5.append(["Difference", len(raw_df) - df.iloc[-1]["Total"]])

            auto_width(ws5)

            # ===== SHEET 6: RAW DATA =====
            ws6 = wb.create_sheet("Raw Data")
            ws6.append(list(raw_df.columns) + ["Status"])

            for row in raw_df.values:
                row_list = list(row)
                status = "Incomplete" if any(pd.isna(x) or str(x).strip()=="" for x in row_list) else "Complete"
                ws6.append(row_list + [status])

                fill = PatternFill(start_color="FFC7CE" if status=="Incomplete" else "C6EFCE", fill_type="solid")
                for cell in ws6[ws6.max_row]:
                    cell.fill = fill

            auto_width(ws6)

            # ===== SHEET 7: FEES COLLECTION (WITH BALANCE) =====
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

            else:
                ws7.append(["No Fees Column Found"])

            auto_width(ws7)

            # ===== SHEET 8: FEES ANALYSIS + PIE =====
            ws8 = wb.create_sheet("Fees Analysis")

            if fee_col:
                summary = temp_fee.groupby(camp_col)[fee_col].sum().reset_index()

                ws8.append(["Camp", "Fees"])

                for row in summary.values:
                    ws8.append([row[0], row[1]])

                pie = PieChart()
                data = Reference(ws8, min_col=2, min_row=1, max_row=len(summary)+1)
                labels = Reference(ws8, min_col=1, min_row=2, max_row=len(summary)+1)

                pie.add_data(data, titles_from_data=True)
                pie.set_categories(labels)

                ws8.add_chart(pie, "E2")

            else:
                ws8.append(["No Fees Column Found"])

            auto_width(ws8)

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
