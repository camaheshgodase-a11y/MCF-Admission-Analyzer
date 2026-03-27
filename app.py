import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter

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
        df[camp_col] = df[camp_col].astype(str).str.strip()

        # -------- PIVOT --------
        pivot = pd.pivot_table(df, index=emp_col, columns=camp_col, aggfunc="size", fill_value=0)

        pivot = pivot.reset_index()
        pivot.rename(columns={emp_col: "Employee Name"}, inplace=True)

        pivot["Total"] = pivot.iloc[:, 1:].sum(axis=1)

        total_row = pd.DataFrame(pivot.iloc[:, 1:].sum()).T
        total_row.insert(0, "Employee Name", "Total")

        final_df = pd.concat([pivot, total_row], ignore_index=True)

        st.subheader("📋 Final Report")
        st.dataframe(final_df, use_container_width=True)

        # -------- AUTO WIDTH --------
        def auto_width(ws):
            for col_cells in ws.iter_cols():
                length = max(len(str(cell.value)) if cell.value else 0 for cell in col_cells)
                ws.column_dimensions[get_column_letter(col_cells[0].column)].width = length + 3

        # -------- EXCEL FUNCTION --------
        def to_excel(df, raw_df):

            wb = Workbook()

            # ===== MAIN REPORT =====
            ws = wb.active
            ws.title = "Admission Report"

            ws.append(list(df.columns))
            for row in df.values:
                ws.append(list(row))

            auto_width(ws)

            # ===== DASHBOARD =====
            ws2 = wb.create_sheet("Dashboard")

            chart = BarChart()
            chart.title = "Employee-wise Admissions"

            data = Reference(ws, min_col=len(df.columns), min_row=1, max_row=len(df))
            cats = Reference(ws, min_col=1, min_row=2, max_row=len(df))

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)

            ws2.add_chart(chart, "A1")

            # ===== CAMP ANALYSIS =====
            ws3 = wb.create_sheet("Camp Analysis")
            ws3.append(["Camp", "Total", "%"])

            total = df.iloc[-1]["Total"]

            for col in df.columns[1:-1]:
                val = df.iloc[-1][col]
                percent = (val/total) if total else 0
                ws3.append([col, val, percent])

            auto_width(ws3)

            # ===== TOP PERFORMERS =====
            ws4 = wb.create_sheet("Top Performers")

            temp = df.iloc[:-1].sort_values(by="Total", ascending=False)
            ws4.append(["Rank", "Employee", "Total"])

            for i, row in enumerate(temp.values, 1):
                ws4.append([i, row[0], row[-1]])

            auto_width(ws4)

            # ===== AUDIT =====
            ws5 = wb.create_sheet("Audit Summary")
            ws5.append(["Metric", "Value"])
            ws5.append(["Total Records", len(raw_df)])
            ws5.append(["Final Count", df.iloc[-1]["Total"]])
            ws5.append(["Difference", len(raw_df) - df.iloc[-1]["Total"]])

            auto_width(ws5)

            # ===== RAW DATA =====
            ws6 = wb.create_sheet("Raw Data")
            ws6.append(list(raw_df.columns) + ["Status"])

            for row in raw_df.values:
                row_list = list(row)
                status = "Incomplete" if any(pd.isna(x) or str(x).strip()=="" for x in row_list) else "Complete"
                ws6.append(row_list + [status])

            auto_width(ws6)

            # ===== FEES COLLECTION =====
            ws7 = wb.create_sheet("Fees Collection")

            fee_col = None
            for col in raw_df.columns:
                if "fee" in col.lower() or "amount" in col.lower():
                    fee_col = col

            if fee_col:
                temp_fee = raw_df.copy()
                temp_fee[fee_col] = pd.to_numeric(temp_fee[fee_col], errors="coerce").fillna(0)

                summary = temp_fee.groupby(emp_col)[fee_col].sum().reset_index()

                ws7.append(["Employee", "Collected", "Expected", "Balance"])

                EXPECTED = 5000

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

            # ===== FEES ANALYSIS =====
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

        # -------- DOWNLOAD --------
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
