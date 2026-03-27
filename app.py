import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="MCF Admission Auditor", layout="wide")

st.title("📊 MCF Admission Analyzer + Audit System")
st.caption("Created by CA Mahesh Godase")

uploaded_file = st.file_uploader("📂 Upload Admission File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_raw.columns = df_raw.columns.astype(str).str.strip()

        df_raw_display = df_raw.fillna("")
        st.subheader("🔍 Raw Data Preview")
        st.dataframe(df_raw_display)

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

        df_calc = df_raw[[emp_col, camp_col]].dropna()

        df_calc[emp_col] = df_calc[emp_col].astype(str).str.strip().str.upper()
        df_calc[camp_col] = df_calc[camp_col].astype(str).str.strip()

        pivot = pd.pivot_table(df_calc, index=emp_col, columns=camp_col, aggfunc="size", fill_value=0)
        pivot = pivot.reset_index()
        pivot.rename(columns={emp_col: "Employee Name"}, inplace=True)

        pivot["Total"] = pivot.iloc[:, 1:].sum(axis=1)

        total_row = pd.DataFrame(pivot.iloc[:, 1:].sum()).T
        total_row.insert(0, "Employee Name", "Total")

        final_df = pd.concat([pivot, total_row], ignore_index=True)

        st.subheader("📋 Final Report")
        st.dataframe(final_df, use_container_width=True)

        def auto_width(ws):
            for col_cells in ws.iter_cols():
                max_len = max((len(str(c.value)) for c in col_cells if c.value), default=0)
                ws.column_dimensions[get_column_letter(col_cells[0].column)].width = max_len + 3

        def to_excel(df, raw_df):

            wb = Workbook()

            header_fill = PatternFill(start_color="1F4E78", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            bold_font = Font(bold=True)

            thin = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )

            # ===== ADMISSION REPORT =====
            ws = wb.active
            ws.title = "Admission Report"
            ws.append(list(df.columns))
            for row in df.values:
                ws.append(list(row))
            auto_width(ws)

            # ===== FEES ANALYSIS =====
            ws8 = wb.create_sheet("Fees Analysis", 1)

            fee_col = None
            for col in raw_df.columns:
                if "fee" in col.lower() or "amount" in col.lower():
                    fee_col = col

            if fee_col:
                temp_fee = raw_df.copy()
                temp_fee[fee_col] = pd.to_numeric(temp_fee[fee_col], errors="coerce").fillna(0)
                summary = temp_fee.groupby(camp_col)[fee_col].sum().reset_index()

                ws8.append(["Camp", "Fees"])
                for row in summary.values:
                    ws8.append([row[0], row[1]])

                pie = PieChart()
                data = Reference(ws8, min_col=2, min_row=1, max_row=len(summary)+1)
                labels = Reference(ws8, min_col=1, min_row=2, max_row=len(summary)+1)

                pie.add_data(data, titles_from_data=True)
                pie.set_categories(labels)

                pie.width = 12
                pie.height = 10

                ws8.add_chart(pie, "E2")

            auto_width(ws8)

            # ===== CAMP ANALYSIS (IMPROVED) =====
            ws3 = wb.create_sheet("Camp Analysis")

            ws3.append(["Rank", "Camp", "Total", "% Contribution"])

            total = df.iloc[-1]["Total"]

            camp_data = []
            for col in df.columns[1:-1]:
                val = df.iloc[-1][col]
                percent = (val/total) if total else 0
                camp_data.append((col, val, percent))

            # Sort descending
            camp_data = sorted(camp_data, key=lambda x: x[1], reverse=True)

            # Write data with rank
            for i, row in enumerate(camp_data, 1):
                ws3.append([i, row[0], row[1], row[2]])

            # Header styling
            for cell in ws3[1]:
                cell.fill = header_fill
                cell.font = header_font

            # % format
            for row in ws3.iter_rows(min_row=2, min_col=4, max_col=4):
                for cell in row:
                    cell.number_format = "0.00%"

            auto_width(ws3)

            # 📊 BAR CHART
            bar = BarChart()
            data = Reference(ws3, min_col=3, min_row=1, max_row=len(camp_data)+1)
            cats = Reference(ws3, min_col=2, min_row=2, max_row=len(camp_data)+1)

            bar.add_data(data, titles_from_data=True)
            bar.set_categories(cats)
            bar.title = "Camp Performance"

            ws3.add_chart(bar, "F2")

            # 🥧 PIE CHART (MEDIUM)
            pie2 = PieChart()
            data = Reference(ws3, min_col=3, min_row=1, max_row=len(camp_data)+1)
            labels = Reference(ws3, min_col=2, min_row=2, max_row=len(camp_data)+1)

            pie2.add_data(data, titles_from_data=True)
            pie2.set_categories(labels)

            pie2.width = 12
            pie2.height = 10

            ws3.add_chart(pie2, "F20")

            # ===== REST SAME =====
            wb.create_sheet("Dashboard")
            wb.create_sheet("Top Performers")
            wb.create_sheet("Raw Data")

            output = BytesIO()
            wb.save(output)
            return output.getvalue()

        st.download_button(
            "📥 Download Premium MIS Excel",
            data=to_excel(final_df, df_raw),
            file_name="MCF_Premium_MIS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error occurred")
        st.exception(e)

else:
    st.info("👆 Upload file to start")
