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

        # Keep blank but don't count them
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

        # 🔥 IMPORTANT: drop NaN ONLY for calculation (not display)
        df_calc = df_raw[[emp_col, camp_col]].dropna()

        df_calc[emp_col] = df_calc[emp_col].astype(str).str.strip().str.upper()
        df_calc[camp_col] = df_calc[camp_col].astype(str).str.strip()

        # -------- PIVOT --------
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
            title_font = Font(size=16, bold=True)
            bold_font = Font(bold=True)

            thin = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )

            center = Alignment(horizontal="center", vertical="center")

            # ===== SHEET 1: ADMISSION REPORT =====
            ws = wb.active
            ws.title = "Admission Report"

            ws.append(list(df.columns))
            for row in df.values:
                ws.append(list(row))

            auto_width(ws)

            # ===== SHEET 2: FEES ANALYSIS (MOVED HERE) =====
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
                pie.title = "Camp-wise Fees Distribution"

                data = Reference(ws8, min_col=2, min_row=1, max_row=len(summary)+1)
                labels = Reference(ws8, min_col=1, min_row=2, max_row=len(summary)+1)

                pie.add_data(data, titles_from_data=True)
                pie.set_categories(labels)

                # 🎨 Add colors
                colors = ["FF6384","36A2EB","FFCE56","4BC0C0","9966FF","FF9F40"]
                pie.series[0].data_points = [DataPoint(idx=i) for i in range(len(summary))]
                for i, dp in enumerate(pie.series[0].data_points):
                    dp.graphicalProperties.solidFill = colors[i % len(colors)]

                # 📊 Bigger size
                pie.width = 20
                pie.height = 15

                ws8.add_chart(pie, "E2")

            else:
                ws8.append(["No Fees Column Found"])

            auto_width(ws8)

            # ===== DASHBOARD =====
            ws2 = wb.create_sheet("Dashboard")

            chart = BarChart()
            chart.title = "Employee-wise Admissions"

            data = Reference(ws, min_col=len(df.columns), min_row=1, max_row=len(df))
            cats = Reference(ws, min_col=1, min_row=2, max_row=len(df))

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)

            ws2.add_chart(chart, "A1")

            auto_width(ws2)

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

            # ===== RAW DATA =====
            ws6 = wb.create_sheet("Raw Data")
            ws6.append(list(raw_df.columns))

            for row in raw_df.fillna("").values:
                ws6.append(list(row))

            auto_width(ws6)

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
