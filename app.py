import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="MCF Admission Auditor", layout="wide")

st.title("📊 MCF Admission Analyzer + Audit System")
st.caption("Created by CA Mahesh Godase")

uploaded_file = st.file_uploader("📂 Upload Admission File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_raw.columns = df_raw.columns.astype(str).str.strip()

        # ✅ FIX 1: Remove NaN from raw data
        df_raw = df_raw.fillna("")

        st.subheader("🔍 Raw Data Preview")
        st.dataframe(df_raw)

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

        df[emp_col] = df[emp_col].astype(str).str.strip().str.upper()
        df[camp_col] = df[camp_col].astype(str).str.strip()

        pivot = pd.pivot_table(df, index=emp_col, columns=camp_col, aggfunc="size", fill_value=0)
        pivot = pivot.reset_index()
        pivot.rename(columns={emp_col: "Employee Name"}, inplace=True)

        pivot["Total"] = pivot.iloc[:, 1:].sum(axis=1)

        total_row = pd.DataFrame(pivot.iloc[:, 1:].sum()).T
        total_row.insert(0, "Employee Name", "Total")

        final_df = pd.concat([pivot, total_row], ignore_index=True)

        # ✅ FIX 2: Remove NaN from final report
        final_df = final_df.fillna("")

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

            # ===== MAIN REPORT =====
            ws = wb.active
            ws.title = "Admission Report"

            ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=len(df.columns))
            ws.cell(1, 1).value = "MCF Admission MIS Report"
            ws.cell(1, 1).font = title_font
            ws.cell(1, 1).alignment = center

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
                    ws.cell(row=r, column=c).alignment = center

            for c in range(1, len(df.columns)+1):
                cell = ws.cell(row=ws.max_row, column=c)
                cell.font = bold_font
                cell.fill = PatternFill(start_color="D9E1F2", fill_type="solid")

            ws.freeze_panes = "A5"
            auto_width(ws)

            # ===== DASHBOARD =====
            ws2 = wb.create_sheet("Dashboard")

            total_adm = df.iloc[-1]["Total"]
            top_emp = df.iloc[:-1].sort_values(by="Total", ascending=False).iloc[0]["Employee Name"]

            ws2["A1"] = "KPI Dashboard"
            ws2["A1"].font = title_font

            ws2.append(["Metric", "Value"])
            ws2.append(["Total Admissions", total_adm])
            ws2.append(["Top Performer", top_emp])

            chart = BarChart()
            chart.title = "Employee-wise Admissions"

            data = Reference(ws, min_col=len(df.columns), min_row=4, max_row=ws.max_row)
            cats = Reference(ws, min_col=1, min_row=5, max_row=ws.max_row-1)

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)

            ws2.add_chart(chart, "E2")

            auto_width(ws2)

            # ===== CAMP ANALYSIS =====
            ws3 = wb.create_sheet("Camp Analysis")
            ws3.append(["Camp", "Total", "% Contribution"])

            total = df.iloc[-1]["Total"]

            for col in df.columns[1:-1]:
                val = df.iloc[-1][col]
                percent = (val/total) if total else 0
                ws3.append([col, val, percent])

            for row in ws3.iter_rows(min_row=2, min_col=3, max_col=3):
                for cell in row:
                    cell.number_format = "0.00%"

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
                status = "Incomplete" if any(str(x).strip()=="" for x in row_list) else "Complete"
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
