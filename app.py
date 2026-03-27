import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.utils import get_column_letter

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="MCF Admission Auditor", layout="wide")

st.title("📊 MCF Admission Analyzer + KPI Dashboard")
st.caption("Created by CA Mahesh Godase")

uploaded_file = st.file_uploader("📂 Upload Admission File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_raw.columns = df_raw.columns.astype(str).str.strip()

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

        df[emp_col] = df[emp_col].astype(str).str.strip().str.upper()
        df[camp_col] = df[camp_col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

        # -------- PIVOT --------
        pivot = pd.pivot_table(df, index=emp_col, columns=camp_col, aggfunc="size", fill_value=0)

        pivot = pivot.reset_index()
        pivot.rename(columns={emp_col: "Employee Name"}, inplace=True)

        pivot["Total"] = pivot.iloc[:, 1:].sum(axis=1)

        total_row = pd.DataFrame(pivot.iloc[:, 1:].sum()).T
        total_row.insert(0, "Employee Name", "Total")

        final_df = pd.concat([pivot, total_row], ignore_index=True)

        st.dataframe(final_df)

        # KPI CALC
        total_adm = int(final_df.iloc[-1]["Total"])
        top_emp = final_df.iloc[:-1].sort_values("Total", ascending=False).iloc[0]["Employee Name"]

        # -------- AUTO WIDTH --------
        def auto_width(ws):
            for col_cells in ws.iter_cols():
                col_letter = get_column_letter(col_cells[0].column)
                max_length = max((len(str(c.value)) for c in col_cells if c.value), default=0)
                ws.column_dimensions[col_letter].width = max_length + 3

        # -------- EXCEL --------
        def to_excel(df, raw_df):
            wb = Workbook()

            header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            red_fill = PatternFill(start_color="FFC7CE", fill_type="solid")
            green_fill = PatternFill(start_color="C6EFCE", fill_type="solid")

            center = Alignment(horizontal="center", vertical="center")

            # Sheet 1
            ws = wb.active
            ws.title = "Admission Report"

            for i, col in enumerate(df.columns, 1):
                cell = ws.cell(row=4, column=i, value=col)
                cell.fill = header_fill
                cell.font = header_font

            for r, row in enumerate(df.values, 5):
                ws.append(list(row))

                # Conditional formatting
                total_val = row[-1]
                fill = red_fill if total_val < 20 else green_fill
                ws.cell(row=r, column=len(df.columns)).fill = fill

            auto_width(ws)

            # KPI Sheet
            ws_kpi = wb.create_sheet("KPI Dashboard")

            ws_kpi["A1"] = "KPI Summary"
            ws_kpi["A1"].font = Font(size=14, bold=True)

            ws_kpi.append(["Metric", "Value"])
            ws_kpi.append(["Total Admissions", total_adm])
            ws_kpi.append(["Top Employee", top_emp])

            auto_width(ws_kpi)

            # Dashboard
            ws2 = wb.create_sheet("Dashboard")

            chart = BarChart()
            data = Reference(ws, min_col=len(df.columns), min_row=4, max_row=len(df)+3)
            cats = Reference(ws, min_col=1, min_row=5, max_row=len(df)+2)

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)

            ws2.add_chart(chart, "A1")

            pie = PieChart()
            data2 = Reference(ws, min_col=2, max_col=len(df.columns)-1,
                              min_row=len(df)+3, max_row=len(df)+3)
            cats2 = Reference(ws, min_col=2, max_col=len(df.columns)-1,
                              min_row=4, max_row=4)

            pie.add_data(data2, titles_from_data=True)
            pie.set_categories(cats2)

            ws2.add_chart(pie, "A20")

            auto_width(ws2)

            output = BytesIO()
            wb.save(output)
            return output.getvalue()

        # -------- PDF --------
        def generate_pdf():
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer)
            styles = getSampleStyleSheet()

            story = []
            story.append(Paragraph("MCF Admission Summary", styles["Title"]))
            story.append(Spacer(1, 10))

            story.append(Paragraph(f"Total Admissions: {total_adm}", styles["Normal"]))
            story.append(Paragraph(f"Top Employee: {top_emp}", styles["Normal"]))

            doc.build(story)
            return buffer.getvalue()

        # DOWNLOAD BUTTONS
        st.download_button("📥 Download Excel", data=to_excel(final_df, df_raw), file_name="MIS.xlsx")

        st.download_button("📄 Download PDF Summary", data=generate_pdf(), file_name="Summary.pdf")

    except Exception as e:
        st.error("❌ Error occurred")
        st.exception(e)

else:
    st.info("Upload file to start")
