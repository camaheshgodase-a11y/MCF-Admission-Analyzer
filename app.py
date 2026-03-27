# ADD THESE IMPORTS AT TOP
from openpyxl.chart import PieChart
from openpyxl.formatting.rule import CellIsRule

# -------- EXCEL FUNCTION --------
def to_excel(df, raw_df):
    wb = Workbook()

    header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    bold_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    # ===== SHEET 1: MAIN REPORT =====
    ws = wb.active
    ws.title = "Admission Report"

    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=len(df.columns))
    ws.cell(row=1, column=1, value="MCF Summer Camp Admission 2026").font = Font(size=16, bold=True)

    start_row = 4

    # Header
    for col_num, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=start_row, column=col_num, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin_border

    # Data
    for r, row in enumerate(df.values, start_row + 1):
        ws.append(list(row))
        for c in range(1, len(df.columns)+1):
            ws.cell(row=r, column=c).border = thin_border

    # Conditional Formatting (Low performers <10)
    ws.conditional_formatting.add(
        f"{get_column_letter(len(df.columns))}{start_row+1}:{get_column_letter(len(df.columns))}{start_row+len(df)}",
        CellIsRule(operator='lessThan', formula=['10'], fill=PatternFill(start_color="FFC7CE", fill_type="solid"))
    )

    auto_width(ws)

    # ===== SHEET 2: DASHBOARD =====
    ws2 = wb.create_sheet("Dashboard")

    total_adm = df.iloc[-1]["Total"]
    top_emp = df.iloc[:-1].sort_values(by="Total", ascending=False).iloc[0]["Employee Name"]

    # KPI
    ws2["A1"] = "KPI Dashboard"
    ws2["A1"].font = Font(size=14, bold=True)

    ws2.append(["Metric", "Value"])
    ws2.append(["Total Admissions", total_adm])
    ws2.append(["Top Performer", top_emp])

    # Chart
    chart1 = BarChart()
    chart1.title = "Employee-wise Admissions"

    data = Reference(ws, min_col=len(df.columns), min_row=4, max_row=len(df)+3)
    cats = Reference(ws, min_col=1, min_row=5, max_row=len(df)+2)

    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)

    ws2.add_chart(chart1, "E2")

    auto_width(ws2)

    # ===== SHEET 3: CAMP ANALYSIS =====
    ws3 = wb.create_sheet("Camp Analysis")
    ws3.append(["Camp", "Total", "% Contribution"])

    total = df.iloc[-1]["Total"]

    for col in df.columns[1:-1]:
        val = df.iloc[-1][col]
        ws3.append([col, val, round(val/total*100, 2)])

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
    ws5.append(["Mismatch", len(raw_df) - df.iloc[-1]["Total"]])

    auto_width(ws5)

    # ===== SHEET 6: FEES COLLECTION =====
    ws7 = wb.create_sheet("Fees Collection")

    fee_col = None
    for col in raw_df.columns:
        if "fee" in col.lower() or "amount" in col.lower():
            fee_col = col

    if fee_col:
        temp_fee = raw_df.copy()
        temp_fee[fee_col] = pd.to_numeric(temp_fee[fee_col], errors="coerce").fillna(0)

        summary = temp_fee.groupby(emp_col)[fee_col].sum().reset_index()

        ws7.append(["Employee", "Collected Fees", "Expected Fees", "Balance"])

        EXPECTED_PER_ADMISSION = 5000  # 🔥 You can change

        for row in summary.values:
            collected = row[1]
            admissions = df[df["Employee Name"] == row[0]]["Total"].values
            admissions = admissions[0] if len(admissions) > 0 else 0

            expected = admissions * EXPECTED_PER_ADMISSION
            balance = expected - collected

            ws7.append([row[0], collected, expected, balance])

    else:
        ws7.append(["No Fees Column Found"])

    auto_width(ws7)

    # ===== SHEET 7: FEES ANALYSIS =====
    ws8 = wb.create_sheet("Fees Analysis")

    if fee_col:
        summary = temp_fee.groupby(camp_col)[fee_col].sum().reset_index()
        total_fee = summary[fee_col].sum()

        ws8.append(["Camp", "Fees", "%"])

        for row in summary.values:
            percent = (row[1]/total_fee*100) if total_fee else 0
            ws8.append([row[0], row[1], round(percent, 2)])

        # PIE CHART
        pie = PieChart()
        pie.title = "Fees Distribution"

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
