def to_excel(df, raw_df):

    wb = Workbook()

    header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    bold_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")

    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    def style_sheet(ws):
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = center

    def auto_width(ws):
        for col_cells in ws.iter_cols():
            col_letter = get_column_letter(col_cells[0].column)
            max_length = max((len(str(c.value)) for c in col_cells if c.value), default=0)
            ws.column_dimensions[col_letter].width = max_length + 3

    # ================= SHEET 1 =================
    ws = wb.active
    ws.title = "Admission Report"

    ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=len(df.columns))
    ws.cell(row=1, column=2, value="MCF Summer Camp Admission 2026").font = Font(size=16, bold=True)

    for i, col in enumerate(df.columns, 1):
        cell = ws.cell(row=4, column=i, value=col)
        cell.fill = header_fill
        cell.font = header_font

    for r, row in enumerate(df.values, 5):
        ws.append(list(row))

    style_sheet(ws)
    auto_width(ws)

    # ================= DASHBOARD =================
    ws2 = wb.create_sheet("Dashboard")

    chart1 = BarChart()
    chart1.title = "Employee-wise Admissions"
    data = Reference(ws, min_col=len(df.columns), min_row=4, max_row=len(df)+3)
    cats = Reference(ws, min_col=1, min_row=5, max_row=len(df)+2)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    ws2.add_chart(chart1, "A1")

    # Pie Chart
    from openpyxl.chart import PieChart
    pie = PieChart()
    pie.title = "Camp Contribution"

    data2 = Reference(ws, min_col=2, max_col=len(df.columns)-1,
                      min_row=len(df)+3, max_row=len(df)+3)
    cats2 = Reference(ws, min_col=2, max_col=len(df.columns)-1,
                      min_row=4, max_row=4)

    pie.add_data(data2, titles_from_data=True)
    pie.set_categories(cats2)

    ws2.add_chart(pie, "A20")

    # ================= CAMP ANALYSIS =================
    ws3 = wb.create_sheet("Camp Analysis")

    ws3.append(["Camp", "Total", "% Contribution"])

    total = df.iloc[-1]["Total"]

    for col in df.columns[1:-1]:
        val = df.iloc[-1][col]
        ws3.append([col, val, round(val/total*100, 2)])

    for cell in ws3[1]:
        cell.fill = header_fill
        cell.font = header_font

    style_sheet(ws3)
    auto_width(ws3)

    # ================= TOP PERFORMERS =================
    ws4 = wb.create_sheet("Top Performers")

    ws4.append(["Rank", "Employee", "Total"])

    temp = df.iloc[:-1].sort_values(by="Total", ascending=False)

    for i, row in enumerate(temp.values, 1):
        ws4.append([i, row[0], row[-1]])

    for cell in ws4[1]:
        cell.fill = header_fill
        cell.font = header_font

    style_sheet(ws4)
    auto_width(ws4)

    # ================= AUDIT =================
    ws5 = wb.create_sheet("Audit Summary")

    ws5.append(["Metric", "Value"])
    ws5.append(["Total Records", len(raw_df)])
    ws5.append(["Final Count", df.iloc[-1]["Total"]])

    for cell in ws5[1]:
        cell.fill = header_fill
        cell.font = header_font

    style_sheet(ws5)
    auto_width(ws5)

    # ================= RAW DATA =================
    ws6 = wb.create_sheet("Raw Data")

    ws6.append(list(raw_df.columns) + ["Status"])

    for row in raw_df.values:
        row_list = list(row)
        status = "Incomplete" if any(pd.isna(x) or str(x).strip()=="" for x in row_list) else "Complete"
        ws6.append(row_list + [status])

    auto_width(ws6)

    # ================= FEES COLLECTION =================
    ws7 = wb.create_sheet("Fees Collection")

    fee_col = next((c for c in raw_df.columns if "fee" in c.lower() or "amount" in c.lower()), None)
    expected_col = next((c for c in raw_df.columns if "total" in c.lower()), None)

    if fee_col:
        temp = raw_df.copy()
        temp[fee_col] = pd.to_numeric(temp[fee_col], errors="coerce").fillna(0)

        if expected_col:
            temp[expected_col] = pd.to_numeric(temp[expected_col], errors="coerce").fillna(0)
        else:
            temp["Expected"] = temp[fee_col]
            expected_col = "Expected"

        summary = temp.groupby(emp_col)[[fee_col, expected_col]].sum().reset_index()

        ws7.append(["Employee", "Collected Fees", "Total Fees", "Balance Fees"])

        for row in summary.values:
            balance = row[2] - row[1]
            ws7.append([row[0], row[1], row[2], balance])

    else:
        ws7.append(["No Fees Column Found"])

    for cell in ws7[1]:
        cell.fill = header_fill
        cell.font = header_font

    style_sheet(ws7)
    auto_width(ws7)

    # ================= FEES ANALYSIS =================
    ws8 = wb.create_sheet("Fees Analysis")

    if fee_col:
        summary = temp.groupby(camp_col)[fee_col].sum().reset_index()
        total_fee = summary[fee_col].sum()

        ws8.append(["Camp", "Fees", "%"])

        for row in summary.values:
            percent = (row[1]/total_fee*100) if total_fee else 0
            ws8.append([row[0], row[1], round(percent, 2)])

    else:
        ws8.append(["No Fees Column Found"])

    for cell in ws8[1]:
        cell.fill = header_fill
        cell.font = header_font

    style_sheet(ws8)
    auto_width(ws8)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()
