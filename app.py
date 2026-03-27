def to_excel(df, raw_df):
    wb = Workbook()

    header_fill = PatternFill(start_color="305496", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    bold_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")

    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    # ===== SHEET 1: MAIN REPORT =====
    ws = wb.active
    ws.title = "Admission Report"

    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=len(df.columns))
    title = ws.cell(row=1, column=1, value="MCF Summer Camp Admission 2026")
    title.font = Font(size=16, bold=True)
    title.alignment = center

    start_row = 4

    # Header
    for col_num, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=start_row, column=col_num, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin

    # Data
    for r, row in enumerate(df.values, start_row + 1):
        ws.append(list(row))
        for c in range(1, len(df.columns)+1):
            ws.cell(row=r, column=c).border = thin

    # Highlight TOTAL row
    last_row = ws.max_row
    for c in range(1, len(df.columns)+1):
        cell = ws.cell(row=last_row, column=c)
        cell.font = bold_font
        cell.fill = PatternFill(start_color="D9E1F2", fill_type="solid")

    # Freeze header
    ws.freeze_panes = "A5"

    auto_width(ws)

    # ===== SHEET 2: DASHBOARD =====
    ws2 = wb.create_sheet("Dashboard")

    total_adm = df.iloc[-1]["Total"]
    top_emp = df.iloc[:-1].sort_values(by="Total", ascending=False).iloc[0]["Employee Name"]

    ws2["A1"] = "KPI Dashboard"
    ws2["A1"].font = Font(size=14, bold=True)

    ws2.append(["Metric", "Value"])
    ws2.append(["Total Admissions", total_adm])
    ws2.append(["Top Performer", top_emp])

    for row in ws2.iter_rows(min_row=2, max_row=4, min_col=1, max_col=2):
        for cell in row:
            cell.border = thin

    # Chart
    chart1 = BarChart()
    chart1.title = "Employee-wise Admissions"

    data = Reference(ws, min_col=len(df.columns), min_row=4, max_row=ws.max_row)
    cats = Reference(ws, min_col=1, min_row=5, max_row=ws.max_row-1)

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
        percent = round(val/total*100, 2) if total else 0
        ws3.append([col, val, percent])

    for row in ws3.iter_rows():
        for cell in row:
            cell.border = thin

    auto_width(ws3)

    # ===== SHEET 4: TOP PERFORMERS =====
    ws4 = wb.create_sheet("Top Performers")

    temp = df.iloc[:-1].sort_values(by="Total", ascending=False)
    ws4.append(["Rank", "Employee", "Total"])

    for i, row in enumerate(temp.values, 1):
        ws4.append([i, row[0], row[-1]])

    for row in ws4.iter_rows():
        for cell in row:
            cell.border = thin

    auto_width(ws4)

    # ===== SHEET 5: AUDIT =====
    ws5 = wb.create_sheet("Audit Summary")

    ws5.append(["Metric", "Value"])
    ws5.append(["Total Records", len(raw_df)])
    ws5.append(["Final Count", df.iloc[-1]["Total"]])
    ws5.append(["Difference", len(raw_df) - df.iloc[-1]["Total"]])

    for row in ws5.iter_rows():
        for cell in row:
            cell.border = thin

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
            cell.border = thin

    auto_width(ws6)

    # ===== SHEET 7: FEES COLLECTION =====
    ws7 = wb.create_sheet("Fees Collection")

    fee_col = None
    for col in raw_df.columns:
        if "fee" in col.lower() or "amount" in col.lower():
            fee_col = col

    if fee_col:
        temp_fee = raw_df.copy()
        temp_fee[fee_col] = pd.to_numeric(temp_fee[fee_col], errors="coerce").fillna(0)

        summary = temp_fee.groupby(emp_col)[fee_col].sum().reset_index()

        ws7.append(["Employee", "Total Fees"])

        for row in summary.values:
            ws7.append(list(row))

    else:
        ws7.append(["No Fees Column Found"])

    for row in ws7.iter_rows():
        for cell in row:
            cell.border = thin

    auto_width(ws7)

    # ===== SHEET 8: FEES ANALYSIS =====
    ws8 = wb.create_sheet("Fees Analysis")

    if fee_col:
        summary = temp_fee.groupby(camp_col)[fee_col].sum().reset_index()
        total_fee = summary[fee_col].sum()

        ws8.append(["Camp", "Fees", "%"])

        for row in summary.values:
            percent = (row[1]/total_fee*100) if total_fee else 0
            ws8.append([row[0], row[1], round(percent, 2)])

    else:
        ws8.append(["No Fees Column Found"])

    for row in ws8.iter_rows():
        for cell in row:
            cell.border = thin

    auto_width(ws8)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()
