from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image
from io import BytesIO

def create_formatted_excel_with_logo(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Admission Report"

    # -------- ADD LOGO --------
    try:
        logo = Image("logo.png")  # keep logo in same folder
        logo.height = 60
        logo.width = 120
        ws.add_image(logo, "A1")
    except:
        pass  # if logo missing, skip

    # -------- ADD TITLE --------
    title = "MCF Summer Camp Admission 2026"

    ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=len(df.columns))
    title_cell = ws.cell(row=1, column=2, value=title)

    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # -------- STYLES --------
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    start_row = 4  # data starts from row 4

    # -------- HEADER --------
    for col_num, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=start_row, column=col_num, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border

    # -------- DATA --------
    for row_num, row in enumerate(df.values, start_row + 1):
        for col_num, value in enumerate(row, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.alignment = center_align
            cell.border = border

            if str(row[0]).strip().upper() == "TOTAL":
                cell.font = bold_font
                cell.fill = PatternFill(start_color="D9D9D9", fill_type="solid")

    # -------- FREEZE HEADER --------
    ws.freeze_panes = f"A{start_row+1}"

    # -------- AUTO WIDTH --------
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        ws.column_dimensions[col_letter].width = max_length + 3

    # -------- SAVE --------
    output = BytesIO()
    wb.save(output)
    return output.getvalue()
