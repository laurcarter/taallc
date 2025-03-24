import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font
from io import BytesIO
import re

def run_full_pl_macro(file_bytes):
    wb = load_workbook(filename=BytesIO(file_bytes))
    ws = wb.active

    focus_ws = wb.create_sheet(title="Focus")
    ssoi_ws = wb.create_sheet(title="SSOI")

    max_row = ws.max_row
    max_col = ws.max_column

    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            focus_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            ssoi_ws.cell(row=cell.row, column=cell.column, value=cell.value)

    def process_sheet(sheet):
        for row in range(1, max_row + 1):
            a_val = sheet.cell(row=row, column=1).value or ""
            if "(" in a_val:
                parts = a_val.split("(", 1)
                sheet.cell(row=row, column=1).value = parts[0].strip()
                sheet.cell(row=row, column=2).value = parts[1].strip()

        for row in range(1, max_row + 1):
            b_val = sheet.cell(row=row, column=2).value or ""
            if "/" in b_val:
                parts = b_val.split("/")
                sheet.cell(row=row, column=2).value = parts[0].strip()
                sheet.cell(row=row, column=3).value = parts[1].replace(")", "").strip()
            sheet.cell(row=row, column=2).value = str(sheet.cell(row=row, column=2).value).replace("(", "")
            sheet.cell(row=row, column=3).value = str(sheet.cell(row=row, column=3).value).replace("/", "").replace(")", "")

    process_sheet(focus_ws)
    process_sheet(ssoi_ws)

    for row in range(1, max_row + 1):
        val = ws.cell(row=row, column=2).value
        focus_ws.cell(row=row, column=5, value=val)
        ssoi_ws.cell(row=row, column=5, value=val)

    for sheet in [focus_ws, ssoi_ws]:
        for row in range(1, max_row + 1):
            sheet.cell(row=row, column=3).value = None
            sheet.cell(row=row, column=4).value = sheet.cell(row=row, column=5).value
            sheet.cell(row=row, column=5).value = None

    for sheet in [focus_ws, ssoi_ws]:
        sheet.insert_cols(1, amount=2)

    def rearrange_rows(sheet):
        for row in range(5, max_row + 1):
            sheet.cell(row=row, column=5).value = sheet.cell(row=row, column=3).value
            sheet.cell(row=row, column=4).value = sheet.cell(row=row, column=2).value
            sheet.cell(row=row, column=3).value = sheet.cell(row=row, column=6).value
            sheet.cell(row=row, column=2).value = None
            sheet.cell(row=row, column=6).value = None

    rearrange_rows(focus_ws)
    rearrange_rows(ssoi_ws)

    focus_ws["C4"] = "Focus"
    focus_ws["D4"] = "Amount"
    focus_ws["E4"] = "Description"
    focus_ws["F4"] = "Totals"

    ssoi_ws["C4"] = "SSOI"
    ssoi_ws["D4"] = "Amount"
    ssoi_ws["E4"] = "Description"
    ssoi_ws["F4"] = "Totals"

    black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    white_font = Font(color="FFFFFF")

    for sheet in [focus_ws, ssoi_ws]:
        for col in range(3, 7):
            cell = sheet.cell(row=7, column=col)
            cell.fill = black_fill
            cell.font = white_font

    # Add comma formatting and set column widths
    for ws_target in [focus_ws, ssoi_ws]:
        for col in ['D', 'F']:
            for row in range(8, max_row + 1):
                cell = ws_target[f"{col}{row}"]
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'
        ws_target.column_dimensions['E'].width = 20

    # Add padding to all single-digit values in column C (as text with leading 0)
    for ws_target in [focus_ws, ssoi_ws]:
        for row in range(5, max_row + 1):
            val = ws_target.cell(row=row, column=3).value
            if val and isinstance(val, (int, str)) and str(val).isdigit() and len(str(val)) == 1:
                ws_target.cell(row=row, column=3).value = f"0{val}"

    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream
