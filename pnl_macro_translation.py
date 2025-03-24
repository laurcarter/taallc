from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def run_full_pl_macro(file_bytes):
    # Ensure the input is BytesIO
    if not isinstance(file_bytes, BytesIO):
        file_bytes = BytesIO(file_bytes)
    
    wb = load_workbook(filename=file_bytes)
    ws = wb.active

    focus_ws = wb.create_sheet(title="Focus")
    ssoi_ws = wb.create_sheet(title="SSOI")

    max_row = ws.max_row
    max_col = ws.max_column

    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            focus_ws.cell(row=cell.row, column=cell.column).value = cell.value
            ssoi_ws.cell(row=cell.row, column=cell.column).value = cell.value

    def format_ssoi_ids(sheet):
        for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=3, max_col=3):
            for cell in row:
                if cell.value is not None:
                    val = str(cell.value).strip()
                    if len(val) == 1:
                        cell.value = f"'0{val}"
                    elif len(val) > 1:
                        cell.value = f"'{val}"

    format_ssoi_ids(ssoi_ws)

    def clean_and_parse(sheet, is_focus=True):
        for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=1):
            for cell in row:
                if cell.value and "(" in str(cell.value):
                    parts = str(cell.value).split("(")
                    cell.value = parts[0].strip()
                    sheet.cell(row=cell.row, column=2).value = parts[1].strip() if len(parts) > 1 else ""

        for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=2, max_col=2):
            for cell in row:
                val = str(cell.value)
                if "/" in val:
                    left, right = val.split("/")
                    if is_focus:
                        cell.value = left.strip()
                        sheet.cell(row=cell.row, column=3).value = right.strip()
                    else:
                        cell.value = right.strip()
                        sheet.cell(row=cell.row, column=3).value = left.strip().rstrip(")")

        for col in [2, 3]:
            for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=col, max_col=col):
                for cell in row:
                    if cell.value:
                        val = str(cell.value).replace("(", "").replace(")", "").replace("/", "").strip()
                        cell.value = val

    clean_and_parse(focus_ws, is_focus=True)
    clean_and_parse(ssoi_ws, is_focus=False)

    def copy_col_b_to_d(source, target):
        for row in range(1, max_row + 1):
            target.cell(row=row, column=4).value = source.cell(row=row, column=2).value

    copy_col_b_to_d(ws, focus_ws)
    copy_col_b_to_d(ws, ssoi_ws)

    def finalize_layout(sheet):
        for row in range(5, max_row + 1):
            sheet.cell(row=row, column=6).value = sheet.cell(row=row, column=3).value
            sheet.cell(row=row, column=3).value = ""
            sheet.cell(row=row, column=4).value = sheet.cell(row=row, column=5).value
            sheet.cell(row=row, column=5).value = ""

        sheet.insert_cols(1, amount=2)
        sheet.insert_rows(4, amount=3)
        sheet.cell(row=4, column=3).value = sheet.title
        sheet.cell(row=4, column=4).value = "Amount"
        sheet.cell(row=4, column=5).value = "Description"
        sheet.cell(row=4, column=6).value = "Totals"

    finalize_layout(focus_ws)
    finalize_layout(ssoi_ws)

    # Skip additional formatting/subtotals for now
    # You can extend this later if needed

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
