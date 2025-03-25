import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO

def run_full_pl_macro(file_bytes):
    # Ensure file_bytes is a BytesIO object
    if isinstance(file_bytes, BytesIO):
        file_obj = file_bytes
    else:
        file_obj = BytesIO(file_bytes)

    wb = load_workbook(filename=file_obj)
    ws = wb.active

    # Create Focus and SSOI worksheets
    focus_ws = wb.create_sheet(title="Focus")
    ssoi_ws = wb.create_sheet(title="SSOI")

    # Find last used row and column
    max_row = ws.max_row
    max_col = ws.max_column

    # Copy content from original sheet to Focus and SSOI as plain values
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            focus_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            ssoi_ws.cell(row=cell.row, column=cell.column, value=cell.value)

    # Format IDs in column C of SSOI sheet as text with leading zeros
    for row in range(1, max_row + 1):
        cell = ssoi_ws.cell(row=row, column=3)
        val = str(cell.value).strip() if cell.value is not None else ""
        if len(val) == 1:
            cell.value = f"0{val}"
        elif len(val) >= 2:
            cell.value = val

    # Delimit Column A by '(' and extract to Column B (both sheets)
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(1, max_row + 1):
            val = sheet.cell(row=row, column=1).value
            if val and '(' in str(val):
                parts = str(val).split('(', 1)
                sheet.cell(row=row, column=1).value = parts[0].strip()
                sheet.cell(row=row, column=2).value = parts[1].strip()

    # Delimit Column B by '/' and process (Focus and SSOI)
    for row in range(1, max_row + 1):
        # Focus Sheet: Left part to B, Right part to C
        val = focus_ws.cell(row=row, column=2).value
        if val and '/' in str(val):
            parts = str(val).split('/')
            focus_ws.cell(row=row, column=2).value = parts[0].strip().replace("(", "")
            focus_ws.cell(row=row, column=3).value = parts[1].strip().replace(")", "").replace("/", "")
            sheet.cell(row=row, column=5).value = None

    # Copy Column B from original sheet to Column E in Focus and SSOI
    for row in range(1, max_row + 1):
        val = ws.cell(row=row, column=2).value
        focus_ws.cell(row=row, column=5, value=val)
        ssoi_ws.cell(row=row, column=5, value=val)
    
    # Clear C and D
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(1, max_row + 1):
            sheet.cell(row=row, column=3).value = None
            sheet.cell(row=row, column=4).value = None
    
    # Move column E over to column D in both sheets
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(1, max_row + 1):
            sheet.cell(row=row, column=4).value = sheet.cell(row=row, column=5).value
            sheet.cell(row=row, column=5).value = None
    
    # Insert two columns at the beginning (A:B) for both sheets
    for sheet in [focus_ws, ssoi_ws]:
        sheet.insert_cols(1, amount=2)
    
    # Move data starting from row 5 for Focus and SSOI sheets
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(5, max_row + 1):
            # Move data as per the rearranging logic
            c_val = sheet.cell(row=row, column=3).value
            d_val = sheet.cell(row=row, column=4).value
            f_val = sheet.cell(row=row, column=6).value
    
            if c_val is not None and c_val != "":
                sheet.cell(row=row, column=5, value=c_val)  # C → E
            if d_val is not None and d_val != "":
                sheet.cell(row=row, column=3, value=d_val)  # D → C
            if f_val is not None and f_val != "":
                sheet.cell(row=row, column=4, value=f_val)  # F → D
    
            # Clear original cells
            sheet.cell(row=row, column=2).value = None
            sheet.cell(row=row, column=6).value = None



    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream
