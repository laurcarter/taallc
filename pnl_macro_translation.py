import openpyxl
from openpyxl import load_workbook
from io import BytesIO

def run_full_pl_macro(file_bytes):
    # Ensure file_bytes is a BytesIO object
    if isinstance(file_bytes, BytesIO):
        file_obj = file_bytes
    else:
        file_obj = BytesIO(file_bytes)

    # Load the workbook
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

        # SSOI Sheet: Right part to B, Left part to C, strip ')' if present
        val = ssoi_ws.cell(row=row, column=2).value
        if val and '/' in str(val):
            parts = str(val).split('/')
            right = parts[1].strip().replace(")", "")
            left = parts[0].strip().replace("(", "")
            ssoi_ws.cell(row=row, column=2).value = right
            ssoi_ws.cell(row=row, column=3).value = left

    # Copy Column B from original ws to Column E in Focus and SSOI
    for row in range(1, max_row + 1):
        val = ws.cell(row=row, column=2).value
        focus_ws.cell(row=row, column=5, value=val)
        ssoi_ws.cell(row=row, column=5, value=val)

    # Clear C and D, move E to D
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(1, max_row + 1):
            sheet.cell(row=row, column=3).value = None
            sheet.cell(row=row, column=4).value = sheet.cell(row=row, column=5).value
            sheet.cell(row=row, column=5).value = None

    # Move the whole sheet over by two columns (Shift A and B to the right)
    for sheet in [focus_ws, ssoi_ws]:
        sheet.insert_cols(1, amount=2)

    # Move data in columns starting from row 5
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(5, max_row + 1):
            c_val = sheet.cell(row=row, column=3).value
            d_val = sheet.cell(row=row, column=4).value
            f_val = sheet.cell(row=row, column=6).value

            sheet.cell(row=row, column=5, value=c_val)  # C → E
            sheet.cell(row=row, column=3, value=d_val)  # D → C
            sheet.cell(row=row, column=4, value=f_val)  # F → D

            # Clear original cells
            sheet.cell(row=row, column=2).value = None
            sheet.cell(row=row, column=6).value = None

    # Add leading zero to single-digit values in SSOI column C (now col 3)
    for row in range(5, max_row + 1):
        val = ssoi_ws.cell(row=row, column=3).value
        if isinstance(val, (int, str)):
            val_str = str(val).strip()
            if val_str.isdigit() and len(val_str) == 1:
                ssoi_ws.cell(row=row, column=3, value=f"0{val_str}")

    # Sort SSOI by column C ascending
    ssoi_data = []
    for row in ssoi_ws.iter_rows(min_row=5, max_row=max_row, min_col=3, max_col=5, values_only=True):
        ssoi_data.append(row)

    ssoi_data.sort(key=lambda x: (str(x[0]) if x[0] is not None else ""))

    for idx, (col_c, col_d, col_e) in enumerate(ssoi_data, start=5):
        ssoi_ws.cell(row=idx, column=3, value=col_c)
        ssoi_ws.cell(row=idx, column=4, value=col_d)
        ssoi_ws.cell(row=idx, column=5, value=col_e)

    # Add column titles in row 4
    focus_ws["C4"] = "Focus"
    focus_ws["D4"] = "Amount"
    focus_ws["E4"] = "Description"
    focus_ws["F4"] = "Totals"

    ssoi_ws["C4"] = "SSOI"
    ssoi_ws["D4"] = "Amount"
    ssoi_ws["E4"] = "Description"
    ssoi_ws["F4"] = "Totals"

    # Shift rows 4-6 down by inserting above row 4
    focus_ws.insert_rows(4, amount=3)
    ssoi_ws.insert_rows(4, amount=3)

    # Sort data in Focus sheet by column C (asc), then D (desc)
    focus_data = []
    for row in focus_ws.iter_rows(min_row=8, max_row=max_row, min_col=3, max_col=5, values_only=True):
        focus_data.append(row)
    focus_data.sort(key=lambda x: (str(x[0]), -float(x[1]) if x[1] is not None and str(x[1]).replace(",", "").replace(".", "").isdigit() else 0))

    for i, row_data in enumerate(focus_data, start=8):
        for j, val in enumerate(row_data, start=3):
            focus_ws.cell(row=i, column=j).value = val

    # Sort data in SSOI sheet by column C (asc), then D (desc)
    ssoi_data = []
    for row in ssoi_ws.iter_rows(min_row=8, max_row=max_row, min_col=3, max_col=5, values_only=True):
        ssoi_data.append(row)
    ssoi_data.sort(key=lambda x: (str(x[0]), -float(x[1]) if x[1] is not None and str(x[1]).replace(",", "").replace(".", "").isdigit() else 0))

    for i, row_data in enumerate(ssoi_data, start=8):
        for j, val in enumerate(row_data, start=3):
            ssoi_ws.cell(row=i, column=j).value = val

    # Style header row (row 7) black background, white font
    black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    white_font = Font(color="FFFFFF")

    for sheet in [focus_ws, ssoi_ws]:
        for col in range(3, 7):  # Columns C to F
            cell = sheet.cell(row=7, column=col)
            cell.fill = black_fill
            cell.font = white_font

    # Format columns D and F to "#,##0"
    for sheet in [focus_ws, ssoi_ws]:
        for col_letter in ["D", "F"]:
            for row in range(8, max_row + 1):
                cell = sheet[f"{col_letter}{row}"]
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0"

    # Double the width of column E
    for sheet in [focus_ws, ssoi_ws]:
        sheet.column_dimensions["E"].width = 20

    # Clear empty rows after last data row in column C
    for sheet in [focus_ws, ssoi_ws]:
        last_data_row = sheet.max_row
        for row in reversed(range(1, last_data_row + 1)):
            if sheet.cell(row=row, column=3).value not in [None, ""]:
                break
        for clear_row in range(row + 1, last_data_row + 1):
            for col in range(1, sheet.max_column + 1):
                sheet.cell(row=clear_row, column=col).value = None

    # Save the result into a BytesIO object and return
    output_stream = BytesIO()
    wb.save
