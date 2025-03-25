import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font
from io import BytesIO
import re

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

    # -----------------------------
    # Step 1: Pre-validate rows with expected formatting
    # -----------------------------
    invalid_rows_focus = set()
    invalid_rows_ssoi = set()
    
    for row in range(1, max_row + 1):
        a_focus = focus_ws.cell(row=row, column=1).value
        a_ssoi = ssoi_ws.cell(row=row, column=1).value
    
        if not (a_focus and isinstance(a_focus, str) and '(' in a_focus):
            invalid_rows_focus.add(row)
        if not (a_ssoi and isinstance(a_ssoi, str) and '(' in a_ssoi):
            invalid_rows_ssoi.add(row)
    
    # -----------------------------
    # Step 2: Delimit Column A by '(', put second part in Column B
    # -----------------------------
    for row in range(1, max_row + 1):
        if row not in invalid_rows_focus:
            val = focus_ws.cell(row=row, column=1).value
            parts = val.split('(', 1)
            focus_ws.cell(row=row, column=1, value=parts[0].strip())
            focus_ws.cell(row=row, column=2, value=parts[1].strip())
    
        if row not in invalid_rows_ssoi:
            val = ssoi_ws.cell(row=row, column=1).value
            parts = val.split('(', 1)
            ssoi_ws.cell(row=row, column=1, value=parts[0].strip())
            ssoi_ws.cell(row=row, column=2, value=parts[1].strip())
    
    # -----------------------------
    # Step 3: Split Column B by '/', reassign parts (Focus: B→B+C, SSOI: B→B+C, swapped)
    # -----------------------------
    for row in range(1, max_row + 1):
        if row not in invalid_rows_focus:
            val = focus_ws.cell(row=row, column=2).value
            if val and '/' in val:
                parts = val.split('/')
                focus_ws.cell(row=row, column=2, value=parts[0].strip().replace("(", ""))
                focus_ws.cell(row=row, column=3, value=parts[1].strip().replace(")", "").replace("/", ""))
            else:
                invalid_rows_focus.add(row)
    
        if row not in invalid_rows_ssoi:
            val = ssoi_ws.cell(row=row, column=2).value
            if val and '/' in val:
                parts = val.split('/')
                right = parts[1].strip().replace(")", "")
                left = parts[0].strip().replace("(", "")
                ssoi_ws.cell(row=row, column=2, value=right)
                ssoi_ws.cell(row=row, column=3, value=left)
            else:
                invalid_rows_ssoi.add(row)
    
    # -----------------------------
    # Step 4: Copy Column B from original ws to Column E in Focus and SSOI
    # -----------------------------
    for row in range(1, max_row + 1):
        if row not in invalid_rows_focus:
            val = ws.cell(row=row, column=2).value
            focus_ws.cell(row=row, column=5, value=val)
        if row not in invalid_rows_ssoi:
            val = ws.cell(row=row, column=2).value
            ssoi_ws.cell(row=row, column=5, value=val)
    
    # -----------------------------
    # Step 5: Clear columns C and D, move E to D
    # -----------------------------
    for row in range(1, max_row + 1):
        if row not in invalid_rows_focus:
            focus_ws.cell(row=row, column=3).value = None
            focus_ws.cell(row=row, column=4, value=focus_ws.cell(row=row, column=5).value)
            focus_ws.cell(row=row, column=5).value = None
    
        if row not in invalid_rows_ssoi:
            ssoi_ws.cell(row=row, column=3).value = None
            ssoi_ws.cell(row=row, column=4, value=ssoi_ws.cell(row=row, column=5).value)
            ssoi_ws.cell(row=row, column=5).value = None


    # Clear C and D, move E to D
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(1, max_row + 1):
            sheet.cell(row=row, column=3).value = None
            sheet.cell(row=row, column=4).value = sheet.cell(row=row, column=5).value
            sheet.cell(row=row, column=5).value = None

        # Shift entire sheet right by two columns
    for sheet in [focus_ws, ssoi_ws]:
        sheet.insert_cols(1, amount=2)

    # Rearrange data starting from row 5
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(5, max_row + 1):
            # Move column C to E, D to C, F to D
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


    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream


