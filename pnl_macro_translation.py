import openpyxl
from openpyxl import load_workbook
from io import BytesIO

def run_full_pl_macro(file_bytes):
    # Ensure file_bytes is a BytesIO object
    if isinstance(file_bytes, BytesIO):
        file_obj = file_bytes
    else:
        file_obj = BytesIO(file_bytes)

    wb = load_workbook(filename=file_obj)
    ws = wb.active  # Get the active worksheet

    # Create Focus and SSOI worksheets
    focus_ws = wb.create_sheet(title="Focus")
    ssoi_ws = wb.create_sheet(title="SSOI")

    # Find the last used row and column in the original sheet
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
            cell.value = f"0{val}"  # Add leading zero for single-character
        elif len(val) >= 2:
            cell.value = f"'{val}"  # Add apostrophe to treat as text

    # Insert a temporary column to format IDs as text in SSOI sheet (mimic the Insert Column behavior)
    ssoi_ws.insert_cols(4)

    # Apply the NumberFormat for column C in the SSOI sheet (mimic Excel's "@")
    ssoi_ws.column_dimensions['C'].number_format = '@'

    # *** NEW CODE START ***
    # Copy the formatted values back to column C and delete the temporary column D
    for row in range(1, max_row + 1):
        ssoi_ws.cell(row=row, column=3, value=ssoi_ws.cell(row=row, column=4).value)
    
    # Delete the temporary column D (column 4)
    ssoi_ws.delete_cols(4)
    
    # *** NEW CODE END ***

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

    # Strip out the opening parenthesis in Column B for both sheets
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(1, max_row + 1):
            val = sheet.cell(row=row, column=2).value
            if val:
                # Ensure val is a string before calling replace()
                val = str(val)
                sheet.cell(row=row, column=2).value = val.replace("(", "")
    
    # Strip out the slash and closing parenthesis in Column C for both sheets
    for sheet in [focus_ws, ssoi_ws]:
        for row in range(1, max_row + 1):
            val = sheet.cell(row=row, column=3).value
            if val:
                # Ensure val is a string before calling replace()
                val = str(val)
                sheet.cell(row=row, column=3).value = val.replace("/", "").replace(")", "")

    # Wipe out column B of the SSOI sheet
    for row in range(1, max_row + 1):
        ssoi_ws.cell(row=row, column=2).value = None
    
    # Copy Column C from Focus sheet to Column B in SSOI sheet
    for row in range(1, max_row + 1):
        focus_value = focus_ws.cell(row=row, column=3).value
        ssoi_ws.cell(row=row, column=2).value = focus_value

    # Wipe out column B of the Focus sheet
    for row in range(1, max_row + 1):
        focus_ws.cell(row=row, column=2).value = None
    
    # Copy Column C from SSOI sheet to Column B in Focus sheet
    for row in range(1, max_row + 1):
        ssoi_value = ssoi_ws.cell(row=row, column=3).value
        focus_ws.cell(row=row, column=2).value = ssoi_value

  # Clear column C in both Focus and SSOI sheets
    for row in range(1, max_row + 1):
        focus_ws.cell(row=row, column=3).value = None
        ssoi_ws.cell(row=row, column=3).value = None

    # Copy Column B from original sheet to Column D in both Focus and SSOI sheets
    for row in range(1, max_row + 1):
        original_value = ws.cell(row=row, column=2).value
        focus_ws.cell(row=row, column=4).value = original_value
        ssoi_ws.cell(row=row, column=4).value = original_value

    # Move the entire sheet over by two columns in both Focus and SSOI sheets
    focus_ws.insert_cols(1, 2)  # Insert two columns at the beginning of the Focus sheet
    ssoi_ws.insert_cols(1, 2)   # Insert two columns at the beginning of the SSOI sheet
    
    # Move data in columns starting from row 5 in the Focus sheet
    for row in range(5, max_row + 1):
        focus_ws.cell(row=row, column=3).offset(0, 2).value = focus_ws.cell(row=row, column=3).value
        focus_ws.cell(row=row, column=3).value = None  # Clear original cell
    
        focus_ws.cell(row=row, column=4).offset(0, -1).value = focus_ws.cell(row=row, column=4).value
        focus_ws.cell(row=row, column=4).value = None  # Clear original cell
    
        focus_ws.cell(row=row, column=6).offset(0, -2).value = focus_ws.cell(row=row, column=6).value
        focus_ws.cell(row=row, column=6).value = None  # Clear original cell
    
    # Same process for the SSOI sheet
    for row in range(5, max_row + 1):
        ssoi_ws.cell(row=row, column=3).offset(0, 2).value = ssoi_ws.cell(row=row, column=3).value
        ssoi_ws.cell(row=row, column=3).value = None  # Clear original cell
    
        ssoi_ws.cell(row=row, column=4).offset(0, -1).value = ssoi_ws.cell(row=row, column=4).value
        ssoi_ws.cell(row=row, column=4).value = None  # Clear original cell
    
        ssoi_ws.cell(row=row, column=6).offset(0, -2).value = ssoi_ws.cell(row=row, column=6).value
        ssoi_ws.cell(row=row, column=6).value = None  # Clear original cell
# new
    # Add leading zeros to all single-digit numbers in column C of SSOI sheet
    for row in range(5, max_row + 1):
        val = str(ssoi_ws.cell(row=row, column=3).value).strip()  # Get the value from column C
        if val.isdigit() and len(val) == 1:  # Check if it's a single-digit number
            ssoi_ws.cell(row=row, column=3).value = f"0{val}"  # Add leading zero
    
    # Sort columns C to E using the updated values in column C for the SSOI sheet
    # Extract data from columns C to E
    ssoi_data = []
    for row in ssoi_ws.iter_rows(min_row=5, max_row=max_row, min_col=3, max_col=5, values_only=True):
        ssoi_data.append(row)
    
    # Sort by the values in column C (ascending)
    ssoi_data.sort(key=lambda x: (str(x[0]) if x[0] is not None else ""))
    
    # Write sorted data back into the SSOI sheet
    for idx, (col_c, col_d, col_e) in enumerate(ssoi_data, start=5):
        ssoi_ws.cell(row=idx, column=3, value=col_c)
        ssoi_ws.cell(row=idx, column=4, value=col_d)
        ssoi_ws.cell(row=idx, column=5, value=col_e)


    
    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream
