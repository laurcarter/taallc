import openpyxl
from openpyxl import load_workbook
from io import BytesIO

#new 
def clean_ss01_column(ssoi_ws, max_row):
    # Loop through each cell in column C of the SSOI sheet starting from row 5
    for row in range(5, max_row + 1):
        cell = ssoi_ws.cell(row=row, column=3)
        cell_value = str(cell.value)  # Ensure cell value is a string
        
        # Check if the value starts with an apostrophe and then handle it
        if cell_value.startswith("'"):
            # Remove the apostrophe and any leading zeros
            cleaned_value = cell_value.lstrip("'0")
            
            # If the cleaned value is empty, set the value to '0', otherwise just the cleaned value
            if cleaned_value == "":
                cleaned_value = "0"
            
            cell.value = cleaned_value  # Update the cell with the cleaned value


# Sorting function for SSOI sheet column C
def sort_ssoi_sheet(ssoi_ws, max_row):
    # Loop through column C starting from row 5 to find and delete rows with empty cells
    for row in range(max_row, 4, -1):  # Go from the bottom to the top to avoid skipping rows
        cell = ssoi_ws.cell(row=row, column=3)
        if cell.value is None or cell.value == "":  # If the cell in column C is empty
            ssoi_ws.delete_rows(row)  # Delete the row

    # Create a list to hold rows with their corresponding values in column C
    rows = []

    # Loop through column C starting from row 5
    for row in range(5, max_row + 1):
        cell = ssoi_ws.cell(row=row, column=3)
        if cell.value is not None:
            c_value = str(cell.value).strip()
            
            # Append the entire row with its value in column C
            rows.append((row, c_value, [ssoi_ws.cell(row=row, column=col).value for col in range(1, ssoi_ws.max_column + 1)]))

    # Sort rows based on the value in column C (ascending)
    # First, sort numeric values (no letters), then alphanumeric
    rows.sort(key=lambda x: (int(x[1]) if x[1].isdigit() else float('inf'), x[1]))

    # Clear the existing values in the sheet starting from row 5
    for row in range(5, max_row + 1):
        for col in range(1, ssoi_ws.max_column + 1):
            ssoi_ws.cell(row=row, column=col).value = None

    # Write the sorted rows back into the sheet
    for idx, (original_row, _, row_values) in enumerate(rows, start=5):
        for col_idx, value in enumerate(row_values, start=1):
            ssoi_ws.cell(row=idx, column=col_idx).value = value

def sort_focus_sheet(focus_ws, max_row):
    # Create a list to store the rows with their values in column C
    focus_data = []
    
    # Loop through column C starting from row 8
    for row in range(8, max_row + 1):
        c_value = focus_ws.cell(row=row, column=3).value
        if c_value is not None:  # Only add rows with non-empty values in column C
            focus_data.append((row, c_value))  # Store the row number and the value in column C
    
    # Sort the rows by the value in column C (ascending order)
    focus_data.sort(key=lambda x: x[1])

    # Write the sorted data back to the Focus sheet
    for idx, (row, value) in enumerate(focus_data, start=8):
        focus_ws.cell(row=row, column=3).value = value  # Sort column C
        # Make sure the entire row moves with the sorted value in column C
        for col in range(1, 7):  # Adjust this if you need to cover more columns
            focus_ws.cell(row=row, column=col).value = focus_ws.cell(row=focus_data[idx][0], column=col).value

    # Now clear the remaining unsorted rows from row 8 downwards in column C (if any)
    for row in range(len(focus_data) + 8, max_row + 1):
        focus_ws.cell(row=row, column=3).value = None  # Clear unsorted rows



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
        
    # Clean the SSOI column C before sorting
    clean_ss01_column(ssoi_ws, max_row)

    # Now, call the sort function to sort the SSOI sheet based on column C
    sort_ssoi_sheet(ssoi_ws, max_row)
# new
    # Add column titles in row 4
    focus_ws["C4"] = "Focus"
    focus_ws["D4"] = "Amount"
    focus_ws["E4"] = "Description"
    focus_ws["F4"] = "Totals"
    
    ssoi_ws["C4"] = "SSOI"
    ssoi_ws["D4"] = "Amount"
    ssoi_ws["E4"] = "Description"
    ssoi_ws["F4"] = "Totals"
    
    # Shift everything below row 4 down by 3 rows in both sheets
    focus_ws.insert_rows(4, amount=3)  # Insert 3 rows at row 4 in focus_ws
    ssoi_ws.insert_rows(4, amount=3)   # Insert 3 rows at row 4 in ssoi_ws

    # Call the sort_focus_sheet function after the rest of the operations in the macro
    sort_focus_sheet(focus_ws, max_row)


    # Ensure to save the workbook after sorting if needed
    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream

