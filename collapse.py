from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def collapse_sheet(file_bytes):
    wb = load_workbook(filename=BytesIO(file_bytes))  # Ensure file_bytes is wrapped in BytesIO
    ws = wb.active  # Active sheet

    # Create a temporary sheet to hold the data
    temp_ws = wb.create_sheet("TempSheet")

    # Copy the content of the active sheet to the temporary sheet
    for row in ws.iter_rows():
        for cell in row:
            temp_ws[cell.coordinate].value = cell.value

    # Step 1: Create the CleanedSheet for the final output
    clean_ws = wb.create_sheet("CleanedSheet")

    # Loop through each row in the temp sheet
    for i, row in enumerate(temp_ws.iter_rows(min_row=1, max_row=temp_ws.max_row), 1):
        account_names = ""
        balance_found = False
    
        # Step 2: Search for the first numeric value (balance column) in columns A to N
        for j in range(1, 15):  # Check columns A to N (columns 1 to 14)
            if j-1 >= len(row):  # Check if j-1 is within the available columns in the row
                continue
            cell = row[j-1]  # Access the cell (1-indexed)
            
            if cell is not None and isinstance(cell.value, (int, float)):  # Check if it's a numeric balance value
                balance = cell.value  # First numeric value is considered the balance
                clean_ws.cell(row=i, column=2).value = balance  # Place the balance in column B
                balance_found = True
                balance_column = j
                break
    
        # Step 3: After finding the balance column, gather all non-numeric values (account names)
        if balance_found:
            account_names = ""
            for j in range(1, balance_column):  # Loop over columns to the left of the balance column
                if j-1 >= len(row):  # Check if j-1 is within the available columns in the row
                    continue
                cell = row[j-1]  # Access the cell (1-indexed)
                
                if cell is not None and not isinstance(cell.value, (int, float)) and cell.value != "":
                    account_names += f" {cell.value}"  # Concatenate account names to the string
    
            account_names = account_names.strip()  # Remove any leading/trailing spaces
            
            if account_names:  # Only write to the sheet if account names are not empty
                clean_ws.cell(row=i, column=1).value = account_names  # Place account names in column A


    # Step 4: Insert a row at the top of the CleanedSheet for headers
    clean_ws.insert_rows(1)
    clean_ws.cell(row=1, column=1).value = "Account Names"
    clean_ws.cell(row=1, column=2).value = "Balance"

    # Move the CleanedSheet to the front of the workbook
    wb._sheets = [wb["CleanedSheet"]] + [ws for ws in wb.worksheets if ws.title != "CleanedSheet"]

    # ---------------------------
    # Scrubbing Non-None Values in Column A
    # ---------------------------
    
    # Iterate through rows in column A of the CleanedSheet to remove 'None' values within cells
    for row in clean_ws.iter_rows(min_col=1, max_col=1, min_row=2, max_row=clean_ws.max_row):
        for cell in row:
            # If the cell value is None or contains only whitespace, replace it with an empty string
            if cell.value is None or str(cell.value).strip() == "":
                cell.value = ""  # Replace None or empty strings with an empty string in column A

    # Save the workbook and return the processed data
    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)

    return output_stream
