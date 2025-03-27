from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def collapse_sheet(file_bytes):
    # Ensure file_bytes is a BytesIO object if it's raw bytes
    if not isinstance(file_bytes, BytesIO):
        file_bytes = BytesIO(file_bytes)  # Convert raw bytes to BytesIO if needed

    wb = load_workbook(file_bytes)  # Load workbook directly from BytesIO
    ws = wb.active  # Active sheet

    # Create a new sheet for the cleaned data
    clean_ws = wb.create_sheet("CleanedSheet")

    # Loop through each row in the sheet
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
        account_names = ""
        balance_found = False
    
        # Step 1: Find the first numeric value (balance column) in the row
        for j in range(0, len(row)):  # Iterate through all columns in the row
            cell = row[j]
            
            # Check for a numeric value (balance)
            if isinstance(cell.value, (int, float)) and cell.value is not None:
                balance_found = True
                balance_column = j  # This is the column of the balance (numeric value)
                balance = cell.value  # The balance value
                clean_ws.cell(row=i, column=2).value = balance  # Store the balance in column B
                break  # We found the balance column, no need to check further
    
        # Step 2: After finding the balance column, collect account names from cells to the left
        if balance_found:
            account_names = ""
            for j in range(0, balance_column):  # Only check the cells to the left of the balance column
                cell = row[j]
                
                # Check if the cell is a non-empty string (account name)
                if isinstance(cell.value, str) and cell.value.strip() != "":
                    account_names += f" {cell.value}"  # Concatenate account names
    
            account_names = account_names.strip()  # Remove any leading/trailing spaces
            
            # Only write to the CleanedSheet if we have non-empty account names
            if account_names:
                clean_ws.cell(row=i, column=1).value = account_names  # Place account names in column A

    # Step 3: Insert a header row for the cleaned sheet
    clean_ws.insert_rows(1)
    clean_ws.cell(row=1, column=1).value = "Account Names"
    clean_ws.cell(row=1, column=2).value = "Balance"

    # Move the CleanedSheet to the front of the workbook
    wb._sheets = [wb["CleanedSheet"]] + [ws for ws in wb.worksheets if ws.title != "CleanedSheet"]

    # Save the workbook and return the processed data
    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)

    return output_stream
