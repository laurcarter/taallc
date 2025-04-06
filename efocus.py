import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font

def match_and_copy_values(focus_ws, focus_target_ws):
    # Loop through rows 8 to 40 in column I of the Focus sheet
    for row in range(8, 41):  # Rows 8 to 40 (inclusive)
        focus_value = focus_ws.cell(row=row, column=9).value  # Value in column I
        
        # If the cell has a value
        if focus_value:
            # Strip any "I" or leading zeros from the Focus value
            focus_value_stripped = str(focus_value).lstrip("I0").strip()
            
            # Now search for this stripped value in FocusTarget column A
            for target_row in range(1, focus_target_ws.max_row + 1):
                target_value = str(focus_target_ws.cell(row=target_row, column=1).value).lstrip("I0").strip()
                
                # If a match is found, get the value from column J in Focus and paste it in column B of FocusTarget
                if focus_value_stripped == target_value:
                    focus_value_j = focus_ws.cell(row=row, column=10).value  # Get value from column J of Focus
                    focus_target_ws.cell(row=target_row, column=2, value=focus_value_j)  # Paste it in column B of FocusTarget
                    break  # Exit loop once a match is found

def efocus_focus(file_bytes, client_data_bytes):
    # Ensure both files are wrapped in BytesIO if they aren't already wrapped
    file_bytes_io = BytesIO(file_bytes)  # Focus file
    client_data_bytes_io = BytesIO(client_data_bytes)  # Client data file
    
    # Load the Focus sheet from the uploaded file (file_bytes)
    wb = load_workbook(filename=file_bytes_io)
    focus_ws = wb['Focus']  # Assuming the Focus sheet is already available

    # Load the client data from the second uploaded file (client_data_bytes)
    client_data = pd.read_excel(client_data_bytes_io, header=None)  # Reading client data without headers

    client_names = []
    for col in range(2, client_data.shape[1], 2):
        cell_value = str(client_data.iloc[0, col]).strip()
        if cell_value and 'Unnamed' not in cell_value:
            client_names.append(cell_value)

    if not client_names:
        return None, None

    selected_client = client_names[0]  # Automatically choose the first client as an example

    # Create the "FocusTarget" sheet
    focus_target_ws = wb.create_sheet(title="FocusTarget")
    
    # Add the client data into FocusTarget sheet
    client_column_a = client_data.iloc[:, 0]
    for i, value in enumerate(client_column_a, start=1):
        focus_target_ws.cell(row=i, column=1, value=value)

    # Set the header for the client name in column B
    focus_target_ws.cell(row=1, column=2, value=selected_client)

    # Add the client data column B and column C
    client_column_b = client_data.iloc[:, 1]
    for i, value in enumerate(client_column_b, start=1):
        focus_target_ws.cell(row=i, column=3, value=value)

    # Call match_and_copy_values function
    match_and_copy_values(focus_ws, focus_target_ws)

    # Rename the FocusTarget sheet
    focus_target_ws.title = "Filing Items Focus"

    # Save the modified workbook to a BytesIO object
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output, selected_client
