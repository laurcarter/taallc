import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

# Function to implement the logic for the 'eFocus' transformation
def efocus_focus(file_bytes, client_data_bytes):
    # Load the uploaded Excel workbook with the "Focus" sheet
    wb = load_workbook(filename=BytesIO(file_bytes))
    
    if 'Focus' not in wb.sheetnames:
        st.error("Focus sheet not found in the uploaded file.")
        return None  # Exit if Focus sheet is not found
    
    focus_ws = wb['Focus']  # Get the Focus sheet
    
    # Load the client data into a DataFrame
    try:
        client_data = pd.read_excel(BytesIO(client_data_bytes))  # Load client data as DataFrame
    except Exception as e:
        st.error(f"Error loading client data: {e}")
        return None  # Exit if client data fails to load

    # Debugging: Check the structure of client data
    st.write("Client Data Columns:", client_data.columns)

    # Extract all non-blank client names from row 1 (column 1 onwards)
    client_names = []
    for col in range(1, len(client_data.columns)):  # Start from column 1 (index 0)
        cell_value = str(client_data.iloc[0, col])  # Check first row, all columns
        if cell_value.strip() and cell_value.lower() != "nan" and "Unnamed" not in cell_value:  # Exclude empty, 'nan', or unnamed columns
            client_names.append(cell_value)

    # Display the client names as radio buttons (bubbles)
    if len(client_names) == 0:
        st.error("No valid client names found in the client data.")
        return None
    
    st.write("Valid client names:", client_names)  # Show the list of valid client names for debugging
    
    client_name = st.radio("Select a client from the list:", client_names)

    # Validate input (ensure the client name is selected)
    if not client_name:
        st.warning("No client selected. Process aborted.")
        return None
    
    # Find the column index of the selected client
    found_cell = None
    for col in range(1, len(client_data.columns)):  # Starting from column 1 (index 0)
        cell_value = str(client_data.iloc[0, col])  # Check first row, all columns
        if client_name.lower() == cell_value.lower():  # Case-insensitive match
            found_cell = col
            break

    # If client name is not found, exit
    if found_cell is None:
        st.error(f"No client found matching '{client_name}'.")
        return None

    # Create a new sheet called "FocusTarget" in the Focus workbook
    ws_target = wb.create_sheet(title="FocusTarget")

    # Step 1: Copy columns A and B from rows 1-275 from the Focus sheet to the FocusTarget sheet
    for row in range(1, 276):  # Rows 1 to 275
        ws_target.cell(row=row, column=1).value = focus_ws.cell(row=row, column=1).value  # Column A
        ws_target.cell(row=row, column=2).value = focus_ws.cell(row=row, column=2).value  # Column B

    # Step 2: Insert new column B in "FocusTarget" for client data
    client_column_data = []
    for row in range(1, 276):  # Get data from rows 2 to 275
        client_column_data.append(client_data.iloc[row - 1, found_cell])  # Adjust for 0-indexing

    # Step 3: Insert client data into the new column B in FocusTarget sheet
    for row, value in enumerate(client_column_data, start=2):
        ws_target.cell(row=row, column=2).value = value
    
    # Set column B header to the full client name found
    ws_target.cell(row=1, column=2).value = client_name

    # Save the transformed workbook to a BytesIO object
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Provide a success message
    st.success(f"eFocus was created with {client_name.upper()} template.")
    st.write(f"Please verify that {client_name.upper()} is what you were referring to.")
    
    # Return the transformed file as BytesIO
    return output



# Streamlit UI for the file upload and processing
st.set_page_config(page_title="eFocus Transformation", layout="wide")

st.title("📂 Upload Your Excel File for eFocus Transformation")

# File uploader for the main Excel file and the client data
uploaded_file = st.file_uploader("Upload the Focus Excel file", type=["xlsx"])
client_data_file = st.file_uploader("Upload the Client Data file", type=["xlsx"])

if uploaded_file and client_data_file:
    # Read the files as bytes
    file_bytes = uploaded_file.read()
    client_data_bytes = client_data_file.read()

    # Process the files with the efocus_focus function
    transformed_file = efocus_focus(file_bytes, client_data_bytes)

    if transformed_file:
        # Store the transformed file in session state
        st.session_state.excel_bytes = transformed_file

        # Provide option to download the transformed file
        st.download_button(
            label="Download Transformed File",
            data=st.session_state.excel_bytes,
            file_name="efocus_transformed_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
