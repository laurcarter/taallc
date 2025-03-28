import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

# Placeholder eFocus Function
def efocus_focus(file_bytes, client_data_bytes):
    # Load the workbook for the uploaded Excel file
    wb = load_workbook(filename=BytesIO(file_bytes))

    # Ensure we are working with the correct sheet (Focus sheet)
    if 'Focus' not in wb.sheetnames:
        st.error("Focus sheet not found in the uploaded file.")
        return None  # Exit if Focus sheet is not found

    focus_ws = wb['Focus']  # Get the Focus sheet

    # Debug: Print the sheet names to verify that we are looking at the right one
    st.write(f"Sheet Names in Workbook: {wb.sheetnames}")
    
    # Load the client data from the uploaded client data Excel file
    client_data = pd.read_excel(BytesIO(client_data_bytes))  # Load client data as DataFrame

    # Debug: Check the columns and preview the first few rows of the client data
    st.write("Client Data Columns:")
    st.write(client_data.columns)
    st.write("First few rows of Client Data:")
    st.write(client_data.head())

    # Ensure that 'ClientName' column exists (checking for exact matches)
    if 'ClientName' not in client_data.columns:
        st.error("'ClientName' column not found in the client data.")
        return None  # Exit if the column is not found

    # Extract client names from row 1 (starting from column C)
    client_names = []
    for col in range(3, focus_ws.max_column + 1):  # Starting from column C (index 3)
        client_name = focus_ws.cell(row=1, column=col).value
        if client_name:
            client_names.append(client_name)

    # Debug: Print out the client names from the Focus sheet
    st.write("Client Names in Focus Sheet:")
    st.write(client_names)

    # Example of checking if the client names match
    for client_name in client_names:
        # Try to find the client data in the client_data (match by 'ClientName')
        client_row = client_data[client_data['ClientName'] == client_name]
        
        if not client_row.empty:
            # Process the client-specific data (for now just showing it)
            st.write(f"Processing data for {client_name}")
            st.write(client_row)
            # Update Focus sheet based on client-specific data if needed
        else:
            st.error(f"No data found for client {client_name}.")

    # Save the workbook and return it
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output  # Return the transformed file as BytesIO



# ---------- Streamlit App Flow ----------
st.set_page_config(page_title="eFocus Transformation", layout="wide")

# Step 1: Upload the main Excel file
st.title("📂 Upload Your Excel File for eFocus Transformation")
uploaded_file = st.file_uploader("Upload your main Excel file", type=["xlsx"])

# Step 2: Upload the client data Excel file
st.title("📂 Upload Client Data Excel File")
uploaded_client_data_file = st.file_uploader("Upload your client data file", type=["xlsx"])

if uploaded_file and uploaded_client_data_file:
    file_bytes = uploaded_file.read()
    client_data_bytes = uploaded_client_data_file.read()

    # Process the file through the eFocus Focus Grouping function
    transformed_file = efocus_focus(file_bytes, client_data_bytes)

    # Store the transformed file in session state
    st.session_state.excel_bytes = transformed_file

    st.success("File processed. The 'Focus' sheet has been created and updated.")

    # Provide option to download the transformed file
    st.download_button(
        label="Download Transformed File",
        data=st.session_state.excel_bytes,
        file_name="efocus_transformed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
