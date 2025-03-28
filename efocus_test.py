import os
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import streamlit as st

# Path to the locally stored client data Excel file
local_file_path = "/Users/lauren.carter/Desktop/client_data.xlsx"

# Function to fetch client data from the local Excel file
def fetch_client_data_locally():
    # Check if the file exists
    if os.path.exists(local_file_path):
        # Load the Excel file into pandas
        client_data = pd.read_excel(local_file_path)
        return client_data
    else:
        print("Error: Client data file not found.")
        return None

# Placeholder function for eFocus transformation (this will be where your processing logic goes)
def efocus_focus(file_bytes):
    # Load the uploaded Excel file
    wb = load_workbook(filename=BytesIO(file_bytes))

    # Check if the "Focus" sheet exists
    if "Focus" not in wb.sheetnames:
        st.error("The uploaded file does not contain a sheet named 'Focus'.")
        return file_bytes  # Return the original file if the sheet is missing

    focus_ws = wb["Focus"]

    # Fetch the client-specific data from the local file
    client_data = fetch_client_data_locally()

    if client_data is None:
        st.error("Client data could not be fetched.")
        return file_bytes  # Return the original file if fetching data fails
    
    # Assume the client data has a column 'ClientName' and we want to use this name to look up data
    client_name = focus_ws.cell(row=1, column=1).value  # Example: Getting client name from 'Focus' sheet

    # Find the client data for this client
    client_row = client_data[client_data['ClientName'] == client_name]

    if not client_row.empty:
        # Apply client-specific transformations using the data
        focus_ws.cell(row=2, column=2).value = client_row.iloc[0]['SpecificColumn']  # Example update
    else:
        st.error(f"No client data found for {client_name}.")
    
    # Save the updated workbook to BytesIO and return it
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output  # Return the updated file as BytesIO

# ---------- Streamlit App Flow ----------
st.set_page_config(page_title="eFocus Transformation", layout="wide")

# Step 1: Upload the file
st.title("📂 Upload Your Excel File for eFocus Transformation")
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    file_bytes = uploaded_file.read()

    # Process the file through the eFocus Focus Grouping function
    transformed_file = efocus_focus(file_bytes)  # Currently just returning the uploaded file without processing

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
