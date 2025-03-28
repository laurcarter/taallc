import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import pandas as pd

# ---------- eFocus Function Placeholder ----------
def efocus_focus(file_bytes, client_data_bytes):
    # Load the uploaded Excel file
    wb = load_workbook(filename=BytesIO(file_bytes))
    focus_ws = wb['Focus']  # Assuming 'Focus' sheet exists in the uploaded workbook

    # Load the client data from the uploaded client_data.xlsx
    client_data = pd.read_excel(BytesIO(client_data_bytes))  # Load client data as a DataFrame

    # Print the columns to inspect
    print("Client Data Columns:", client_data.columns)

    # Get the client names from row 1 (starting from column C onwards)
    client_names = []
    for col in range(3, focus_ws.max_column + 1):  # Starting from column C (column index 3)
        client_name = focus_ws.cell(row=1, column=col).value
        if client_name:  # If the cell is not empty
            client_names.append(client_name)

    # Assume you get the client name from the 'Focus' sheet for which the transformation should apply
    for client_name in client_names:
        # Check if the client name exists in the client data
        if 'ClientName' in client_data.columns:
            client_row = client_data[client_data['ClientName'] == client_name]

            if not client_row.empty:
                # Apply client-specific transformations using the data
                # For example, update a specific cell in the 'Focus' sheet with client-specific info
                focus_ws.cell(row=2, column=2).value = client_row.iloc[0]['SpecificColumn']  # Example update
            else:
                st.error(f"No client data found for {client_name}.")
        else:
            st.error("'ClientName' column not found in the client data.")

    # Save the updated workbook to BytesIO and return it
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output  # Return the updated file as BytesIO



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
