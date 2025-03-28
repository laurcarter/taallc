import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import pandas as pd
import streamlit as st
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook

def efocus_focus(file_bytes, client_data_bytes):
    # Load the Focus sheet from the uploaded file (file_bytes)
    wb = load_workbook(filename=BytesIO(file_bytes))
    focus_ws = wb['Focus']  # Assuming the Focus sheet is already available

    # Load the client data from the second uploaded file (client_data_bytes)
    client_data = pd.read_excel(BytesIO(client_data_bytes), header=None)  # Reading client data without headers

    # Check the client data structure
    st.write("Client Data Loaded", client_data.head())  # Display the first few rows of the client data for debugging

    # Initialize list to hold valid client names
    client_names = []

    # Loop through columns starting from C (column 3), skipping alternate columns (C, E, G, etc.)
    for col in range(2, client_data.shape[1], 2):  # Starting from column 2 (C in 0-indexed), increment by 2 (skip alternating)
        cell_value = str(client_data.iloc[0, col]).strip()  # Get client name from row 1 (adjusted for 0-indexing)
        
        # Only add to the list if the cell contains a valid client name (not empty or "Unnamed")
        if cell_value and 'Unnamed' not in cell_value:
            client_names.append(cell_value)

    # If no valid client names were found, show a message and exit
    if not client_names:
        st.error("No valid client names found in the client data.")
        return None

    # Display the valid client names as a list of buttons for selection
    client_name = st.selectbox("Choose a client name from the list:", client_names)
    
    # After client selection, return the column of the selected client name
    client_column = client_names.index(client_name) * 2 + 3  # Adjust column index for 0-indexed list
    
    # Now, you can do something with the client column (e.g., use it to paste data into the FocusTarget sheet)
    
    # For debugging purposes, we'll return the selected column for now.
    st.write(f"Selected Client: {client_name}")
    st.write(f"Client Column: {client_column}")  # This is the column index in the client data
    
    return client_column  # You can use this column index to fetch the data for pasting into FocusTarget




# Streamlit UI for the file upload and processing
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
    else:
        st.error("Error processing the files. Please check the files and try again.")

else:
    st.info("Please upload both the Focus and Client Data files to proceed.")
