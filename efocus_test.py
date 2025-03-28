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

def efocus_focus(file_bytes, client_data_bytes):
    # Load the Focus sheet from the uploaded file (file_bytes)
    wb = load_workbook(filename=BytesIO(file_bytes))
    focus_ws = wb['Focus']  # Assuming the Focus sheet is already available

    # Load the client data from the second uploaded file (client_data_bytes)
    client_data = pd.read_excel(BytesIO(client_data_bytes), header=None)  # Reading client data without headers

    # Initialize list to hold client names that are valid (skipping blank columns)
    client_names = []

    # Loop through columns C, E, G, ..., until column 3 + 100 columns
    for col in range(3, 203, 2):  # 3 for column C, 203 is 3 + (100*2)
        cell_value = str(client_data.iloc[0, col - 1]).strip()  # Get client name from row 1 (adjusted for 0-indexing)
        
        # Only add to the list if the client name is non-empty and not just a header like 'Unnamed'
        if cell_value and 'Unnamed' not in cell_value:
            client_names.append(cell_value)

    # If no valid client names were found, show a message
    if not client_names:
        st.error("No valid client names found in the client data.")
        return None

    # Show the list of client names as select buttons in Streamlit
    client_name = st.selectbox("Choose a client name from the list:", client_names)
    
    # After client selection, return the column of the selected client name
    client_column = client_names.index(client_name) * 2 + 3  # Adjust column index for 0-indexed list
    
    # Now do something with the client column (For example: create FocusTarget sheet in the main workbook)
    # This is where you would process further depending on your logic
    
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
