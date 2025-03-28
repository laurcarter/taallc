import streamlit as st
from io import BytesIO
from openpyxl import load_workbook

# Function to load the uploaded file and check the sheet contents
def load_excel(file_bytes):
    wb = load_workbook(BytesIO(file_bytes))
    sheet = wb.active
    return wb, sheet

# Step 7: After the first download, prompt the user for eFocus continuation
if "step" not in st.session_state:
    st.session_state.step = 1
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

# Step 1: File Upload and Worksheet Loading
if st.session_state.step == 1:
    st.title("📂 Upload Your Excel File for eFocus")  # Title for Step 1
    st.write("Upload your Excel file to continue to eFocus.")  # Description for Step 1

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file:
        st.session_state.excel_bytes = uploaded_file.read()
        st.success("File uploaded successfully!")
        # Proceed to Step 2 to review or process the file
        st.session_state.step = 2

# Step 2: Review or Process the Uploaded File
elif st.session_state.step == 2:
    st.title("📊 Review Your Uploaded Excel File")  # Title for Step 2
    st.write("Review the uploaded Excel file and choose to proceed with eFocus.")  # Description for Step 2

    # Load the uploaded file to preview its contents
    wb, sheet = load_excel(st.session_state.excel_bytes)
    
    # Display basic info about the sheet
    st.write(f"**Sheet Name:** {sheet.title}")
    st.write(f"**Number of Rows:** {sheet.max_row}")
    st.write(f"**Number of Columns:** {sheet.max_column}")
    
    # Preview the first few rows to give a sense of the data
    preview_rows = 5
    data_preview = []
    for row in sheet.iter_rows(min_row=1, max_row=preview_rows, values_only=True):
        data_preview.append(row)
    
    st.write("**Preview of the Data:**")
    st.write(data_preview)

    # Ask if the user wants to continue to eFocus
    continue_to_eFocus = st.button("Continue to eFocus")
    if continue_to_eFocus:
        st.session_state.step = 3  # Move to Step 3 to initiate eFocus

# Step 3: Initiate eFocus Process (Post-Download Actions)
elif st.session_state.step == 3:
    st.title("🔄 Initiating eFocus Process")  # Title for Step 3
    st.write("The eFocus process is now being initiated.")  # Description for Step 3

    # Here, you would add the functionality that kicks off the eFocus process
    # For now, let's just simulate with a success message:
    st.success("eFocus process started! You can now continue with eFocus operations.")
    
    # Provide an option to download the result
    st.download_button(
        label="Download Final eFocus Processed File",
        data=st.session_state.excel_bytes,
        file_name="eFocus_processed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Optional: Add a button to start over if the user wants to process a new file
    if st.button("Start Over"):
        for key in ["step", "excel_bytes"]:
            st.session_state.pop(key, None)
        st.session_state.step = 1  # Reset to Step 1

