import openpyxl
from openpyxl import load_workbook
from io import BytesIO
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
import streamlit as st


def balance_focus_grouping(file_bytes):
    # Ensure file_bytes is a BytesIO object
    if not isinstance(file_bytes, BytesIO):
        # If file_bytes is raw data (not a BytesIO object), wrap it into BytesIO
        file_bytes = BytesIO(file_bytes)

    # Load the workbook from the BytesIO object
    wb = load_workbook(filename=file_bytes)
    ws = wb.active  # Get the active worksheet

    # Create Focus worksheet
    focus_ws = wb.create_sheet(title="Focus")

    # Find the last used row and column in the original sheet
    max_row = ws.max_row
    max_col = ws.max_column

    # Step 1: Copy the content from the original sheet and paste it as plain text into the Focus sheet
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            focus_ws.cell(row=cell.row, column=cell.column, value=cell.value)

    # Step 2: Delimit Column A by the opening parenthesis and put the second part in Column B in Focus sheet
    for row in range(1, max_row + 1):
        val = focus_ws.cell(row=row, column=1).value
        if val and '(' in str(val):
            # Split the value at the opening parenthesis and store the parts
            parts = str(val).split('(', 1)
            focus_ws.cell(row=row, column=1).value = parts[0].strip()
            focus_ws.cell(row=row, column=2).value = parts[1].strip()

    # Step 3: Remove parentheses in Column B (Focus sheet)
    for row in range(1, max_row + 1):
        val_b = focus_ws.cell(row=row, column=2).value
        if val_b:
            focus_ws.cell(row=row, column=2).value = str(val_b).replace("(", "").replace(")", "")

    # Step 4: Copy Column B from the original sheet to Column E in Focus sheet
    for row in range(1, max_row + 1):
        original_value = ws.cell(row=row, column=2).value
        focus_ws.cell(row=row, column=5).value = original_value

    # Step 5: Clear Columns C and D in Focus sheet
    for row in range(1, max_row + 1):
        focus_ws.cell(row=row, column=3).value = None
        focus_ws.cell(row=row, column=4).value = None

    # Step 6: Move Column E to Column D in Focus sheet
    for row in range(1, max_row + 1):
        focus_ws.cell(row=row, column=4).value = focus_ws.cell(row=row, column=5).value
        focus_ws.cell(row=row, column=5).value = None

    # Save the modified workbook to a BytesIO object
    output = BytesIO()
    wb.save(output)
    output.seek(0)  # Move cursor to the beginning of the BytesIO object
    return output.read()  # Return the transformed file as bytes



# Step 1: Upload the file
st.title("📂 Upload Your Excel File for Balance Sheet Transformation")
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    file_bytes = uploaded_file.read()

    # Process the file through the Balance Focus Grouping function
    transformed_file = balance_focus_grouping(file_bytes)

    # Store the transformed file in session state
    st.session_state.excel_bytes = transformed_file

    st.success("File processed. The 'Focus' sheet has been created and updated.")

    # Provide option to download the transformed file
    st.download_button(
        label="Download Transformed File",
        data=st.session_state.excel_bytes,
        file_name="balance_transformed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
