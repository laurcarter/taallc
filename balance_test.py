import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# Function to replicate the "BalanceFocusGrouping" operation
def balance_focus_grouping(file_bytes):
    # Load the workbook from the uploaded file
    wb = load_workbook(filename=BytesIO(file_bytes))
    ws = wb.active  # Assuming we're working with the active sheet

    # Add a new worksheet for the "Focus" sheet
    focus_ws = wb.create_sheet("Focus")

    # Find the last row and column of the data
    last_row = ws.max_row
    last_col = ws.max_column

    # Copy the data from the original sheet and paste it into the "Focus" sheet
    for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=last_col):
        for col_idx, cell in enumerate(row, 1):
            focus_ws.cell(row=row[0].row, column=col_idx, value=cell.value)

    # Split values in Column A by the parenthesis and move the second part to Column B
    for row in range(1, last_row + 1):
        cell_value = focus_ws.cell(row=row, column=1).value
        if isinstance(cell_value, str) and "(" in cell_value:
            split_values = cell_value.split("(")
            focus_ws.cell(row=row, column=1, value=split_values[0].strip())  # First part
            focus_ws.cell(row=row, column=2, value=split_values[1].strip(")"))  # Second part without closing parenthesis

    # Save the modified workbook to a BytesIO object
    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)

    return output_stream


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
