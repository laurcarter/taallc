import openpyxl
from openpyxl import load_workbook
from io import BytesIO
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

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

    # Step 1: Copy Column A from the original sheet into Focus sheet (Column A in Focus)
    for row in range(1, max_row + 1):
        focus_ws.cell(row=row, column=1).value = ws.cell(row=row, column=1).value

    # Step 2: Split Column A by the opening parenthesis and move to Column B and D
    for row in range(1, max_row + 1):
        val = focus_ws.cell(row=row, column=1).value
        if val and '(' in str(val):
            # Split the value at the opening parenthesis and store the parts
            parts = str(val).split('(', 1)
            focus_ws.cell(row=row, column=1).value = parts[0].strip()  # First part goes to Column A
            focus_ws.cell(row=row, column=2).value = parts[1].strip()  # Second part goes to Column D

    # Step 3: Remove parentheses in Column D (Focus sheet)
    for row in range(1, max_row + 1):
        val_d = focus_ws.cell(row=row, column=2).value
        if val_d:
            focus_ws.cell(row=row, column=2).value = str(val_d).replace("(", "").replace(")", "")
    # Step 4: Copy Column B from the original sheet to Column D in Focus sheet (as value only)
    for row in range(1, max_row + 1):
        # Get the calculated value (not the formula) from Column B in the original sheet
        original_value = ws.cell(row=row, column=2).value  # Get the value, not the formula
        focus_ws.cell(row=row, column=4).value = original_value  # Directly paste the value into Column D

    # Insert two columns at the beginning of the Focus sheet (Columns A and B become empty)
    focus_ws.insert_cols(1, 2)
    
    # Shift data from columns starting from row 5 (move the data from columns C, D, F)
    for row in range(5, max_row + 1):
        # Move data in Column C to Column E
        focus_ws.cell(row=row, column=5).value = focus_ws.cell(row=row, column=3).value
        focus_ws.cell(row=row, column=3).value = None  # Clear original cell
    
        # Move data in Column D to Column C
        focus_ws.cell(row=row, column=3).value = focus_ws.cell(row=row, column=4).value
        focus_ws.cell(row=row, column=4).value = None  # Clear original cell
    
        # Move data in Column F to Column D
        focus_ws.cell(row=row, column=4).value = focus_ws.cell(row=row, column=6).value
        focus_ws.cell(row=row, column=6).value = None  # Clear original cell
    
    # Add column titles in row 4 (ensure this is done after the rows are shifted)
    focus_ws["C4"] = "Focus"
    focus_ws["D4"] = "Amount"
    focus_ws["E4"] = "Description"
    focus_ws["F4"] = "Totals"
    
    # Shift everything below row 4 down by 3 rows in Focus sheet
    focus_ws.insert_rows(4, amount=3)  # Insert 3 rows at row 4 in focus_ws
    
    # Define fill color (black) and font color (white)
    black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    white_font = Font(color="FFFFFF")
    
    # Fill columns C to F in row 7 with black and change the text color to white in the Focus sheet
    for col in ["C", "D", "E", "F"]:
        focus_ws[f"{col}7"].fill = black_fill
        focus_ws[f"{col}7"].font = white_font
    
    # Format column D and F in the Focus sheet to show numbers with thousand commas
    for row in range(1, max_row + 1):
        focus_ws.cell(row=row, column=4).number_format = "#,##0"  # Column D
        focus_ws.cell(row=row, column=6).number_format = "#,##0"  # Column F
    
    # Increase the width of column E to double the default width in the Focus sheet
    focus_ws.column_dimensions["E"].width = focus_ws.column_dimensions["E"].width * 2.5

    

    #subtotals
    #apply_subtotals_for_sheet(focus_ws, max_row)
    
    # Call this function for both sheets
    #delete_blank_rows(focus_ws, max_row)  # For Focus sheet
    
    

    # You can now use income_sum and expense_sum in your further calculations
    #apply_income_expense_totals(focus_ws, max_row)


    #create_summary(focus_ws, max_row)
    #apply_focus_summary_formatting(focus_ws, max_row)


    #apply_random_formatting(focus_ws, max_row)


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
