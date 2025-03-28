import streamlit as st
from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font
import re
from collapse import collapse_sheet


# ---------- Utility Functions ----------
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def check_and_prompt_for_net_income(focus_ws):
    # Step 1: Iterate through rows starting from row 8 in Column A
    for row in range(8, focus_ws.max_row + 1):
        a_value = str(focus_ws.cell(row=row, column=1).value).strip()  # Value in Column A

        # Step 2: Check if "Net Income" is in the cell and if it doesn't have parentheses
        if "Net Income" in a_value and "(" not in a_value:
            # Display warning and a text box for the user to input the Focus box number
            net_income_input = st.text_input(f"Warning: 'Net Income' in row {row} is not coded with parentheses. Please enter the correct Focus box number:")

            if net_income_input:
                # Add parentheses around the user's input and update the cell in Column A
                new_value = f"Net Income ({net_income_input})"
                focus_ws.cell(row=row, column=1).value = new_value  # Update the value in Column A
                st.success(f"Net Income for row {row} has been updated with Focus box number: {net_income_input}")
                return True  # Indicate that the Net Income was updated

    return False  # If no updates were made, return False




# Function to highlight flagged totals
def highlight_and_flag_totals(file_bytes):
    # Ensure file_bytes is wrapped in BytesIO before passing to load_workbook
    if not isinstance(file_bytes, BytesIO):
        file_bytes = BytesIO(file_bytes)

    wb = load_workbook(filename=file_bytes, data_only=True)
    flagged_cells = []

    for ws in wb.worksheets:
        max_row = ws.max_row
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=8):
            for cell in row:
                value = str(cell.value) if cell.value else ""
                if "Total" in value and "(" in value and ")" in value:
                    cell.fill = yellow_fill
                    flagged_cells.append((ws.title, cell.coordinate, value))

    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream, flagged_cells


# Function to clean flagged totals (remove parentheses and 'Total')
def clean_flagged_totals(file_bytes):
    if isinstance(file_bytes, BytesIO):
        file_obj = file_bytes
    else:
        file_obj = BytesIO(file_bytes)

    wb = load_workbook(filename=file_obj, data_only=True)

    def remove_parentheses_content(text):
        return re.sub(r'\s*\([^)]*\)', '', text).strip()

    for ws in wb.worksheets:
        max_row = ws.max_row
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=8):
            for cell in row:
                if cell.value and "Total" in str(cell.value):
                    cell.value = remove_parentheses_content(str(cell.value))  # Clean text
                    cell.fill = PatternFill()  # Reset the highlight fill

    # Save the updated file and return as BytesIO
    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream



# Function to perform P&L transformation
def perform_pnl_transformation(file_bytes):
    from pnl_macro_translation import run_full_pl_macro
    return run_full_pl_macro(file_bytes)


# Function to perform Balance transformation
def perform_balance_transformation(file_bytes):
    from balance import balance_focus_grouping 
    return balance_focus_grouping(file_bytes)

# ---------- Streamlit App Flow ----------
st.set_page_config(page_title="Personal Information Collection", layout="wide")

# Step Handling
if "step" not in st.session_state:
    st.session_state.step = 1
if "personal_info" not in st.session_state:
    st.session_state.personal_info = {}

# Step 1: Personal Information Collection
if st.session_state.step == 1:
    st.title("🧾 Personal Information Collection")  # Title for Step 1
    st.write("To get started, we'll need some information about you.")  # Description for Step 1

    with st.expander("Step 1: Tell us about yourself", expanded=True):
        first_name = st.text_input("First Name")
        middle_initial = st.text_input("Middle Initial")
        last_name = st.text_input("Last Name")
        suffix = st.selectbox("Suffix (Jr., Sr., III)", ["", "Jr.", "Sr.", "III"])
        occupation = st.text_input("Occupation")
        employer = st.text_input("Employer")
        dob = st.date_input("Date of Birth")  # Streamlit will handle the date format automatically
        phone_number = st.text_input("Phone Number")

        if st.button("Continue"):
            # Store the collected information in session state
            st.session_state.personal_info = {
                "first_name": first_name,
                "middle_initial": middle_initial,
                "last_name": last_name,
                "suffix": suffix,
                "occupation": occupation,
                "employer": employer,
                "dob": dob,
                "phone_number": phone_number
            }
            st.session_state.step = 2

# Step 2: Personal Information Summary
elif st.session_state.step == 2:
    st.title("🔍 Review Your Personal Information")  # Title for Step 2
    st.write("Please review your personal information before continuing.")  # Description for Step 2

    with st.expander("Step 2: Review Your Information", expanded=True):
        st.write(f"**First Name:** {st.session_state.personal_info['first_name']}")
        st.write(f"**Middle Initial:** {st.session_state.personal_info['middle_initial']}")
        st.write(f"**Last Name:** {st.session_state.personal_info['last_name']}")
        st.write(f"**Suffix:** {st.session_state.personal_info['suffix']}")
        st.write(f"**Occupation:** {st.session_state.personal_info['occupation']}")
        st.write(f"**Employer:** {st.session_state.personal_info['employer']}")
        st.write(f"**Date of Birth:** {st.session_state.personal_info['dob'].strftime('%m/%d/%Y')}")
        st.write(f"**Phone Number:** {st.session_state.personal_info['phone_number']}")

        if st.button("Continue"):
            st.session_state.step = 3

# Step 3: Upload Excel File
elif st.session_state.step == 3:
    st.title("📂 Upload Your Excel File")  # Title for Step 3
    st.write("Please upload your Excel file for auto-filing.")  # Description for Step 3

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file:
        file_bytes = uploaded_file.read()

        # Load the workbook to check the number of sheets
        wb = load_workbook(filename=BytesIO(file_bytes))
        sheet_names = wb.sheetnames  # List of sheet names
        if len(sheet_names) > 1:
            # If multiple sheets are detected, ask the user to resubmit a new file with a single sheet
            st.error("Your file contains multiple sheets. Please resubmit a new file with only one sheet.")
            st.session_state.step = 3  # Stay on Step 3 so they can upload again
        else:
            # If there is only one sheet, proceed with the regular logic
            sheet = wb.active

            # Now highlight and flag totals on the selected sheet
            highlighted_file, flagged = highlight_and_flag_totals(file_bytes)
            st.session_state.excel_bytes = highlighted_file  # Store the highlighted file in session state
            st.session_state.flagged_cells = flagged  # Store the flagged cells
    
            st.success(f"Found {len(flagged)} potentially incorrect 'Total' cells.")
    
            # Continue button to move to the next step
            if st.button("Continue"):
                    st.session_state.step = 4  # Move to Step 4





# Step 4: Show flagged cells for review and clean if needed
elif st.session_state.step == 4:
    st.title("🔍 Review Flagged Cells")  # Title for Step 4
    st.write("Review and clean any flagged cells, or leave them as-is.")  # Description for Step 4

    # Ensure flagged_cells exist in session_state
    if 'flagged_cells' not in st.session_state:
        st.session_state.flagged_cells = []

    if st.session_state.flagged_cells:
        # Display all the flagged cells with their coordinates
        for sheet, coord, val in st.session_state.flagged_cells:
            st.write(f"- **{sheet}**!{coord} → `{val}`")

        col1, col2 = st.columns(2)

        # Button for cleaning flagged totals
        with col1:
            if st.button("Yes, clean these cells"):
                # Get the current file from session state
                file_bytes = st.session_state.excel_bytes
                
                # Clean the flagged totals (this function will modify the file)
                cleaned_file = clean_flagged_totals(file_bytes)

                # Update the session state with the cleaned file
                st.session_state.excel_bytes = cleaned_file  # Store the cleaned file in session state

                # Proceed to Step 5 after cleaning
                st.session_state.step = 5  

        # Button for leaving the flagged totals as-is (just keep them highlighted)
        with col2:
            if st.button("No, leave them as-is"):
                # Ensure we keep the file as-is with highlighted totals, no cleaning
                file_bytes = st.session_state.excel_bytes
                
                # Just highlight the totals again and move to Step 5
                highlighted_file, _ = highlight_and_flag_totals(file_bytes)  # Re-highlight if needed
                st.session_state.excel_bytes = highlighted_file  # Store the highlighted file in session state

                # Proceed to Step 5 without cleaning
                st.session_state.step = 5  

    else:
        # If no flagged cells, display a message
        st.info("No problematic 'Total' cells found. Skipping ahead.")
        if st.button("Continue"):
            # Proceed to Step 5 if no flagged cells
            st.session_state.step = 5  # Skip to Step 5 if no flagged cells



# Step 5: Choose Transformation Type
elif st.session_state.step == 5:
    # **Silent Collapse Check**
    file_bytes = st.session_state.excel_bytes  # The current file in session state

    # Load the workbook
    wb = load_workbook(file_bytes)  # Pass file_bytes as a valid file-like object
    sheet = wb.active

    # Check for empty cells in column A (rows 5-10) and collapse if necessary
    empty_cell_count = 0
    total_cells_to_check = 6  # Checking rows 5 to 10 (6 rows total)
    for row in range(5, 11):  # Rows 5 to 10 in column A
        cell = sheet.cell(row=row, column=1)
        if cell.value is None or str(cell.value).strip() == "":
            empty_cell_count += 1
    if empty_cell_count == total_cells_to_check:
        collapsed_file = collapse_sheet(file_bytes)
        if isinstance(collapsed_file, BytesIO):  # Ensure it returns BytesIO
            st.session_state.excel_bytes = collapsed_file
        else:
            st.error("Error: Collapse function did not return a valid file.")
            
    st.title("🔧 What type of filing is this?")
    st.write("Select the type of filing for this document.")

    choice = st.radio("Select your filing type:", ["Profit & Loss (P&L)", "Balance Sheet"], index=0)

    if st.button("Run Transformation"):
        if choice == "Profit & Loss (P&L)":
            st.session_state.excel_bytes = perform_pnl_transformation(st.session_state.excel_bytes)
        
        elif choice == "Balance Sheet":
            # Before transforming balance sheet, check and possibly update "Net Income"
            wb = load_workbook(st.session_state.excel_bytes)
            focus_ws = wb.active  # Assuming the relevant sheet is active; adjust if necessary
            
            # Check and update the "Net Income" if necessary
            net_income_updated = check_and_prompt_for_net_income(focus_ws)
            
            if net_income_updated:
                # If Net Income was updated, proceed with balance sheet transformation
                st.session_state.excel_bytes = perform_balance_transformation(st.session_state.excel_bytes)
            else:
                # If no update was made, proceed with balance sheet transformation anyway
                st.session_state.excel_bytes = perform_balance_transformation(st.session_state.excel_bytes)

        st.session_state.step = 6  # Move to the final step for download


            
# Step 6: Download Final Processed File
elif st.session_state.step == 6:
    st.title("✅ Final Step: Download Processed File")  # Title for Step 6
    st.write("Download the final processed file.")  # Description for Step 6

    st.download_button(
        label="Download Final Excel",
        data=st.session_state.excel_bytes,
        file_name="final_filing.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    if st.button("Start Over"):
        for key in ["step", "excel_bytes", "flagged_cells"]:
            st.session_state.pop(key, None)

