import streamlit as st
from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import re

# ---------- Utility Functions ----------
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def highlight_and_flag_totals(file_bytes):
    wb = load_workbook(filename=BytesIO(file_bytes), data_only=True)
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
                    cell.value = remove_parentheses_content(str(cell.value))
                    cell.fill = PatternFill()  # Reset to default empty fill

    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream

def perform_pnl_transformation(file_bytes):
    from pnl_macro_translation import run_full_pl_macro
    return run_full_pl_macro(file_bytes)

# ---------- Streamlit App Flow ----------
st.set_page_config(page_title="Filing Cleanup Wizard", layout="centered")
st.title("🧾 Filing Cleanup Wizard")

if "step" not in st.session_state:
    st.session_state.step = 1
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "flagged_cells" not in st.session_state:
    st.session_state.flagged_cells = []

# Step 1: Upload Excel
if st.session_state.step == 1:
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file:
        file_bytes = uploaded_file.read()
        highlighted_file, flagged = highlight_and_flag_totals(file_bytes)
        st.session_state.excel_bytes = highlighted_file
        st.session_state.flagged_cells = flagged

        st.success(f"Found {len(flagged)} potentially incorrect 'Total' cells.")
        if st.button("Next"):
            st.session_state.step = 2

# Step 2: Show flagged cells
elif st.session_state.step == 2:
    st.subheader("🔍 Review Flagged Cells")
    if st.session_state.flagged_cells:
        for sheet, coord, val in st.session_state.flagged_cells:
            st.write(f"- **{sheet}**!{coord} → `{val}`")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Yes, clean these cells"):
                cleaned_file = clean_flagged_totals(st.session_state.excel_bytes)
                st.session_state.excel_bytes = cleaned_file
                st.session_state.step = 3
        with col2:
            if st.button("No, leave them as-is"):
                st.session_state.step = 3

    else:
        st.info("No problematic 'Total' cells found. Skipping ahead.")
        if st.button("Continue"):
            st.session_state.step = 3

# Step 3: Choose transformation
elif st.session_state.step == 3:
    st.subheader("🔧 What type of filing is this?")
    choice = st.radio("Select your filing type:", ["Profit & Loss (P&L)", "Balance Sheet"], index=0)
    if st.button("Run Transformation"):
        if choice == "Profit & Loss (P&L)":
            st.session_state.excel_bytes = perform_pnl_transformation(st.session_state.excel_bytes)
        st.session_state.step = 4

# Step 4: Download Final Output
elif st.session_state.step == 4:
    st.subheader("✅ Final Step: Download Processed File")
    st.download_button(
        label="Download Final Excel",
        data=st.session_state.excel_bytes,
        file_name="final_filing.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    if st.button("Start Over"):
        for key in ["step", "excel_bytes", "flagged_cells"]:
            st.session_state.pop(key, None)
