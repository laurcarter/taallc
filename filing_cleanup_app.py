import streamlit as st

# ---------- Streamlit App Flow ----------
st.set_page_config(page_title="Personal Information Form", layout="centered")
st.title("📋 Personal Information Form")

# Initialize session state variables
if "step" not in st.session_state:
    st.session_state.step = 1

if "personal_info" not in st.session_state:
    st.session_state.personal_info = {}

if "phone_numbers" not in st.session_state:
    st.session_state.phone_numbers = []

# Step 1: Tell us about yourself
if st.session_state.step == 1:
    st.subheader("To get started, we'll need some information about you.")

    # Form for Personal Information
    with st.form(key="personal_info_form"):
        first_name = st.text_input("First Name:")
        middle_initial = st.text_input("Middle Initial:")
        last_name = st.text_input("Last Name:")
        suffix = st.selectbox("Jr., Sr., III:", ["", "Jr.", "Sr.", "III"])
        occupation = st.text_input("Occupation:")
        employer = st.text_input("Employer:")
        date_of_birth = st.date_input("Date of Birth (mm/dd/yyyy):")

        # Submit button
        submit_button = st.form_submit_button("Continue")
    
    # Save the data
    if submit_button:
        st.session_state.personal_info = {
            "First Name": first_name,
            "Middle Initial": middle_initial,
            "Last Name": last_name,
            "Jr., Sr., III": suffix,
            "Occupation": occupation,
            "Employer": employer,
            "Date of Birth": date_of_birth
        }
        st.session_state.step = 2  # Move to Step 2 (Personal Information Summary)

# Step 2: Personal Information Summary
elif st.session_state.step == 2:
    st.subheader("Personal Information Summary")
    st.write("Here's the information you've provided:")

    # Display the personal information
    for key, value in st.session_state.personal_info.items():
        st.write(f"{key}: {value}")

    # Phone numbers management
    st.write("Listed below are your account phone number(s). You can add up to three phone numbers. We'll only use them to help you access your account.")

    # Display phone numbers and options to edit or delete
    phone_number_input = st.text_input("Phone Number:")

    if phone_number_input:
        if len(st.session_state.phone_numbers) < 3:
            st.session_state.phone_numbers.append(phone_number_input)
            st.experimental_rerun()

    # Display phone numbers and manage them
    for idx, phone in enumerate(st.session_state.phone_numbers):
        col1, col2 = st.columns([4, 1])
        with col1:
            st.write(phone)
        with col2:
            if st.button(f"Delete {phone}", key=f"delete_{idx}"):
                st.session_state.phone_numbers.remove(phone)
                st.experimental_rerun()

    # Option to go back or proceed
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("Edit Personal Info"):
            st.session_state.step = 1  # Go back to Step 1 to edit personal information
    with col2:
        if st.button("Continue"):
            st.session_state.step = 3  # Proceed to the next step (File Selection)

# Step 3: Select File to Auto File
elif st.session_state.step == 3:
    st.subheader("Select the file you wish to auto file today")
    
    # File uploader for P&L file selection
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    
    if uploaded_file:
        file_bytes = uploaded_file.read()
        st.session_state.excel_bytes = file_bytes  # Save the file for processing
        st.success("File uploaded successfully!")
        
        # Process the file (e.g., by calling the PNL transformation logic)
        from pnl_macro_translation import run_full_pl_macro  # Assuming this function is already defined
        processed_file = run_full_pl_macro(file_bytes)
        st.session_state.excel_bytes = processed_file

        st.success("File processed successfully!")

        # Proceed to download or other steps after processing
        st.session_state.step = 4  # Move to next step (e.g., download or confirmation)

# Step 4: Download Final Output or Other Options
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
