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
            st.session_state.step = 3  # Proceed to the next step

# Step 3: Other operations (e.g., continue to next section or save data)
elif st.session_state.step == 3:
    st.subheader("Other Operations or Final Step")
    # Add your logic for further steps here...
