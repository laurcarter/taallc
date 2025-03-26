import streamlit as st

# ---------- Utility Functions ----------
def perform_pnl_transformation(file_bytes):
    from pnl_macro_translation import run_full_pl_macro
    return run_full_pl_macro(file_bytes)

# ---------- Streamlit App Flow ----------
st.set_page_config(page_title="Filing Cleanup Wizard", layout="centered")
st.title("🧾 Filing Cleanup Wizard")

# New Step: Collect user information
if "step" not in st.session_state:
    st.session_state.step = 1

if "user_info" not in st.session_state:
    st.session_state.user_info = {}

# Step 1: Tell us about yourself (User Information Form)
if st.session_state.step == 1:
    st.subheader("👤 To get started, we'll need some information about you.")
    
    with st.form(key='user_info_form'):
        first_name = st.text_input("First Name:")
        middle_initial = st.text_input("Middle Initial:")
        last_name = st.text_input("Last Name:")
        jr_sr_iii = st.text_input("Jr., Sr., III:")
        occupation = st.text_input("Occupation:")
        employer = st.text_input("Employer:")
        date_of_birth = st.date_input("Date of Birth (mm/dd/yyyy)")

        submit_button = st.form_submit_button(label="Next")

    if submit_button:
        # Save user information into session state
        st.session_state.user_info = {
            "first_name": first_name,
            "middle_initial": middle_initial,
            "last_name": last_name,
            "jr_sr_iii": jr_sr_iii,
            "occupation": occupation,
            "employer": employer,
            "date_of_birth": date_of_birth
        }

        st.session_state.step = 2

# Step 2: Show uploaded file and proceed to the next steps
elif st.session_state.step == 2:
    st.subheader("🔼 Tell us about your file!")
    
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file:
        file_bytes = uploaded_file.read()
        st.session_state.excel_bytes = file_bytes

        if st.button("Next"):
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
        for key in ["step", "excel_bytes", "user_info"]:
            st.session_state.pop(key, None)


