import streamlit as st
import os
import openpyxl
from datetime import datetime
import pandas as pd
import zipfile
import io

st.set_page_config(page_title="Team Suggestions Portal", page_icon="💡", layout="centered")

UPLOAD_DIR = "uploaded_attachments"
EXCEL_FILE = "Subordinate_Suggestions_Log.xlsx"
ADMIN_PIN = "1234"  # You can change this to your preferred PIN

os.makedirs(UPLOAD_DIR, exist_ok=True)

# Initialize Excel log if it does not exist
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Suggestions"
    ws.append(["Timestamp", "Employee Name", "Department", "Suggestion Text", "Attachment File"])
    wb.save(EXCEL_FILE)

# --- SIDEBAR: ADMIN SECTION ---
st.sidebar.header("🔒 Plant Head / Admin Access")
admin_pin_input = st.sidebar.text_input("Enter Admin PIN", type="password")

if admin_pin_input == ADMIN_PIN:
    st.sidebar.success("Access Granted")
    st.sidebar.subheader("Download Records")
    
    # Download Excel Log
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.sidebar.download_button(
                label="📥 Download Excel Log",
                data=f,
                file_name=f"Suggestions_Log_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    # Download all photos/attachments as a single ZIP archive
    attachment_files = [os.path.join(UPLOAD_DIR, f) for f in os.listdir(UPLOAD_DIR) if os.path.isfile(os.path.join(UPLOAD_DIR, f))]
    if attachment_files:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for file in attachment_files:
                zip_file.write(file, arcname=os.path.basename(file))
        zip_buffer.seek(0)
        
        st.sidebar.download_button(
            label="📁 Download All Attachments (.zip)",
            data=zip_buffer,
            file_name=f"Plant_Attachments_{datetime.now().strftime('%Y%m%d')}.zip",
            mime="application/zip"
        )
    else:
        st.sidebar.info("No attachments uploaded yet.")

# --- MAIN FORM (EMPLOYEE VIEW) ---
st.title("💡 Suggestion & Feedback Box")
st.write("Share your ideas, operational bottlenecks, or improvements directly with the management.")

with st.form("suggestion_form", clear_on_submit=True):
    name = st.text_input("Your Name *", placeholder="Enter full name")
    dept = st.selectbox("Department *", [
        "Production & Processing",
        "Quality Control (QC)",
        "Maintenance & Utilities",
        "Supply Chain & Logistics",
        "Administration & Safety",
        "Other"
    ])
    suggestion = st.text_area("Suggestion / Feedback Details *", placeholder="Explain your suggestion or issue clearly...", height=150)
    
    uploaded_file = st.file_uploader(
        "Attach supporting file (Photo, Excel, PDF, or Text)", 
        type=["jpg", "jpeg", "png", "xlsx", "xls", "pdf", "txt", "csv"]
    )
    
    submitted = st.form_submit_button("Submit Suggestion")

    if submitted:
        if not name.strip() or not suggestion.strip():
            st.error("Please enter your name and suggestion before submitting.")
        else:
            saved_filename = "No Attachment"
            
            if uploaded_file is not None:
                timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                clean_name = f"{timestamp_str}_{uploaded_file.name}"
                file_path = os.path.join(UPLOAD_DIR, clean_name)
                
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                saved_filename = clean_name

            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb["Suggestions"]
            ws.append([
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                name,
                dept,
                suggestion,
                saved_filename
            ])
            wb.save(EXCEL_FILE)

            st.success("✅ Thank you! Your suggestion and attachment have been received.")

# --- ADMIN DISPLAY TABLE ---
if admin_pin_input == ADMIN_PIN:
    st.divider()
    st.subheader("📋 Submissions Log (Admin View)")
    try:
        df = pd.read_excel(EXCEL_FILE)
        st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.write("Unable to read data log.")
