import streamlit as st
import os
import openpyxl
from datetime import datetime

# Configure page for mobile screens
st.set_page_config(page_title="Team Suggestions Portal", page_icon="💡", layout="centered")

# Folders and tracking files
UPLOAD_DIR = "uploaded_attachments"
EXCEL_FILE = "Subordinate_Suggestions_Log.xlsx"

os.makedirs(UPLOAD_DIR, exist_ok=True)

# Initialize Excel log if it does not exist
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Suggestions"
    ws.append(["Timestamp", "Employee Name", "Department", "Suggestion Text", "Attachment File"])
    wb.save(EXCEL_FILE)

# --- USER INTERFACE ---
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
    
    # File uploader supporting photos, documents, and spreadsheets
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
            
            # Save uploaded file to disk
            if uploaded_file is not None:
                # Add timestamp to filename to prevent overwriting files with identical names
                timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                clean_name = f"{timestamp_str}_{uploaded_file.name}"
                file_path = os.path.join(UPLOAD_DIR, clean_name)
                
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                saved_filename = clean_name

            # Append entry to the Excel tracker
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