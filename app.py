import streamlit as st
import os
from file_converter import convert_file

# Configuration
TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

# Streamlit App
st.title("Universal File Converter")

uploaded_file = st.file_uploader("Upload your file", type=[
    'txt', 'docx', 'pptx', 'xlsx', 'jpg', 'jpeg', 'png'
])

if uploaded_file:
    # Get file info
    input_ext = os.path.splitext(uploaded_file.name)[1].lower()
    base_name = os.path.splitext(uploaded_file.name)[0]

    # Define conversion options
    conversion_options = {
        '.txt': ['PDF', 'DOCX', 'PNG'],
        '.docx': ['TXT', 'PDF', 'PNG', 'XLSX'],
        '.pptx': ['PDF', 'ZIP'],
        '.xlsx': ['PDF', 'DOCX', 'TXT', 'PNG'],
    }

    # Show format selector
    output_format = st.selectbox("Convert to:", conversion_options.get(input_ext, []))

    # Prepare file paths
    input_path = os.path.join(TEMP_DIR, uploaded_file.name)
    output_ext = f".{output_format.lower()}"
    output_file_name = f"{base_name}_converted{output_ext}"
    output_path = os.path.join(TEMP_DIR, output_file_name)

    # Save uploaded file
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Conversion and download
    if st.button("Convert Now"):
        if convert_file(input_path, output_path):
            st.success("✅ Conversion Successful!")
            
            # Show download button
            mime_types = {
                '.pdf': 'application/pdf',
                '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                '.png': 'image/png',
                '.txt': 'text/plain',
                '.zip': 'application/zip'
            }
            
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Converted File",
                    data=f,
                    file_name=output_file_name,
                    mime=mime_types.get(output_ext, 'application/octet-stream')
                )
        else:
            st.error("❌ Conversion Failed. Please check the file format.")
