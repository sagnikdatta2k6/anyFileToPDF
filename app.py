import streamlit as st
import os
from file_converter import convert_file  # Adjust if needed

# Temporary folder to save files during the session
TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

st.title("File to PDF Converter with Manual Download")

uploaded_file = st.file_uploader("Upload a file")

if uploaded_file is not None:
    # Save uploaded file temporarily
    input_path = os.path.join(TEMP_DIR, uploaded_file.name)
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success(f"Uploaded file saved temporarily.")

    base_name, _ = os.path.splitext(uploaded_file.name)
    output_path = os.path.join(TEMP_DIR, base_name + ".pdf")

    try:
        # Convert the file
        convert_file(input_path, output_path)
        st.success("File converted successfully!")

        # Provide download button for the converted file
        with open(output_path, "rb") as f:
            st.download_button(
                label="Click here to download the converted PDF",
                data=f,
                file_name=os.path.basename(output_path),
                mime="application/pdf"
            )
    except Exception as e:
        st.error(f"Conversion failed: {e}")
