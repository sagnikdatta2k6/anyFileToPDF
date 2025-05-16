import streamlit as st
import os
from file_converter import convert_file  # Adjust if your file_converter.py is elsewhere

# Directory to save uploaded and converted files
SAVE_DIR = "saved_files"
os.makedirs(SAVE_DIR, exist_ok=True)

st.title("Any File to PDF Converter with Local Storage")

uploaded_file = st.file_uploader("Upload a file")

if uploaded_file is not None:
    # Save uploaded file locally
    input_path = os.path.join(SAVE_DIR, uploaded_file.name)
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    input_path = os.path.abspath(input_path)
    st.success(f"Uploaded file saved at: {input_path}")

    # Define output file path with .pdf extension (or keep original extension if needed)
    base_name, _ = os.path.splitext(uploaded_file.name)
    output_path = os.path.join(SAVE_DIR, base_name + ".pdf")
    output_path = os.path.abspath(output_path)

    # Convert the file
    try:
        convert_file(input_path, output_path)

        # Check if output file was created
        if os.path.exists(output_path):
            st.success(f"File converted and saved at: {output_path}")

            # Provide download button
            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download Converted PDF",
                    data=f,
                    file_name=os.path.basename(output_path),
                    mime="application/pdf"
                )
        else:
            st.error(f"Conversion failed: Output file not found at {output_path}")
    except Exception as e:
        st.error(f"Conversion failed with error: {e}")
