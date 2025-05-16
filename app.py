import streamlit as st
import os
from file_converter import convert_file  # Make sure this matches your file_converter.py location

# Get the Desktop path of the current user
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
os.makedirs(desktop_path, exist_ok=True)  # Ensure Desktop folder exists

st.title("File to PDF Converter - Save Files to Desktop")

uploaded_file = st.file_uploader("Upload a file")

if uploaded_file is not None:
    # Save uploaded file to Desktop
    input_path = os.path.join(desktop_path, uploaded_file.name)
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success(f"Uploaded file saved at: {input_path}")

    # Define output PDF file path on Desktop
    base_name, _ = os.path.splitext(uploaded_file.name)
    output_path = os.path.join(desktop_path, base_name + ".pdf")

    try:
        # Convert the file using your conversion function
        convert_file(input_path, output_path)
        st.success(f"Converted file saved at: {output_path}")

        # Provide a download button for convenience
        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Converted PDF",
                data=f,
                file_name=os.path.basename(output_path),
                mime="application/pdf"
            )
    except Exception as e:
        st.error(f"Conversion failed: {e}")
