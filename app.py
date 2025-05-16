import streamlit as st
import os
from file_converter import convert_file

TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

st.title("Universal File Converter")

uploaded_file = st.file_uploader("Upload a file")

if uploaded_file is not None:
    supported_formats = ['.pdf', '.docx', '.txt']  # Extend this list as you add conversions
    uploaded_ext = os.path.splitext(uploaded_file.name)[1].lower()
    output_options = [ext for ext in supported_formats if ext != uploaded_ext]

    if not output_options:
        st.warning("No other output formats available for this file type.")
    else:
        output_format = st.selectbox("Select output format", output_options)

        input_path = os.path.join(TEMP_DIR, uploaded_file.name)
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success("Uploaded file saved temporarily.")

        base_name, _ = os.path.splitext(uploaded_file.name)
        output_file_name = base_name + output_format
        output_path = os.path.join(TEMP_DIR, output_file_name)

        success = convert_file(input_path, output_path)

        if success and os.path.exists(output_path):
            st.success(f"Converted to {output_format} successfully!")
            mime_types = {
                '.pdf': 'application/pdf',
                '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                '.txt': 'text/plain',
            }
            mime_type = mime_types.get(output_format, 'application/octet-stream')
            with open(output_path, 'rb') as f:
                st.download_button(
                    label=f"Download {output_file_name}",
                    data=f,
                    file_name=output_file_name,
                    mime=mime_type
                )
        else:
            st.error("Conversion failed or unsupported file type.")
