import streamlit as st
import os
from file_converter import convert_file

TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

st.title("Any File Converter with Format Selection")

uploaded_file = st.file_uploader("Upload a file")

if uploaded_file is not None:
    # Supported output formats (extend if you add more)
    output_formats = [".pdf"]
    output_format = st.selectbox("Select output file format", options=output_formats)

    input_path = os.path.join(TEMP_DIR, uploaded_file.name)
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success("Uploaded file saved temporarily.")

    base_name, _ = os.path.splitext(uploaded_file.name)
    output_file_name = base_name + output_format
    output_path = os.path.join(TEMP_DIR, output_file_name)

    success = False
    try:
        success = convert_file(input_path, output_path)
    except Exception as e:
        st.error(f"Conversion failed: {e}")

    if success and os.path.exists(output_path):
        st.success(f"File converted successfully to {output_format}!")
        with open(output_path, "rb") as f:
            mime_type = "application/pdf" if output_format == ".pdf" else "application/octet-stream"
            st.download_button(
                label=f"Download {output_file_name}",
                data=f,
                file_name=output_file_name,
                mime=mime_type
            )
    else:
        st.error("Conversion failed or unsupported file type.")
