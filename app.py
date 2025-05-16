import streamlit as st
import os
from file_converter import convert_file

TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

st.title("Custom File Converter")

uploaded_file = st.file_uploader("Upload a file")

if uploaded_file is not None:
    # Supported output formats for each input type
    conversion_options = {
        '.txt': ['.pdf', '.docx', '.png'],
        '.docx': ['.txt', '.pdf', '.png', '.xlsx'],
        '.pptx': ['.pdf', '.png'],
        '.xlsx': ['.pdf', '.docx', '.txt', '.png'],
    }
    input_ext = os.path.splitext(uploaded_file.name)[1].lower()
    output_options = conversion_options.get(input_ext, [])

    if not output_options:
        st.warning("No supported conversions for this file type.")
    else:
        output_format = st.selectbox("Convert to", output_options)

        input_path = os.path.join(TEMP_DIR, uploaded_file.name)
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success("File uploaded.")

        base_name, _ = os.path.splitext(uploaded_file.name)
        output_file_name = base_name + output_format
        output_path = os.path.join(TEMP_DIR, output_file_name)

        if st.button("Convert"):
            success = convert_file(input_path, output_path)
            if success and os.path.exists(output_path):
                st.success(f"Conversion successful! Download below.")
                mime_types = {
                    '.pdf': 'application/pdf',
                    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    '.png': 'image/png',
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
                st.error("Conversion failed or unsupported conversion.")
