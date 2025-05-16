import streamlit as st
import tempfile
import uuid
import zipfile
import traceback
import os
from io import BytesIO
from file_converter import convert_file

SUPPORTED_FORMATS = {
    'txt': ['pdf', 'docx'],
    'docx': ['txt', 'pdf', 'xlsx'],
    'pptx': ['pdf', 'zip'],
    'xlsx': ['pdf', 'docx', 'txt'],
    'jpg': ['pdf', 'png'],
    'jpeg': ['pdf', 'png'],
    'png': ['txt', 'pdf', 'jpg']
}

MIME_TYPES = {
    'pdf': 'application/pdf',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'txt': 'text/plain',
    'zip': 'application/zip'
}

st.title("Universal File Converter")

def main():
    uploaded_file = st.file_uploader("Upload a file", type=list(SUPPORTED_FORMATS.keys()))
    
    if uploaded_file:
        input_ext = uploaded_file.name.split('.')[-1].lower()
        base_name = '.'.join(uploaded_file.name.split('.')[:-1])
        output_formats = SUPPORTED_FORMATS.get(input_ext, [])
        
        if not output_formats:
            st.error("This file type is not supported")
            return
            
        output_format = st.selectbox("Convert to:", options=output_formats)
        
        if st.button("Convert Now"):
            try:
                with st.spinner("Converting..."):
                    # Use in-memory file for upload
                    file_id = uuid.uuid4().hex
                    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{input_ext}") as input_tmp:
                        input_tmp.write(uploaded_file.getvalue())
                        input_path = input_tmp.name

                    output_path = os.path.join(tempfile.gettempdir(), f"output_{file_id}.{output_format}")

                    # Convert
                    success = convert_file(input_path, output_path)
                    
                    if success and os.path.exists(output_path):
                        with open(output_path, 'rb') as f:
                            file_bytes = f.read()
                        st.success("✅ Conversion successful!")
                        st.download_button(
                            label=f"Download {base_name}.{output_format}",
                            data=file_bytes,
                            file_name=f"{base_name}.{output_format}",
                            mime=MIME_TYPES.get(output_format, 'application/octet-stream')
                        )
                    else:
                        st.error("❌ Conversion failed. Please check the file format.")
            except Exception as e:
                st.error(f"❌ Conversion failed!\n\n**Reason:** {str(e)}\n\n**Technical Details:**\n``````")
            finally:
                # Cleanup temp files
                try:
                    if 'input_path' in locals() and os.path.exists(input_path):
                        os.remove(input_path)
                    if 'output_path' in locals() and os.path.exists(output_path):
                        os.remove(output_path)
                except Exception:
                    pass

if __name__ == "__main__":
    main()
