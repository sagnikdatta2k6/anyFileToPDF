import streamlit as st
import os
import tempfile
import uuid
import zipfile
from io import BytesIO
from file_converter import convert_file

SUPPORTED_FORMATS = {
    'txt': ['pdf', 'docx', 'png'],
    'docx': ['txt', 'pdf', 'png', 'xlsx'],
    'pptx': ['pdf', 'zip'],
    'xlsx': ['pdf', 'docx', 'txt', 'png'],
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
        file_name = uploaded_file.name
        input_ext = os.path.splitext(file_name)[1].lower().lstrip('.')
        base_name = os.path.splitext(file_name)[0]
        output_formats = SUPPORTED_FORMATS.get(input_ext, [])
        if not output_formats:
            st.error("Unsupported file type")
            return
            
        output_format = st.selectbox("Convert to:", options=output_formats)
        
        file_id = str(uuid.uuid4())
        temp_dir = tempfile.gettempdir()
        input_path = os.path.join(temp_dir, f"input_{file_id}{os.path.splitext(file_name)[1]}")
        output_file_name = f"{base_name}_converted.{output_format}"
        output_path = os.path.join(temp_dir, f"output_{file_id}.{output_format}")

        try:
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        except PermissionError:
            st.error("❌ Permission denied while saving the uploaded file. Please close any open files in the temp directory or run as administrator.")
            return
        
        if st.button("Convert File"):
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
                success = convert_file(input_path, output_path)
                
                if success:
                    # For ZIP, verify contents before allowing download
                    if output_format == 'zip':
                        with open(output_path, 'rb') as f:
                            zip_bytes = f.read()
                            with zipfile.ZipFile(BytesIO(zip_bytes)) as zf:
                                if len(zf.namelist()) == 0:
                                    raise ValueError("Empty ZIP file generated")
                    
                    st.success("✅ Conversion successful!")
                    mime_type = MIME_TYPES.get(output_format, 'application/octet-stream')
                    with open(output_path, 'rb') as f:
                        st.download_button(
                            label=f"Download {output_file_name}",
                            data=f,
                            file_name=output_file_name,
                            mime=mime_type
                        )
                else:
                    st.error("❌ Conversion failed. Please check the file format.")
                
            except Exception as e:
                st.error(f"❌ Conversion failed: {str(e)}")
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                    except Exception:
                        pass
            finally:
                if os.path.exists(input_path):
                    try:
                        os.remove(input_path)
                    except Exception:
                        pass

if __name__ == "__main__":
    main()
