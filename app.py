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

def main():
    st.title("Universal File Converter")
    
    uploaded_file = st.file_uploader("Upload a file", type=list(SUPPORTED_FORMATS.keys()))
    
    if uploaded_file:
        file_name = uploaded_file.name
        input_ext = os.path.splitext(file_name)[1].lower().lstrip('.')
        base_name = os.path.splitext(file_name)[0]
        output_formats = SUPPORTED_FORMATS.get(input_ext, [])
        
        if not output_formats:
            st.error("This file type is not supported")
            return
            
        output_format = st.selectbox("Convert to:", options=output_formats)
        
        # File paths with unique IDs
        file_id = uuid.uuid4().hex
        temp_dir = tempfile.gettempdir()
        input_path = os.path.join(temp_dir, f"input_{file_id}{os.path.splitext(file_name)[1]}")
        output_path = os.path.join(temp_dir, f"output_{file_id}.{output_format}")
        
        # Save uploaded file
        try:
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.info(f"📁 Temporary file saved: {input_path}")
        except Exception as e:
            st.error(f"❌ Failed to save file: {str(e)}")
            return
        
        if st.button("Convert Now"):
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
                
                success = convert_file(input_path, output_path)
                
                if success and os.path.exists(output_path):
                    # Verify ZIP contents
                    if output_format == 'zip':
                        with zipfile.ZipFile(output_path, 'r') as zf:
                            if len(zf.namelist()) == 0:
                                raise ValueError("ZIP file is empty")
                    
                    # Prepare download
                    st.success("Conversion successful! Download your file:")
                    mime_type = MIME_TYPES.get(output_format, 'application/octet-stream')
                    with open(output_path, 'rb') as f:
                        st.download_button(
                            label=f"Download {os.path.basename(output_path)}",
                            data=f,
                            file_name=os.path.basename(output_path),
                            mime=mime_type
                        )
                else:
                    st.error("Conversion failed - no output file created")
                    
            except Exception as e:
                st.error(f"""
                ❌ Conversion failed!
                **Reason:** {str(e)}
                **Technical Details:**  
                ```
                {traceback.format_exc()}
                ```
                """)
            finally:
                # Cleanup
                for path in [input_path, output_path]:
                    if path and os.path.exists(path):
                        try:
                            os.remove(path)
                        except:
                            pass

if __name__ == "__main__":
    main()
