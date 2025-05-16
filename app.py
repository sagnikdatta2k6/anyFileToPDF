import streamlit as st
import os
from file_converter import convert_file

# Configuration
TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

# Supported formats
SUPPORTED_FORMATS = {
    'txt': ['pdf', 'docx', 'png'],
    'docx': ['txt', 'pdf', 'png', 'xlsx'],
    'pptx': ['pdf', 'zip'],
    'xlsx': ['pdf', 'docx', 'txt', 'png'],
    'jpg': ['pdf', 'png'],
    'jpeg': ['pdf', 'png'],
    'png': ['pdf', 'jpg']
}

# MIME types for download
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
    uploaded_file = st.file_uploader("Upload a file", type=SUPPORTED_FORMATS.keys())
    
    if uploaded_file:
        # Get file info
        file_name = uploaded_file.name
        input_ext = os.path.splitext(file_name)[1].lower().lstrip('.')
        base_name = os.path.splitext(file_name)[0]
        
        # Show output format selector
        output_formats = SUPPORTED_FORMATS.get(input_ext, [])
        if not output_formats:
            st.error("Unsupported file type")
            return
            
        output_format = st.selectbox("Convert to:", options=output_formats)
        
        # Prepare file paths
        input_path = os.path.join(TEMP_DIR, file_name)
        output_file_name = f"{base_name}_converted.{output_format}"
        output_path = os.path.join(TEMP_DIR, output_file_name)
        
        # Save uploaded file
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Conversion button
        if st.button("Convert File"):
            try:
                # Clear previous output
                if os.path.exists(output_path):
                    os.remove(output_path)
                
                # Perform conversion
                success = convert_file(input_path, output_path)
                
                if success:
                    # Verify ZIP contents
                    if output_format == 'zip':
                        with open(output_path, 'rb') as f:
                            zip_bytes = f.read()
                            with zipfile.ZipFile(BytesIO(zip_bytes)) as zf:
                                if len(zf.namelist()) == 0:
                                    raise ValueError("Empty ZIP file generated")
                    
                    st.success("✅ Conversion successful!")
                    
                    # Show download button
                    mime_type = MIME_TYPES.get(output_format, 'application/octet-stream')
                    with open(output_path, 'rb') as f:
                        st.download_button(
                            label=f"Download {output_file_name}",
                            data=f,
                            file_name=output_file_name,
                            mime=mime_type
                        )
                
            except Exception as e:
                st.error(f"❌ Conversion failed: {str(e)}")
                # Cleanup failed output
                if os.path.exists(output_path):
                    os.remove(output_path)

if __name__ == "__main__":
    main()
