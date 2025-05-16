import streamlit as st
import os
from file_converter import convert_file

TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

st.title("Professional Document Converter")

uploaded_file = st.file_uploader("Upload your file", type=[
    'txt', 'docx', 'pptx', 'xlsx', 'jpg', 'jpeg', 'png'
])

if uploaded_file:
    input_ext = os.path.splitext(uploaded_file.name)[1].lower()
    
    conversion_options = {
        '.docx': ['PDF', 'PNG', 'Excel'],
        '.pptx': ['PDF', 'PNG'],
        '.xlsx': ['PDF', 'PNG'],
        '.txt': ['PDF'],
        '.jpg': ['PDF', 'PNG'],
        '.jpeg': ['PDF', 'PNG'],
        '.png': ['PDF', 'JPG']
    }
    
    ext_map = {
        'PDF': '.pdf',
        'PNG': '.png',
        'Excel': '.xlsx',
        'JPG': '.jpg'
    }
    
    if input_ext not in conversion_options:
        st.error("Unsupported file type")
    else:
        output_format = st.selectbox("Convert to:", conversion_options[input_ext])
        output_ext = ext_map[output_format]
        
        input_path = os.path.join(TEMP_DIR, uploaded_file.name)
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        base_name = os.path.splitext(uploaded_file.name)[0]
        output_path = os.path.join(TEMP_DIR, f"{base_name}_converted{output_ext}")
        
        if st.button("Convert Now"):
            if convert_file(input_path, output_path):
                st.success("Conversion successful! Download your file:")
                with open(output_path, "rb") as f:
                    st.download_button(
                        "Download File",
                        f,
                        file_name=os.path.basename(output_path)
                    )
            else:
                st.error("Conversion failed. Please check the file format.")
