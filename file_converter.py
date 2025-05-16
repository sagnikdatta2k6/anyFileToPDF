import os
from fpdf import FPDF
from PIL import Image
from docx import Document
import openpyxl
import pdfplumber
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO

try:
    import comtypes.client
    POWERPOINT_INSTALLED = True
except ImportError:
    POWERPOINT_INSTALLED = False

# Conversion functions (examples)
def convert_txt_to_pdf(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as f:
        text = f.read()
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.multi_cell(0, 10, text)
    pdf.output(output_file)

def convert_pdf_to_docx(input_file, output_file):
    doc = Document()
    pdf = fitz.open(input_file)
    for page in pdf:
        pix = page.get_pixmap(dpi=200)
        img = Image.open(BytesIO(pix.tobytes("png")))
        img_io = BytesIO()
        img.save(img_io, format="PNG")
        img_io.seek(0)
        doc.add_picture(img_io, width=Inches(6.5))
        doc.add_paragraph("")
    doc.save(output_file)

# Add other conversion functions here...

# Conversion map: (input_ext, output_ext) -> function
CONVERSION_MAP = {
    ('.txt', '.pdf'): convert_txt_to_pdf,
    ('.pdf', '.docx'): convert_pdf_to_docx,
    # Add all supported conversions here
}

def convert_file(input_file, output_file):
    input_file = os.path.abspath(input_file)
    output_file = os.path.abspath(output_file)
    input_ext = os.path.splitext(input_file)[1].lower()
    output_ext = os.path.splitext(output_file)[1].lower()

    if input_ext == output_ext:
        # No conversion needed, just copy
        import shutil
        shutil.copy(input_file, output_file)
        return True

    func = CONVERSION_MAP.get((input_ext, output_ext))
    if func:
        try:
            func(input_file, output_file)
            return True
        except Exception as e:
            print(f"Conversion error: {e}")
            return False
    else:
        print(f"No conversion function for {input_ext} to {output_ext}")
        return False
