import os
from fpdf import FPDF
from PIL import Image
from docx import Document
import openpyxl

try:
    import comtypes.client
    POWERPOINT_INSTALLED = True
except ImportError:
    POWERPOINT_INSTALLED = False


# Convert Text file to PDF
def convert_txt_to_pdf(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        text = file.read()
    
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.multi_cell(0, 10, text)
    pdf.output(output_file)
    print(f"Converted TXT to PDF: {output_file}")
