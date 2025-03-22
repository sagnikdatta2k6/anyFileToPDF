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


# Convert Image file to PDF
def convert_image_to_pdf(input_file, output_file):
    image = Image.open(input_file)
    pdf = image.convert("RGB")
    pdf.save(output_file)
    print(f"Converted Image to PDF: {output_file}")


# Convert Word file to PDF
def convert_docx_to_pdf(input_file, output_file):
    doc = Document(input_file)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    for para in doc.paragraphs:
        pdf.multi_cell(0, 10, para.text)
    
    pdf.output(output_file)
    print(f"Converted DOCX to PDF: {output_file}")


# Convert Excel file to PDF
def convert_excel_to_pdf(input_file, output_file):
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    for row in sheet.iter_rows(values_only=True):
        line = "\t".join(str(cell) for cell in row if cell is not None)
        pdf.multi_cell(0, 10, line)
    
    pdf.output(output_file)
    print(f"Converted Excel to PDF: {output_file}")


