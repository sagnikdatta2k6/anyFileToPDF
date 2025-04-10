from fpdf import FPDF
from PIL import Image
import os
from docx import Document
from docx.shared import Inches
import openpyxl

try:
    import comtypes.client
    POWERPOINT_INSTALLED = True
except ImportError:
    POWERPOINT_INSTALLED = False

import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO


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


def convert_image_to_pdf(input_file, output_file):
    image = Image.open(input_file)
    pdf = image.convert("RGB")
    pdf.save(output_file)
    print(f"Converted Image to PDF: {output_file}")


def convert_docx_to_pdf(input_file, output_file):
    doc = Document(input_file)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for para in doc.paragraphs:
        pdf.multi_cell(0, 10, para.text)
    pdf.output(output_file)
    print(f"Converted DOCX to PDF: {output_file}")


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


def convert_pptx_to_pdf(input_file, output_file, format_type=32):
    if not POWERPOINT_INSTALLED:
        print("Error: comtypes module not installed. PowerPoint conversion will not work.")
        return
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        input_file = os.path.abspath(input_file)
        output_file = os.path.abspath(output_file)
        presentation = powerpoint.Presentations.Open(input_file, WithWindow=False)
        presentation.SaveAs(output_file, format_type)
        presentation.Close()
        powerpoint.Quit()
        print(f"Converted PowerPoint to PDF: {output_file}")
    except Exception as e:
        print(f"Error converting PPTX to PDF: {e}")
        try:
            powerpoint.Quit()
        except:
            pass


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
    print(f"Converted PDF to DOCX with images: {output_file}")


def convert_pdf_to_pptx(input_file, output_file):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    pdf = fitz.open(input_file)
    for page in pdf:
        pix = page.get_pixmap(dpi=150)
        image = Image.open(BytesIO(pix.tobytes("png")))
        img_io = BytesIO()
        image.save(img_io, format="PNG")
        img_io.seek(0)
        slide = prs.slides.add_slide(blank_slide_layout)
        slide.shapes.add_picture(img_io, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
    prs.save(output_file)
    print(f"Converted PDF to PPTX with images: {output_file}")


def convert_file(input_file, output_file):
    file_extension = os.path.splitext(input_file)[1].lower()
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return
    if output_file.lower().endswith('.pdf'):
        if file_extension == '.txt':
            convert_txt_to_pdf(input_file, output_file)
        elif file_extension in ['.jpg', '.jpeg', '.png']:
            convert_image_to_pdf(input_file, output_file)
        elif file_extension == '.docx':
            convert_docx_to_pdf(input_file, output_file)
        elif file_extension == '.xlsx':
            convert_excel_to_pdf(input_file, output_file)
        elif file_extension == '.pptx':
            convert_pptx_to_pdf(input_file, output_file)
        else:
            print(f"Unsupported input file type for PDF conversion: {file_extension}")
    elif file_extension == '.pdf':
        if output_file.lower().endswith('.docx'):
            convert_pdf_to_docx(input_file, output_file)
        elif output_file.lower().endswith('.pptx'):
            convert_pdf_to_pptx(input_file, output_file)
        else:
            print(f"Unsupported output format for PDF source: {output_file}")
    else:
        print(f"Unsupported conversion from {file_extension} to {os.path.splitext(output_file)[1].lower()}")


# Example usage
input_file = r"Sagnik.pdf"
output_file = r"sagnik_converted.docx"

convert_file(input_file, output_file)
print(f"Conversion complete. Output saved as {output_file}")
