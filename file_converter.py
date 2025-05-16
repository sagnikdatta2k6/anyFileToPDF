import os
from fpdf import FPDF
from PIL import Image
from docx import Document
import openpyxl
import pdfplumber
import fitz  # PyMuPDF
from pptx import Presentation
from io import BytesIO

try:
    import comtypes.client
    POWERPOINT_INSTALLED = True
except ImportError:
    POWERPOINT_INSTALLED = False

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

def convert_txt_to_docx(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        text = file.read()
    doc = Document()
    doc.add_paragraph(text)
    doc.save(output_file)
    print(f"Converted TXT to DOCX: {output_file}")

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

def convert_docx_to_txt(input_file, output_file):
    doc = Document(input_file)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(full_text))
    print(f"Converted DOCX to TXT: {output_file}")

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

def convert_pptx_to_pdf(input_file, output_file):
    if not POWERPOINT_INSTALLED:
        print("Error: comtypes module not installed. PowerPoint conversion will not work.")
        return False
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return False
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1

        input_file_abs = os.path.abspath(input_file)
        output_file_abs = os.path.abspath(output_file)

        if not output_file_abs.lower().endswith('.pdf'):
            output_file_abs += '.pdf'

        print(f"Opening PowerPoint file: {input_file_abs}")
        print(f"Saving PDF to: {output_file_abs}")

        presentation = powerpoint.Presentations.Open(input_file_abs, WithWindow=False)
        presentation.SaveAs(output_file_abs, 32)  # 32 = PDF format
        presentation.Close()
        powerpoint.Quit()

        if os.path.exists(output_file_abs):
            print(f"Converted PowerPoint to PDF successfully: {output_file_abs}")
            return True
        else:
            print(f"Failed to create PDF at: {output_file_abs}")
            return False

    except Exception as e:
        print(f"Error converting PPTX to PDF: {e}")
        try:
            powerpoint.Quit()
        except:
            pass
        return False

def convert_file(input_file, output_file):
    input_file = os.path.abspath(input_file)
    output_file = os.path.abspath(output_file)

    input_ext = os.path.splitext(input_file)[1].lower()
    output_ext = os.path.splitext(output_file)[1].lower()

    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return False

    # TXT conversions
    if input_ext == '.txt':
        if output_ext == '.pdf':
            convert_txt_to_pdf(input_file, output_file)
            return True
        elif output_ext == '.docx':
            convert_txt_to_docx(input_file, output_file)
            return True
        else:
            print(f"Unsupported output format for TXT: {output_ext}")
            return False

    # DOCX conversions
    elif input_ext == '.docx':
        if output_ext == '.pdf':
            convert_docx_to_pdf(input_file, output_file)
            return True
        elif output_ext == '.txt':
            convert_docx_to_txt(input_file, output_file)
            return True
        else:
            print(f"Unsupported output format for DOCX: {output_ext}")
            return False

    # Image to PDF
    elif input_ext in ['.jpg', '.jpeg', '.png']:
        if output_ext == '.pdf':
            convert_image_to_pdf(input_file, output_file)
            return True
        else:
            print(f"Unsupported output format for image: {output_ext}")
            return False

    # Excel to PDF
    elif input_ext == '.xlsx':
        if output_ext == '.pdf':
            convert_excel_to_pdf(input_file, output_file)
            return True
        else:
            print(f"Unsupported output format for Excel: {output_ext}")
            return False

    # PPTX to PDF
    elif input_ext == '.pptx':
        if output_ext == '.pdf':
            return convert_pptx_to_pdf(input_file, output_file)
        else:
            print(f"Unsupported output format for PPTX: {output_ext}")
            return False

    else:
        print(f"Unsupported input file type: {input_ext}")
        return False
