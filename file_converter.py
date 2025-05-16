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

# TXT to PDF
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

# TXT to DOCX
def convert_txt_to_docx(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        text = file.read()
    doc = Document()
    doc.add_paragraph(text)
    doc.save(output_file)
    print(f"Converted TXT to DOCX: {output_file}")

# Image to PDF
def convert_image_to_pdf(input_file, output_file):
    image = Image.open(input_file)
    pdf = image.convert("RGB")
    pdf.save(output_file)
    print(f"Converted Image to PDF: {output_file}")

# DOCX to PDF
def convert_docx_to_pdf(input_file, output_file):
    doc = Document(input_file)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for para in doc.paragraphs:
        pdf.multi_cell(0, 10, para.text)
    pdf.output(output_file)
    print(f"Converted DOCX to PDF: {output_file}")

# DOCX to TXT
def convert_docx_to_txt(input_file, output_file):
    doc = Document(input_file)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(full_text))
    print(f"Converted DOCX to TXT: {output_file}")

# Excel to PDF
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

# PPTX to PDF (Windows only)
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

# PDF to DOCX (images)
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

# PDF to PPTX (images)
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

# PDF to Excel (extract tables)
def convert_pdf_to_excel(input_file, output_file):
    with pdfplumber.open(input_file) as pdf:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Extracted Tables"
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    ws.append(row)
        wb.save(output_file)
        print(f"Converted PDF tables to Excel: {output_file}")

# Main conversion dispatcher
def convert_file(input_file, output_file):
    input_file = os.path.abspath(input_file)
    output_file = os.path.abspath(output_file)

    input_ext = os.path.splitext(input_file)[1].lower()
    output_ext = os.path.splitext(output_file)[1].lower()

    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return False

    # Conversions TO PDF
    if output_ext == '.pdf':
        if input_ext == '.txt':
            convert_txt_to_pdf(input_file, output_file)
            return True
        elif input_ext in ['.jpg', '.jpeg', '.png']:
            convert_image_to_pdf(input_file, output_file)
            return True
        elif input_ext == '.docx':
            convert_docx_to_pdf(input_file, output_file)
            return True
        elif input_ext == '.xlsx':
            convert_excel_to_pdf(input_file, output_file)
            return True
        elif input_ext == '.pptx':
            return convert_pptx_to_pdf(input_file, output_file)
        else:
            print(f"Unsupported input file type for PDF conversion: {input_ext}")
            return False

    # Conversions FROM PDF
    elif input_ext == '.pdf':
        if output_ext == '.docx':
            convert_pdf_to_docx(input_file, output_file)
            return True
        elif output_ext == '.pptx':
            convert_pdf_to_pptx(input_file, output_file)
            return True
        elif output_ext == '.xlsx':
            convert_pdf_to_excel(input_file, output_file)
            return True
        else:
            print(f"Unsupported output format for PDF input: {output_ext}")
            return False

    # TXT conversions
    elif input_ext == '.txt':
        if output_ext == '.docx':
            convert_txt_to_docx(input_file, output_file)
            return True
        else:
            print(f"Unsupported output format for TXT input: {output_ext}")
            return False

    # DOCX conversions
    elif input_ext == '.docx':
        if output_ext == '.txt':
            convert_docx_to_txt(input_file, output_file)
            return True
        else:
            print(f"Unsupported output format for DOCX input: {output_ext}")
            return False

    else:
        print(f"Unsupported conversion from {input_ext} to {output_ext}")
        return False
