import os
from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont
from docx import Document
import openpyxl
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO

# TXT to PDF
def convert_txt_to_pdf(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        text = file.read()
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.multi_cell(0, 10, text)
    pdf.output(output_file)

# TXT to DOCX
def convert_txt_to_docx(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        text = file.read()
    doc = Document()
    doc.add_paragraph(text)
    doc.save(output_file)

# TXT to IMAGE (PNG)
def convert_txt_to_image(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        text = file.read()
    font = ImageFont.load_default()
    lines = text.split('\n')
    max_width = max(font.getsize(line)[0] for line in lines) + 20
    line_height = font.getsize('A')[1] + 5
    img_height = line_height * len(lines) + 20
    image = Image.new('RGB', (max_width, img_height), color='white')
    draw = ImageDraw.Draw(image)
    y = 10
    for line in lines:
        draw.text((10, y), line, font=font, fill='black')
        y += line_height
    image.save(output_file)

# DOCX to TXT
def convert_docx_to_txt(input_file, output_file):
    doc = Document(input_file)
    with open(output_file, 'w', encoding='utf-8') as f:
        for para in doc.paragraphs:
            f.write(para.text + '\n')

# DOCX to PDF
def convert_docx_to_pdf(input_file, output_file):
    doc = Document(input_file)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for para in doc.paragraphs:
        pdf.multi_cell(0, 10, para.text)
    pdf.output(output_file)

# DOCX to IMAGE (PNG)
def convert_docx_to_image(input_file, output_file):
    doc = Document(input_file)
    text = '\n'.join([para.text for para in doc.paragraphs])
    font = ImageFont.load_default()
    lines = text.split('\n')
    max_width = max(font.getsize(line)[0] for line in lines) + 20
    line_height = font.getsize('A')[1] + 5
    img_height = line_height * len(lines) + 20
    image = Image.new('RGB', (max_width, img_height), color='white')
    draw = ImageDraw.Draw(image)
    y = 10
    for line in lines:
        draw.text((10, y), line, font=font, fill='black')
        y += line_height
    image.save(output_file)

# DOCX to EXCEL
def convert_docx_to_excel(input_file, output_file):
    doc = Document(input_file)
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, para in enumerate(doc.paragraphs, start=1):
        ws.cell(row=i, column=1, value=para.text)
    wb.save(output_file)

# PPTX to PDF (slides as images in PDF)
def convert_pptx_to_pdf(input_file, output_file):
    prs = Presentation(input_file)
    pdf = FPDF()
    for slide in prs.slides:
        img = slide_to_image(slide, prs)
        img_buffer = BytesIO()
        img.save(img_buffer, format='PNG')
        img_buffer.seek(0)
        pdf.add_page()
        pdf.image(img_buffer, x=10, y=10, w=180)
    pdf.output(output_file)

# PPTX to IMAGE (first slide as PNG)
def convert_pptx_to_image(input_file, output_file):
    prs = Presentation(input_file)
    img = slide_to_image(prs.slides[0], prs)
    img.save(output_file)

def slide_to_image(slide, prs):
    # Create a blank image with slide dimensions
    width, height = int(prs.slide_width // 9525), int(prs.slide_height // 9525)
    image = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(image)
    # Placeholder: Draw slide title and content text
    y = 10
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            draw.text((10, y), shape.text, fill='black')
            y += 20
    return image

# EXCEL to PDF
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

# EXCEL to DOCX
def convert_excel_to_docx(input_file, output_file):
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active
    doc = Document()
    for row in sheet.iter_rows(values_only=True):
        line = "\t".join(str(cell) for cell in row if cell is not None)
        doc.add_paragraph(line)
    doc.save(output_file)

# EXCEL to TXT
def convert_excel_to_txt(input_file, output_file):
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active
    with open(output_file, 'w', encoding='utf-8') as f:
        for row in sheet.iter_rows(values_only=True):
            line = "\t".join(str(cell) for cell in row if cell is not None)
            f.write(line + '\n')

# EXCEL to IMAGE (PNG)
def convert_excel_to_image(input_file, output_file):
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active
    lines = []
    for row in sheet.iter_rows(values_only=True):
        line = "\t".join(str(cell) for cell in row if cell is not None)
        lines.append(line)
    text = '\n'.join(lines)
    font = ImageFont.load_default()
    lines = text.split('\n')
    max_width = max(font.getsize(line)[0] for line in lines) + 20
    line_height = font.getsize('A')[1] + 5
    img_height = line_height * len(lines) + 20
    image = Image.new('RGB', (max_width, img_height), color='white')
    draw = ImageDraw.Draw(image)
    y = 10
    for line in lines:
        draw.text((10, y), line, font=font, fill='black')
        y += line_height
    image.save(output_file)

def convert_file(input_file, output_file):
    input_file = os.path.abspath(input_file)
    output_file = os.path.abspath(output_file)
    input_ext = os.path.splitext(input_file)[1].lower()
    output_ext = os.path.splitext(output_file)[1].lower()

    conversion_map = {
        ('.txt', '.pdf'): convert_txt_to_pdf,
        ('.txt', '.docx'): convert_txt_to_docx,
        ('.txt', '.png'): convert_txt_to_image,
        ('.docx', '.txt'): convert_docx_to_txt,
        ('.docx', '.pdf'): convert_docx_to_pdf,
        ('.docx', '.png'): convert_docx_to_image,
        ('.docx', '.xlsx'): convert_docx_to_excel,
        ('.pptx', '.pdf'): convert_pptx_to_pdf,
        ('.pptx', '.png'): convert_pptx_to_image,
        ('.xlsx', '.pdf'): convert_excel_to_pdf,
        ('.xlsx', '.docx'): convert_excel_to_docx,
        ('.xlsx', '.txt'): convert_excel_to_txt,
        ('.xlsx', '.png'): convert_excel_to_image,
    }

    func = conversion_map.get((input_ext, output_ext))
    if func:
        try:
            func(input_file, output_file)
            return True
        except Exception as e:
            print(f"Conversion error: {e}")
            return False
    else:
        print(f"Unsupported conversion from {input_ext} to {output_ext}")
        return False
