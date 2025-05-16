import os
from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont
from docx import Document
import openpyxl
from pptx import Presentation
from io import BytesIO

# For PPTX to PDF/PNG using Microsoft PowerPoint (Windows only)
try:
    import win32com.client
    import pythoncom
    POWERPOINT_AUTOMATION = True
except ImportError:
    POWERPOINT_AUTOMATION = False

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

# PPTX to PDF (using PowerPoint automation)
def convert_pptx_to_pdf(input_file, output_file):
    if not POWERPOINT_AUTOMATION:
        raise RuntimeError("pywin32 is not installed or not available on this system.")
    pythoncom.CoInitialize()
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    input_file_abs = os.path.abspath(input_file)
    output_file_abs = os.path.abspath(output_file)
    presentation = powerpoint.Presentations.Open(input_file_abs, WithWindow=False)
    presentation.SaveAs(output_file_abs, 32)  # 32 = PDF format
    presentation.Close()
    powerpoint.Quit()
    pythoncom.CoUninitialize()

# PPTX to PNG (using PowerPoint automation, exports all slides)
def convert_pptx_to_png(input_file, output_file):
    if not POWERPOINT_AUTOMATION:
        raise RuntimeError("pywin32 is not installed or not available on this system.")
    pythoncom.CoInitialize()
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    input_file_abs = os.path.abspath(input_file)
    output_dir = os.path.dirname(output_file)
    base_name = os.path.splitext(os.path.basename(output_file))[0]
    presentation = powerpoint.Presentations.Open(input_file_abs, WithWindow=False)
    # Export all slides as PNGs
    for i in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides(i)
        slide.Export(os.path.join(output_dir, f"{base_name}_slide{i}.png"), "PNG", 1920, 1080)
    presentation.Close()
    powerpoint.Quit()
    pythoncom.CoUninitialize()

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
        ('.pptx', '.png'): convert_pptx_to_png,
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
