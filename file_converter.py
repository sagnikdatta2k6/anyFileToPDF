import os
import zipfile
import tempfile
import pythoncom
import win32com.client
from io import BytesIO
from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont
from docx import Document
import openpyxl
from pptx import Presentation
import shutil

# Text Conversions
def convert_txt_to_pdf(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        text = file.read()
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.multi_cell(0, 10, text)
    pdf.output(output_file)

def convert_txt_to_docx(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        text = file.read()
    doc = Document()
    doc.add_paragraph(text)
    doc.save(output_file)

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

# DOCX Conversions
def convert_docx_to_txt(input_file, output_file):
    doc = Document(input_file)
    with open(output_file, 'w', encoding='utf-8') as f:
        for para in doc.paragraphs:
            f.write(para.text + '\n')

def convert_docx_to_pdf(input_file, output_file):
    doc = Document(input_file)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for para in doc.paragraphs:
        pdf.multi_cell(0, 10, para.text)
    pdf.output(output_file)

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

def convert_docx_to_excel(input_file, output_file):
    doc = Document(input_file)
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, para in enumerate(doc.paragraphs, start=1):
        ws.cell(row=i, column=1, value=para.text)
    wb.save(output_file)

# PPTX Conversions
def convert_pptx_to_pdf(input_file, output_file):
    pythoncom.CoInitialize()
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(os.path.abspath(input_file), WithWindow=False)
    presentation.SaveAs(os.path.abspath(output_file), 32)
    presentation.Close()
    powerpoint.Quit()
    pythoncom.CoUninitialize()

def convert_pptx_to_zip(input_file, output_file):
    """Converts PPTX to ZIP archive containing all slides as PNGs with proper error handling."""
    presentation = None
    powerpoint = None
    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(os.path.abspath(input_file), WithWindow=False)
        
        with tempfile.TemporaryDirectory() as temp_dir:
            # Export slides to temporary directory
            for i in range(1, presentation.Slides.Count + 1):
                slide = presentation.Slides(i)
                slide.Export(os.path.join(temp_dir, f"slide_{i}.png"), "PNG")
            
            # Create ZIP file only if slides were exported
            if not os.listdir(temp_dir):
                raise RuntimeError("No slides were converted to PNG")
            
            # Create ZIP in temporary location first
            temp_zip = output_file + ".tmp"
            with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for img_name in os.listdir(temp_dir):
                    img_path = os.path.join(temp_dir, img_name)
                    zip_file.write(img_path, img_name)
            
            # Replace final file only if successful
            if os.path.exists(output_file):
                os.remove(output_file)
            os.rename(temp_zip, output_file)
        
        return True
        
    except Exception as e:
        # Cleanup failed files
        if 'temp_zip' in locals() and os.path.exists(temp_zip):
            os.remove(temp_zip)
        raise RuntimeError(f"PPTX conversion failed: {str(e)}")
    finally:
        try:
            if presentation:
                presentation.Close()
            if powerpoint:
                powerpoint.Quit()
            pythoncom.CoUninitialize()
        except:
            pass

# Excel Conversions
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

def convert_excel_to_docx(input_file, output_file):
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active
    doc = Document()
    for row in sheet.iter_rows(values_only=True):
        line = "\t".join(str(cell) for cell in row if cell is not None)
        doc.add_paragraph(line)
    doc.save(output_file)

def convert_excel_to_txt(input_file, output_file):
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active
    with open(output_file, 'w', encoding='utf-8') as f:
        for row in sheet.iter_rows(values_only=True):
            line = "\t".join(str(cell) for cell in row if cell is not None)
            f.write(line + '\n')

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

# Main Conversion Function
def convert_file(input_file, output_file):
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
        ('.pptx', '.zip'): convert_pptx_to_zip,
        ('.xlsx', '.pdf'): convert_excel_to_pdf,
        ('.xlsx', '.docx'): convert_excel_to_docx,
        ('.xlsx', '.txt'): convert_excel_to_txt,
        ('.xlsx', '.png'): convert_excel_to_image,
    }

    if input_ext == output_ext:
        raise ValueError("Input and output formats cannot be the same")

    func = conversion_map.get((input_ext, output_ext))
    if not func:
        raise ValueError(f"Unsupported conversion: {input_ext} to {output_ext}")

    try:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        func(input_file, output_file)
        if not os.path.exists(output_file):
            raise RuntimeError("Output file was not created")
        if output_ext == '.zip':
            with zipfile.ZipFile(output_file, 'r') as zip_ref:
                if len(zip_ref.namelist()) == 0:
                    os.remove(output_file)
                    raise RuntimeError("Converted ZIP file is empty")
        return True
    except Exception as e:
        if os.path.exists(output_file):
            if os.path.isdir(output_file):
                shutil.rmtree(output_file)
            else:
                os.remove(output_file)
        raise RuntimeError(f"Conversion failed: {e}")
    
