import os
import zipfile
import tempfile
import pythoncom
import win32com.client
import pytesseract
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from io import BytesIO
from fpdf import FPDF
from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageFont
from docx import Document
import openpyxl
import shutil
import traceback

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
POPPLER_PATH = r'C:\Program Files\poppler-23.11.0\Library\bin'  # Update if needed

def convert_txt_to_pdf(input_file, output_file):
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            text = file.read()
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, text)
        pdf.output(output_file)
        return True
    except Exception as e:
        print(f"[TXT→PDF] Error: {e}")
        traceback.print_exc()
        return False

def convert_txt_to_docx(input_file, output_file):
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            text = file.read()
        doc = Document()
        doc.add_paragraph(text)
        doc.save(output_file)
        return True
    except Exception as e:
        print(f"[TXT→DOCX] Error: {e}")
        traceback.print_exc()
        return False

def convert_docx_to_txt(input_file, output_file):
    try:
        doc = Document(input_file)
        with open(output_file, 'w', encoding='utf-8') as f:
            for para in doc.paragraphs:
                f.write(para.text + '\n')
        return True
    except Exception as e:
        print(f"[DOCX→TXT] Error: {e}")
        traceback.print_exc()
        return False

def convert_docx_to_pdf(input_file, output_file):
    try:
        doc = Document(input_file)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for para in doc.paragraphs:
            pdf.multi_cell(0, 10, para.text)
        pdf.output(output_file)
        return True
    except Exception as e:
        print(f"[DOCX→PDF] Error: {e}")
        traceback.print_exc()
        return False

def convert_docx_to_excel(input_file, output_file):
    try:
        doc = Document(input_file)
        wb = openpyxl.Workbook()
        ws = wb.active
        for i, para in enumerate(doc.paragraphs, start=1):
            ws.cell(row=i, column=1, value=para.text)
        wb.save(output_file)
        return True
    except Exception as e:
        print(f"[DOCX→XLSX] Error: {e}")
        traceback.print_exc()
        return False

def convert_pptx_to_pdf(input_file, output_file):
    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(os.path.abspath(input_file), WithWindow=False)
        presentation.SaveAs(os.path.abspath(output_file), 32)
        presentation.Close()
        powerpoint.Quit()
        pythoncom.CoUninitialize()
        return True
    except Exception as e:
        print(f"[PPTX→PDF] Error: {e}")
        traceback.print_exc()
        return False

def convert_pptx_to_zip(input_file, output_file):
    presentation = None
    powerpoint = None
    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(os.path.abspath(input_file), WithWindow=False)
        with tempfile.TemporaryDirectory() as temp_dir:
            for i in range(1, presentation.Slides.Count + 1):
                slide = presentation.Slides(i)
                slide.Export(os.path.join(temp_dir, f"slide_{i}.png"), "PNG")
            if not os.listdir(temp_dir):
                raise RuntimeError("No slides were converted to PNG")
            temp_zip = output_file + ".tmp"
            with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for img_name in os.listdir(temp_dir):
                    img_path = os.path.join(temp_dir, img_name)
                    zip_file.write(img_path, img_name)
            if os.path.exists(output_file):
                os.remove(output_file)
            os.rename(temp_zip, output_file)
        return True
    except Exception as e:
        if 'temp_zip' in locals() and os.path.exists(temp_zip):
            os.remove(temp_zip)
        print(f"[PPTX→ZIP] Error: {e}")
        traceback.print_exc()
        return False
    finally:
        try:
            if presentation:
                presentation.Close()
            if powerpoint:
                powerpoint.Quit()
            pythoncom.CoUninitialize()
        except:
            pass

def convert_excel_to_pdf(input_file, output_file):
    try:
        wb = openpyxl.load_workbook(input_file)
        sheet = wb.active
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for row in sheet.iter_rows(values_only=True):
            line = "\t".join(str(cell) for cell in row if cell is not None)
            pdf.multi_cell(0, 10, line)
        pdf.output(output_file)
        return True
    except Exception as e:
        print(f"[XLSX→PDF] Error: {e}")
        traceback.print_exc()
        return False

def convert_excel_to_docx(input_file, output_file):
    try:
        wb = openpyxl.load_workbook(input_file)
        sheet = wb.active
        doc = Document()
        for row in sheet.iter_rows(values_only=True):
            line = "\t".join(str(cell) for cell in row if cell is not None)
            doc.add_paragraph(line)
        doc.save(output_file)
        return True
    except Exception as e:
        print(f"[XLSX→DOCX] Error: {e}")
        traceback.print_exc()
        return False

def convert_excel_to_txt(input_file, output_file):
    try:
        wb = openpyxl.load_workbook(input_file)
        sheet = wb.active
        with open(output_file, 'w', encoding='utf-8') as f:
            for row in sheet.iter_rows(values_only=True):
                line = "\t".join(str(cell) for cell in row if cell is not None)
                f.write(line + '\n')
        return True
    except Exception as e:
        print(f"[XLSX→TXT] Error: {e}")
        traceback.print_exc()
        return False

def convert_png_to_txt(input_file, output_file):
    try:
        img = Image.open(input_file)
        text = pytesseract.image_to_string(img)
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(text)
        return True
    except Exception as e:
        print(f"OCR Error: {e}")
        traceback.print_exc()
        return False

def convert_image_to_pdf(input_file, output_file):
    try:
        image = Image.open(input_file)
        if image.mode != 'RGB':
            image = image.convert('RGB')
        pdf = FPDF(unit='pt', format=[image.width, image.height])
        pdf.add_page()
        temp_img = output_file + ".temp.jpg"
        image.save(temp_img)
        pdf.image(temp_img, 0, 0, image.width, image.height)
        pdf.output(output_file)
        os.remove(temp_img)
        return True
    except Exception as e:
        print(f"Image to PDF Error: {e}")
        traceback.print_exc()
        return False

def convert_png_to_jpg(input_file, output_file):
    try:
        image = Image.open(input_file)
        if image.mode != 'RGB':
            image = image.convert('RGB')
        image.save(output_file, 'JPEG', quality=95)
        return True
    except Exception as e:
        print(f"PNG to JPG Error: {e}")
        traceback.print_exc()
        return False

def convert_jpg_to_png(input_file, output_file):
    try:
        image = Image.open(input_file)
        image.save(output_file, 'PNG')
        return True
    except Exception as e:
        print(f"JPG to PNG Error: {e}")
        traceback.print_exc()
        return False

def convert_file(input_file, output_file):
    input_ext = os.path.splitext(input_file)[1].lower()
    output_ext = os.path.splitext(output_file)[1].lower()

    conversion_map = {
        ('.txt', '.pdf'): convert_txt_to_pdf,
        ('.txt', '.docx'): convert_txt_to_docx,
        ('.docx', '.txt'): convert_docx_to_txt,
        ('.docx', '.pdf'): convert_docx_to_pdf,
        ('.docx', '.xlsx'): convert_docx_to_excel,
        ('.pptx', '.pdf'): convert_pptx_to_pdf,
        ('.pptx', '.zip'): convert_pptx_to_zip,
        ('.xlsx', '.pdf'): convert_excel_to_pdf,
        ('.xlsx', '.docx'): convert_excel_to_docx,
        ('.xlsx', '.txt'): convert_excel_to_txt,
        ('.png', '.txt'): convert_png_to_txt,
        ('.png', '.pdf'): convert_image_to_pdf,
        ('.png', '.jpg'): convert_png_to_jpg,
        ('.jpg', '.pdf'): convert_image_to_pdf,
        ('.jpg', '.png'): convert_jpg_to_png,
    }

    if input_ext == output_ext:
        raise ValueError("Input and output formats cannot be the same")

    func = conversion_map.get((input_ext, output_ext))
    if not func:
        raise ValueError(f"No converter for {input_ext}→{output_ext}")

    try:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        result = func(input_file, output_file)
        if not os.path.exists(output_file):
            raise RuntimeError("Output file was not created")
        if output_ext == '.zip':
            with zipfile.ZipFile(output_file, 'r') as zip_ref:
                if len(zip_ref.namelist()) == 0:
                    os.remove(output_file)
                    raise RuntimeError("Converted ZIP file is empty")
        return True
    except Exception as e:
        print(f"[convert_file] Error: {e}")
        traceback.print_exc()
        if os.path.exists(output_file):
            if os.path.isdir(output_file):
                shutil.rmtree(output_file)
            else:
                os.remove(output_file)
        raise RuntimeError(f"Conversion failed: {e}")
