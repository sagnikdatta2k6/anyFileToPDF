import os
import zipfile
import tempfile
import pythoncom
import win32com.client
import pytesseract
import textwrap
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Set non-GUI backend
import matplotlib.pyplot as plt
from io import BytesIO
from fpdf import FPDF
from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageFont
from docx import Document
import openpyxl
import shutil
import traceback

# ========================
# CONFIGURATION (UPDATE THESE PATHS)
# ========================
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
POPPLER_PATH = r'C:\Program Files\poppler-23.11.0\Library\bin'
WINDOWS_FONT_PATH = r'C:\Windows\Fonts\cour.ttf'

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# ========================
# TEXT CONVERSIONS (DEBUGGED)
# ========================
def convert_txt_to_pdf(input_file, output_file):
    print(f"\n[TXT→PDF] Converting {input_file}")
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            text = file.read()
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, text)
        pdf.output(output_file)
        print(f"[TXT→PDF] Success: {output_file}")
        return True
    except Exception as e:
        print(f"[TXT→PDF] Error: {str(e)}")
        traceback.print_exc()
        return False

def convert_txt_to_image(input_file, output_file):
    print(f"\n[TXT→PNG] Converting {input_file}")
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            text = file.read()

        # Font configuration
        if os.path.exists(WINDOWS_FONT_PATH):
            font = ImageFont.truetype(WINDOWS_FONT_PATH, 14)
        else:
            font = ImageFont.load_default()
            print("[TXT→PNG] Using fallback font")

        # Text wrapping
        lines = []
        for paragraph in text.split('\n'):
            lines.extend(textwrap.wrap(paragraph, width=80, replace_whitespace=False))
        
        # Image dimensions
        line_height = int(font.size * 1.2)
        img_width = 800  # Fixed width for better readability
        img_height = line_height * len(lines) + 40
        
        image = Image.new('RGB', (img_width, img_height), color='white')
        draw = ImageDraw.Draw(image)
        
        # Draw text
        y = 20
        for line in lines:
            draw.text((20, y), line, font=font, fill='black')
            y += line_height
        
        image.save(output_file)
        print(f"[TXT→PNG] Saved: {output_file}")
        return True
    except Exception as e:
        print(f"[TXT→PNG] Error: {str(e)}")
        traceback.print_exc()
        return False

# ========================
# DOCX CONVERSIONS (DEBUGGED)
# ========================
def convert_docx_to_image(input_file, output_file):
    print(f"\n[DOCX→PNG] Converting {input_file}")
    temp_pdf = None
    try:
        # Step 1: Convert DOCX to PDF
        temp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        convert_docx_to_pdf(input_file, temp_pdf.name)
        print(f"[DOCX→PNG] Intermediate PDF: {temp_pdf.name}")

        # Step 2: Convert PDF to PNG
        images = convert_from_path(temp_pdf.name, poppler_path=POPPLER_PATH)
        if images:
            images[0].save(output_file, 'PNG')
            print(f"[DOCX→PNG] Success: {output_file}")
            return True
        print("[DOCX→PNG] No images generated")
        return False
    except Exception as e:
        print(f"[DOCX→PNG] Error: {str(e)}")
        traceback.print_exc()
        return False
    finally:
        if temp_pdf and os.path.exists(temp_pdf.name):
            os.remove(temp_pdf.name)

# ========================
# EXCEL CONVERSIONS (DEBUGGED)
# ========================
def convert_excel_to_image(input_file, output_file):
    print(f"\n[XLSX→PNG] Converting {input_file}")
    try:
        df = pd.read_excel(input_file)
        plt.figure(figsize=(12, min(4 + len(df)*0.3, 20)))
        ax = plt.gca()
        ax.axis('off')
        table = ax.table(
            cellText=df.values,
            colLabels=df.columns,
            cellLoc='center',
            loc='center'
        )
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        plt.savefig(output_file, bbox_inches='tight', dpi=150)
        plt.close()
        print(f"[XLSX→PNG] Saved: {output_file}")
        return True
    except Exception as e:
        print(f"[XLSX→PNG] Error: {str(e)}")
        traceback.print_exc()
        return False

# ========================
# IMAGE CONVERSIONS (DEBUGGED)
# ========================
def convert_image_to_pdf(input_file, output_file):
    print(f"\n[IMAGE→PDF] Converting {input_file}")
    try:
        image = Image.open(input_file)
        if image.mode != 'RGB':
            image = image.convert('RGB')
        pdf = FPDF(unit='pt', format=[image.width, image.height])
        pdf.add_page()
        temp_img = f"{output_file}.temp.jpg"
        image.save(temp_img)
        pdf.image(temp_img, 0, 0, image.width, image.height)
        pdf.output(output_file)
        os.remove(temp_img)
        print(f"[IMAGE→PDF] Success: {output_file}")
        return True
    except Exception as e:
        print(f"[IMAGE→PDF] Error: {str(e)}")
        traceback.print_exc()
        return False

# [Keep other conversion functions from previous versions]

# ========================
# MAIN CONVERSION FUNCTION (DEBUGGED)
# ========================
def convert_file(input_file, output_file):
    print(f"\n🔧 Starting conversion: {input_file} → {output_file}")
    input_ext = os.path.splitext(input_file)[1].lower()
    output_ext = os.path.splitext(output_file)[1].lower()

    conversion_map = {
        # ... [keep existing conversion mappings] ...
    }

    try:
        if input_ext == output_ext:
            raise ValueError("Input and output formats are the same")

        func = conversion_map.get((input_ext, output_ext))
        if not func:
            raise ValueError(f"No converter for {input_ext}→{output_ext}")

        print(f"Using converter: {func.__name__}")
        func(input_file, output_file)
        
        if not os.path.exists(output_file):
            raise FileNotFoundError(f"Output file {output_file} not created")

        print("✅ Conversion successful")
        return True

    except Exception as e:
        print(f"❌ Critical error: {str(e)}")
        traceback.print_exc()
        if os.path.exists(output_file):
            os.remove(output_file)
        raise RuntimeError(f"Conversion failed: {str(e)}")
