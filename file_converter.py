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


# Convert PowerPoint file to PDF
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


# Convert file to PDF based on file extension
def convert_file_to_pdf(input_file, output_file):
    file_extension = os.path.splitext(input_file)[1].lower()

    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return

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
        print(f"Unsupported file type: {file_extension}")

# Example usage
if __name__ == "__main__":
    input_file = r"sample_file.pptx"
    output_file = r"converted_file.pdf"

    convert_file_to_pdf(input_file, output_file)
    print(f"Conversion complete. Output saved as {output_file}")
