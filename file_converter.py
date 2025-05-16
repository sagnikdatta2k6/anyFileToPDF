import os
from PIL import Image
import aspose.words as aw
import aspose.slides as slides
import aspose.cells as ac
from io import BytesIO

# DOCX conversions
def convert_docx_to_png(input_file, output_file):
    doc = aw.Document(input_file)
    options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
    options.resolution = 300
    options.page_set = aw.saving.PageSet(0)
    doc.save(output_file, options)

def convert_docx_to_pdf(input_file, output_file):
    doc = aw.Document(input_file)
    doc.save(output_file)

def convert_docx_to_excel(input_file, output_file):
    doc = aw.Document(input_file)
    doc.save(output_file, aw.SaveFormat.XLSX)

# PPTX conversions
def convert_pptx_to_pdf(input_file, output_file):
    pres = slides.Presentation(input_file)
    pres.save(output_file, slides.export.SaveFormat.PDF)

def convert_pptx_to_png(input_file, output_file):
    pres = slides.Presentation(input_file)
    for idx in range(pres.slides.length):
        slide = pres.slides[idx]
        slide.get_thumbnail().save(f"{output_file}_slide{idx+1}.png", slides.export.ImageFormat.PNG)

# Excel conversions
def convert_excel_to_png(input_file, output_file):
    workbook = ac.Workbook(input_file)
    sheet = workbook.worksheets[0]
    img_opts = ac.rendering.ImageOrPrintOptions()
    img_opts.image_format = ac.drawing.ImageFormat.png
    img_opts.horizontal_resolution = 300
    img_opts.vertical_resolution = 300
    renderer = ac.rendering.SheetRender(sheet, img_opts)
    renderer.to_image(0).save(output_file)

def convert_excel_to_pdf(input_file, output_file):
    workbook = ac.Workbook(input_file)
    workbook.save(output_file, ac.SaveFormat.PDF)

# Image conversions
def convert_image_to_pdf(input_file, output_file):
    img = Image.open(input_file)
    img.save(output_file, "PDF", resolution=100.0)

def convert_image_to_image(input_file, output_file):
    img = Image.open(input_file)
    img.save(output_file)

# Text conversions
def convert_txt_to_pdf(input_file, output_file):
    doc = aw.Document()
    builder = aw.DocumentBuilder(doc)
    with open(input_file, "r") as f:
        builder.write(f.read())
    doc.save(output_file)

def convert_file(input_file, output_file):
    input_ext = os.path.splitext(input_file)[1].lower()
    output_ext = os.path.splitext(output_file)[1].lower()

    conversion_map = {
        # Document conversions
        ('.docx', '.png'): convert_docx_to_png,
        ('.docx', '.pdf'): convert_docx_to_pdf,
        ('.docx', '.xlsx'): convert_docx_to_excel,
        
        # Presentation conversions
        ('.pptx', '.pdf'): convert_pptx_to_pdf,
        ('.pptx', '.png'): convert_pptx_to_png,
        
        # Spreadsheet conversions
        ('.xlsx', '.png'): convert_excel_to_png,
        ('.xlsx', '.pdf'): convert_excel_to_pdf,
        
        # Image conversions
        ('.jpg', '.pdf'): convert_image_to_pdf,
        ('.jpeg', '.pdf'): convert_image_to_pdf,
        ('.png', '.pdf'): convert_image_to_pdf,
        ('.jpg', '.png'): convert_image_to_image,
        ('.png', '.jpg'): convert_image_to_image,
        
        # Text conversions
        ('.txt', '.pdf'): convert_txt_to_pdf
    }

    try:
        if input_ext == output_ext:
            return False
        
        if (input_ext, output_ext) not in conversion_map:
            raise ValueError(f"Unsupported conversion: {input_ext} to {output_ext}")
        
        conversion_map[(input_ext, output_ext)](input_file, output_file)
        return True
        
    except Exception as e:
        print(f"Conversion error: {str(e)}")
        return False
